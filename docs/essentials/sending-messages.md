---
title: "Sending Messages"
sidebar_position: "4"
---

# Sending Messages

Sending messages is a core part of an agent's functionality. With all activity handlers, a `send` method is provided which allows your handlers to send a message back to the user to the relevant conversation.

:::info[SDK 2.1]
In SDK 2.1, the per-turn `context.Send`, `context.Reply`, and `context.Quote` helpers work the same way. The main difference is **streaming** — see the tabbed examples below.
:::

```typescript
app.on('message', async ({ activity, send }) => {
  await send(`You said: ${activity.text}`);
});
```

In the above example, the handler gets a `message` activity, and uses the `send` method to send a reply to the user.

```typescript
app.on('signin.verify-state', async ({ send }) => {
  await send('You have successfully signed in!');
});
```
You are not restricted to only replying to `message` activities. In the above example, the handler is listening to `signin.verify-state` events, which are sent when a user successfully signs in.

:::tip
This shows an example of sending a text message. Additionally, you are able to send back things like [adaptive cards](../in-depth-guides/adaptive-cards.md) by using the same `send` method. Look at the [adaptive card](../in-depth-guides/adaptive-cards.md) section for more details.
:::

## Streaming

You may also stream messages to the user which can be useful for long messages, or AI generated messages. The SDK makes this simple for you by providing a `stream` function which you can use to send messages in chunks.

Streaming changes how the response arrives, not how fast your agent produces it. Text appears incrementally as it is generated, so the user sees progress rather than a blank screen, and you can show a thinking or status indicator before any real output exists. A streamed response is also stoppable: when the user stops it, the SDK surfaces the cancellation to your handler so you can abandon the remaining work. The gain is in perceived time to first output, which is why streaming suits long or model generated replies and adds little to a short one you could simply send.

```typescript
app.on('message', async ({ activity, stream }) => {
  stream.emit('hello');
  stream.emit(', ');
  stream.emit('world!');

  // result message: "hello, world!"
});
```

:::note
Streaming is currently only supported in 1:1 conversations, not group chats or channels. Teams also supports only one concurrent streaming response per chat at a time.

Every inbound activity gets its own stream, and the SDK closes that stream at the end of the turn. It does not serialize turns, though, and handlers can run concurrently, so a second activity can arrive while the first is still streaming. Streaming therefore brings message handling concurrency concerns that do not exist when you just send a reply, and resolving them may be implemented in your custom code by serializing turns per conversation or by finalizing the in-flight stream before starting the next one.
:::

![Animated image showing agent response text incrementally appearing in the chat window.](/screenshots/streaming-chat.gif)

### Informative updates and stopping a stream

An informative update shows a status line such as "Thinking..." while the agent works, before any response text exists.

`stream.update()` sends an informative update.
An update emitted after response text has accumulated is ignored.
The SDK learns that the user stopped the stream from the answer to a chunk it sends.
From then on `stream.canceled` is true and `stream.emit()` throws `StreamCancelledError`.
A stream creates its message from response text or an informative update.
If the first output is only attachments, `stream.close()` creates no message, so send the attachments with `send()` instead.

### Chunk pacing and retries

You do not need to throttle your own chunks. Rapid `stream.emit()` calls are coalesced and spaced roughly 500 milliseconds apart, so a token-by-token loop becomes a couple of network calls per second rather than one call per token. Every streamed chunk carries the full text accumulated so far rather than just the newest delta, so collapsing several chunks into one send never loses content. Wrapping your own throttle or sleep around each chunk mostly slows the response down, since it would be pacing a path that is already paced.

`stream.emit()` is fire and forget. It appends to a queue and returns immediately instead of awaiting the network, and a flush then drains the whole queue into one accumulated buffer and sends it as a single typing activity. When further chunks arrive while a send is already in flight, the next flush is scheduled 500ms later, which is what collapses a fast loop into roughly two sends per second. Chunks that arrive slower than a round trip are sent as they come, so the SDK never adds latency to a slow producer.

Failed sends are retried for you with exponential backoff: up to 5 attempts, waiting 500ms, 1s, 2s, and 4s between them. This is a blind retry on transient failures rather than rate limit aware backoff, since the SDK does not read `Retry-After` or treat `429` specially. Terminal `403`s are deliberately excluded, so a stream that has already timed out or been cancelled is never retried back to life.

None of this extends the [two minute window](#the-two-minute-streaming-limit), which runs on wall clock from the first chunk. Pacing keeps a stream under the streaming API's rate limit, but does not buy more time. It also applies only to the streaming path: the plain activity updates go out exactly when they are sent, so that cadence stays yours to manage.

### The two minute streaming limit

Teams keeps a message stream open for two minutes, measured from its first chunk. After that the service rejects further chunks for that message with a `403` and the [error code](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/streaming-ux#error-codes) `ContentStreamNotAllowed` (`Content stream finished due to exceeded streaming time.`).

However, hitting the limit does not mean the response is lost. The SDK marks the stream as timed out and stops pushing chunks, but it keeps buffering everything your agent produces, and `stream.close()` finalizes by updating the original message in place with the full buffered text. 

:::warning
The two minute clock starts on the **first streamed chunk**, and an informative update (`stream.update()`) is a chunk. An agent that posts a status update immediately and then thinks for three minutes has spent its entire window before emitting a single token. When the slow work happens before you have any output, do that work first and open the stream once real content starts.
:::

### Running longer than two minutes

To outlast the window, hand off before it closes: stream normally, then finalize the stream and continue the response growing with plain activity updates. No streaming window applies to plain updates, so the answer can keep arriving for as long as your agent runs.

```typescript

app.on('message', async ({ activity, send, stream }) => {
  let opened: number | undefined;
  let text = '';
  let editing = false;
  let messageId: string | undefined;
  let lastEdit = 0;

  for await (const chunk of runAgent(activity.text)) {
    text += chunk;

    // Phase 1: real streaming. The window opens on the first streamed chunk,
    // not when the handler starts, so the clock starts here.
    if (!editing && (opened === undefined || Date.now() - opened < 110_000)) {
      opened ??= Date.now();
      stream.emit(chunk);
      continue;
    }

    // Hand off once, keeping the finalized message's id.
    if (!editing) {
      messageId = (await stream.close())?.id;
      editing = true;
    }

    // Phase 2: plain edits on the same message. No two minute ceiling.
    if (messageId && Date.now() - lastEdit > 3_000) {
      await send(new MessageActivityInput(text).withId(messageId));
      lastEdit = Date.now();
    }
  }

  if (!editing) {
    await stream.close();
  } else if (messageId) {
    await send(new MessageActivityInput(text).withId(messageId));
  } else {
    // close() published nothing, so send the buffered text rather than drop it.
    await send(new MessageActivityInput(text));
  }
});
```

Setting an id on an outgoing activity routes the send through the update path, so each edit replaces the finalized message instead of posting a new one.

Plain updates are subject to ordinary activity throttling rather than the streaming path's allowance, so pace them at a cadence measured in seconds rather than sending one per token.

### Starting another streamed message

Calling `stream.emit()` again after the stream has been closed posts a second, separate message in the conversation by reusing the SDK object in your custom handler; it clears its buffered state and starts the new message. The second message gets its own stream id, so it also gets its own fresh two minute window. The message already finalized stays finalized and is not reopened or extended. Use this when your agent has a genuinely separate second answer to stream, not as a way to give one response more time.

### Streaming in threads

The first message of a stream is placed the same way as a response sent with `send()`.
See [Reactive Threading](#reactive-threading-within-a-handler) for the rules.
In a 1:1 chat the stream starts in the thread when the inbound message is inside a thread, and at the top level otherwise.
The SDK applies the same rules in channels: a channel stream starts in the thread of the inbound message.
Without `channelData.replyChain.id` on the inbound activity, the stream is not placed in a reply chain.
Teams itself supports streaming only in 1:1 chats, as noted above.

Only the first chunk of each streamed message uses the reply chain route.
The later chunks are sent to the inbound conversation with the stream id of the first message.
In a 1:1 chat inside a thread, the streamed message appears in the thread pane (checked live with TypeScript and Python).
In a channel, if the service answers 404 to the reply chain route, the first chunk falls back to the inbound conversation, as with `send()`, and stays in the thread.
In a 1:1 chat, a failed reply chain route does not fall back.
An error from the first chunk is logged, and `stream.emit()` does not throw it.

`stream` in a handler is created with this placement already applied.
The `StreamOptions` type, with `replyChain:  chatOrChannelId, replyChainId `, is what `IActivitySender.createStream(ref, options)` accepts when you implement a custom sender, and it is experimental.
The stream falls back to the inbound conversation on a 404 only when `ref.conversation.conversationType` is `channel`.

## @Mention

Sending a message at `@mentions` a user is as simple including the details of the user using the `addMention` method

```typescript

app.on('message', async ({ send, activity }) => {
  await send(new MessageActivityInput('hi!').addMention(activity.from));
});
```

## Targeted Messages

Targeted messages, also known as ephemeral messages, are delivered to a specific user in a shared conversation. From a single user's perspective, they appear as regular inline messages in a conversation. Other participants won't see these messages, making them useful for authentication flows, help or error responses, personal reminders, or sharing contextual information without cluttering the group conversation.

To send a targeted message when responding to an incoming activity, use the `withRecipient` method with the recipient account and set the targeting flag to true.

```typescript

app.on('message', async ({ send, activity }) => {
  // Using withRecipient with isTargeted=true explicitly targets the specified recipient
  await send(
    new MessageActivityInput('This message is only visible to you!')
      .withRecipient(activity.from, true)
  );
});
```

### Prompt Preview

:::info[Coming Soon]
Prompt Preview is coming soon in June 2026.
:::

Prompt Preview shows a compact, collapsible preview of the targeted message your agent is replying to, helping carry context from a private user-to-agent message into the reply.

#### Prompt Preview in targeted reply

In a targeted (private) reply, both the prompt preview and the bot response are visible only to the targeted user.

![Prompt Preview in a targeted reply.](/screenshots/private-prompt-preview.png)

#### Prompt Preview in public reply

In a public reply, the same prompt preview appears above the bot response and is visible to everyone in the conversation.

![Prompt Preview in a public reply.](/screenshots/public-prompt-preview.png)

In reactive scenarios, when replying to an inbound targeted activity through `send()` or `reply()`, the SDK automatically includes targeted message info.

For proactive scenarios (using `app.send()`), attach targeted message info using the targeted message ID you are replying to.

```typescript

const targetedMessageId = '1772050244572';
const conversationId = '19:groupchat-id@thread.v2';
const userAccount: Account = {
  id: '29:1AbCDef...',
  name: 'Adele Vance',
};

const targetedMessage = new MessageActivityInput('Here is the result!')
  .addTargetedMessageInfo(targetedMessageId)
  .withRecipient(userAccount, true);

// Targeted reply (only the user sees it)
await app.send(conversationId, targetedMessage);

// OR public reply (everyone sees it)
const publicMessage = new MessageActivityInput('Here is the result!')
  .addTargetedMessageInfo(targetedMessageId);
await app.send(conversationId, publicMessage);
```

## Reactions

Reactions allow your agent to add or remove emoji reactions on messages in a conversation, and to receive reactions added by users. See the [Message Reactions guide](../in-depth-guides/message-reactions.md) for full coverage.

## Threading

In Teams channels and 1:1 agent chats, the SDK can send messages into reply chains (threads).
The SDK provides helpers to simplify working with threads.

In this guide, top-level means a message in the main chat or channel conversation, and in-thread means a message in a reply chain (thread).

Where a message lands depends on the method and the scope:

| Method | 1:1 chat | Channel |
|---|---|---|
| `context.send()` | Matches level of user's message | In-thread |
| `context.reply()` | Matches level of user's message | In-thread |

The SDK places a response in a reply chain only when `channelData.replyChain.id` is present on the inbound activity.
With no `replyChain`, these methods send to the inbound conversation.

The reply method, `context.reply()`, also quotes the user's message.
To reply in-thread to a top-level message in a 1:1 chat, or to send a top-level message from inside a thread, see [Proactive Messaging](./sending-messages/proactive-messaging.md#proactive-threading).

### Reactive Threading (Within a Handler)

When your agent receives a message in a reply chain (thread), the context sends into that same reply chain.
In a channel, replying to a top-level post also sends into that post's thread.
In a 1:1 chat, a response to a top-level message is also top-level, and a thread is started explicitly with a proactive reply (see below).
Use `send()` to send reactively without quoting, or `reply()` to send a reactive, quoted reply.
In a channel the answer is in the same thread, and in a 1:1 chat it matches the level of the inbound message.

```typescript
app.on('message', async ({ send, reply }) => {
  // Send reactively without quoting.
  // In a channel, sends in the same thread.
  // In a 1:1 chat, matches the level of the inbound message.
  await send('Acknowledged');

  // Send a reactive, quoted reply.
  // In a channel, sends in the same thread.
  // In a 1:1 chat, matches the level of the inbound message.
  await reply('Got it!');
});
```

#### Where the SDK takes the destination from

The SDK reads `channelData` on the inbound activity to find the destination and does not parse the conversation id.
To tell a channel from a chat, it also uses `conversation.conversationType`: `channel` selects the channel rules below, and anything else selects the chat rules.

- **Channel:** a response to a `message` is sent through the reply chain `channelData.replyChain.id` of the channel `channelData.channel.id`.
  If the reply chain route answers HTTP `404`, the SDK sends to the inbound conversation instead.
  For a channel message, the inbound conversation id already identifies the same thread, so the response stays in it.
  The SDK checks only the status code.
- **1:1 chat:** when the inbound activity is inside a thread (see the table below), the response is sent through the reply chain `channelData.replyChain.id` of the chat `channelData.chat.id`.
  Otherwise it is sent to the inbound conversation, at the top level.
  An error from the reply chain route is thrown to your handler, and the response is not sent anywhere else.
- **Missing ids:** when `channelData.replyChain`, or the id of the scope (`channelData.chat.id` in a 1:1 chat, `channelData.channel.id` in a channel), is missing, the SDK does not use the reply chain route and sends to the inbound conversation.
  In a channel, the inbound conversation id can still keep the response in the thread.

Typing indicators sent with `send()` are not placed.
Only message responses are.

#### Responding to edits, reactions and card actions

In a 1:1 chat the SDK also places a response to an inbound activity that is not a typed message.
It compares `channelData.replyChain.id` with the id of the message the activity acts on, and the activity is inside a thread when they differ.

| Inbound activity | Message it acts on | Inside a thread when |
|---|---|---|
| `message` typed by the user | The message, `activity.id` | `replyChain.id` differs from `activity.id` |
| `messageUpdate`, `messageDelete` | The edited or deleted message, `activity.id` | `replyChain.id` differs from `activity.id` |
| `messageReaction` | The reacted message, `replyToId` | `replyToId` is a plain number other than `0` that differs from `replyChain.id` |
| `message` with `channelData.source`, sent by a card button (`messageBack`, `Action.Submit`) | The message with the card, `replyToId` | `replyToId` is a plain number other than `0` that differs from `replyChain.id` |
| `invoke` with `value.messagePayload` (a message action) | The message the action ran on, `value.messagePayload.id` | That id is a plain number other than `0` that differs from `replyChain.id` |
| Any other `invoke` with a `replyToId`, such as `adaptiveCard/action` from `Action.Execute`, `task/fetch` and `task/submit` from a card button, or `message/submitAction` feedback | The message with the card or feedback buttons, `replyToId` | `replyToId` is a plain number other than `0` that differs from `replyChain.id` |
| `invoke` with neither, such as a compose box action or sign-in | None | `replyChain.id` is present |
| Any other activity, such as `installationUpdate` or `conversationUpdate` | None | Never |

So a response to a card button clicked on a top-level message is top-level, and a response to a card button clicked on a card inside a thread is sent into that thread.
A `replyToId` that is not a plain number, such as an encrypted id, is never inside a thread.

The SDK uses `conversation.conversationType` only to tell a channel from a chat.
In a chat it does not require a particular type to place a response.
In the 1:1 chat of an agentic user, card button clicks (`messageBack`, `Action.Submit`, `Action.Execute`) and the `task/fetch` and `task/submit` invokes they open have arrived with `conversationType` set to `groupChat`, a service issue that is being corrected.
The SDK places the response to them from `channelData` like any other activity, so a response to a click inside a thread is sent into that thread.

In a channel only a `message` is placed, and a card button click that arrives as a `message` is placed like any other.
The SDK sends a response to any other activity in a channel to the inbound conversation.

For proactive threading (sending to a thread outside of a handler), see [Proactive Messaging](./sending-messages/proactive-messaging.md#proactive-threading).

#### Experimental status

The reply chain types and operations are marked `@experimental` (diagnostic `ExperimentalTeamsThreading`) in the API reference.
The marker is documentation only, with no compiler or runtime warning.
It covers `channelData.chat`, `channelData.replyChain` and `channelData.source`, `api.conversations.createReplyChainActivity` and `createTargetedReplyChainActivity` (and `createReplyChain` and `createTargetedReplyChain` on `ConversationActivityClient`), `ActivitySenderOptions.replyChainId` and `StreamOptions`.
`send`, `reply` and `app.reply` are not marked.

## Quoted Replies

Quoted replies let your agent reference a previous message in the conversation. When a user sends a message that quotes another message, your agent receives structured metadata about the quoted content. Your agent can also send messages that quote previous messages.

### Receiving Quoted Replies

When a user quotes a message and sends it to your agent, the quoted reply metadata is available on the inbound activity. Use the `getQuotedMessages()` method to access all quoted reply entities.

```typescript
app.on('message', async ({ activity, reply }) => {
  const quotes = activity.getQuotedMessages();

  if (quotes.length > 0) {
    const quote = quotes[0].quotedReply;
    await reply(
      `You quoted message ${quote.messageId} from ${quote.senderName}: "${quote.preview}"`
    );
  }
});
```

Each quoted reply entity contains the quoted message's ID, sender information, a preview of the quoted text, and whether the quoted message has been deleted.

### Sending a Quoted Reply

When your agent calls `reply()`, the SDK automatically stamps a quoted reply entity referencing the inbound message when its id is a plain message id. The reply will appear as a quoted reply in Teams.

This applies when `activity.id` is a plain numeric message id, as it is for messages, edits and deletes.
Otherwise, for example a card button click or another invoke with an id like `f:<guid>`, the SDK does not quote, and the response is sent like `send()`, without a quote.
To quote in a card or invoke handler, quote the message that holds the card with `quote()` and its `replyToId`.

```typescript
app.on('message', async ({ reply }) => {
  // reply() automatically quotes the inbound message
  await reply('Got it!');
});
```

To quote a different message in the same conversation (not the inbound message), use the `quote()` method with the message ID you want to quote.

```typescript
app.on('message', async ({ quote }) => {
  // Quote a specific message by its ID
  const parentMessageId = '1772050244572';
  await quote(parentMessageId, 'Referencing an earlier message');
});
```

### Building Quoted Replies for Proactive Send

For proactive scenarios (using `app.send()`) or when quoting multiple messages, use the `addQuote()` method on a message activity. Pass the message ID and an optional response text.

```typescript

const parentMessageId = '1772050244572';
const firstMessageId = '1772050244573';
const secondMessageId = '1772050244574';

// Single quote with response below it
let msg = new MessageActivityInput()
  .addQuote(parentMessageId, 'Here is my response');
await app.send(conversationId, msg);

// Multiple quotes with interleaved responses
msg = new MessageActivityInput()
  .addQuote(firstMessageId, 'response to first')
  .addQuote(secondMessageId, 'response to second');
await app.send(conversationId, msg);

// Grouped quotes — omit response to group quotes together
msg = new MessageActivityInput('see below for previous messages')
  .addQuote(firstMessageId)
  .addQuote(secondMessageId, 'response to both');
await app.send(conversationId, msg);
```
