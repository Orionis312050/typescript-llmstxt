---
title: "Proactive Messaging"
sidebar_position: "1"
---

# Proactive Messaging

In [Sending Messages](#), you were shown how to respond to an event when it happens. However, there are times when you want to send a message to the user without them sending a message first. This is called proactive messaging. You can do this by using the `send` method on your app instance. This approach is useful for sending notifications or reminders to the user.

The main thing to note is that you need to have the `conversationId` of the chat or channel that you want to send the message to. It's a good idea to store this value somewhere from an activity handler so that you can use it for proactive messaging later.

```typescript

// ...

// This would be some persistent storage
const myConversationIdStorage = new Map<string, string>();

// Installation is just one place to get the conversation id. All activities
// have the conversation id, so you can use any activity to get it.
app.on('install.add', async ({ activity, send }) => {
  // Save the conversation id in
  myConversationIdStorage.set(activity.from.aadObjectId!, activity.conversation.id);

  await send('Hi! I am going to remind you to say something to me soon!');
  notificationQueue.addReminder(activity.from.aadObjectId!, sendProactiveNotification, 10_000);
});
```

Then, when you want to send a proactive message, you can retrieve the `conversationId` from storage and use it to send the message.

```typescript

// ...

const sendProactiveNotification = async (userId: string) => {
  const conversationId = myConversationIdStorage.get(userId);
  if (!conversationId) {
    return;
  }
  const activity = new MessageActivityInput('Hey! It\'s been a while. How are you?');
  await app.send(conversationId, activity);
};
```

:::tip
In this example, you see how to get the `conversationId` using one of the activity handlers. This is a good place to store the conversation id, but you can also do this in other places like when the user installs the app or when they sign in. The important thing is that you have the conversation id stored somewhere so you can use it later.
:::

## Targeted Proactive Messages

Targeted messages, also known as ephemeral messages, are delivered to a specific user in a shared conversation. From a single user's perspective, they appear as regular inline messages in a conversation. Other participants won't see these messages.

When sending targeted messages proactively, you must explicitly specify the recipient account.

```typescript

// When sending proactively, you must provide an explicit recipient account
const sendTargetedNotification = async (conversationId: string, recipient: Account) => {
  await app.send(
    conversationId,
    new MessageActivityInput('This is a private notification just for you!')
      .withRecipient(recipient, true)
  );
};
```

## Proactive Threading

The SDK sends messages into reply chains (threads) in channels and 1:1 chats.
Whether a thread is shown depends on the service and the Teams client.
The SDK sends a response from a handler through the reply chain route only when `channelData.replyChain.id` is present on the inbound activity.
With no `replyChain`, the SDK skips that route and sends to the inbound conversation.
In a channel, the suffix of that conversation id can still keep the response in the thread.

In this guide, top-level means a message in the main chat or channel conversation, and in-thread means a message in a reply chain (thread).

The `app` methods take explicit ids, so they work outside a handler.
Given the id of the chat or channel (`channelData.chat.id` or `channelData.channel.id`), in 1:1 chats and channels they place the message as follows:

| Method | 1:1 chat | Channel |
|---|---|---|
| `app.send()` | Top-level | Top-level |
| `app.reply()` | In-thread | In-thread |

To proactively send a message as a reply in a thread, use `app.reply()` with the id of the chat or channel and the id of the top-level message of the thread.
Take the ids from `channelData` on the inbound activity:

- **Chat id:** `channelData.chat.id`, in a 1:1 chat.
- **Channel id:** `channelData.channel.id`, in a channel.
- **Reply chain id:** `channelData.replyChain.id`, the id of the top-level message of the thread.
  - A `replyChain.id` that matches the message's own `activity.id` indicates a top-level post.
    A different `replyChain.id` indicates the message is in-thread.
  - If `replyChain` is absent, there is no reply chain id to take from the inbound activity, so pass an id you stored yourself.
    The SDK does not take it from the conversation id.

```typescript
// chatOrChannelId is channelData.chat.id or channelData.channel.id
// replyChainId is channelData.replyChain.id
await app.reply(chatOrChannelId, replyChainId, 'Thread update!');
```

To send a new top-level (non-threaded) message instead, send to the id of the chat or channel with `app.send()`.

```typescript
// Send a new top-level message to the chat or channel
await app.send(chatOrChannelId, 'Hello!');
```

If the reply chain call fails, `app.reply()` throws the error to your code.
It does not fall back to another route.

Inside a handler, a response goes to the `serviceUrl` of the inbound activity.
`app.send()` and `app.reply()` have no inbound activity and take no service URL argument, so they go to the app's service URL: the `serviceUrl` app option, else the `SERVICE_URL` environment variable, else `https://smba.trafficmanager.net/teams`.

A service URL without its cloud segment can be accepted and still post the message flat, outside the thread.
Store the full service URL of the inbound activity, and pass or configure it for proactive calls, instead of rebuilding or trimming it.

The API client's reply operation, `api.conversations.replyToActivity()`, is a plain send to the conversation it is given.
It posts to the activities route of that conversation and does not place the message in a reply chain.
To call the reply chain route directly, use `api.conversations.createReplyChainActivity()`, which is experimental.
It takes the chat or channel id and the reply chain id, and has no fallback.

If you stored a conversation id from a channel message, it ends in `;messageid=<root post id>`.
That form identifies the thread, not the chat or channel, and the reply chain endpoint rejects it.
Store `channelData.chat.id` or `channelData.channel.id` and `channelData.replyChain.id` instead.

A proactive call has no inbound activity to take an identity from.
When the recipient is an agentic user, pass its identity from the inbound activity's recipient.
Without it, the call uses the app's own credentials instead of the agentic user's.
An agentic app cannot get an app-only token, so the call fails with an authentication error (`AADSTS82001`) and nothing is sent.

```typescript

// agenticIdentity is undefined unless the recipient is an agentic user
const agenticIdentity = getAgenticIdentity(context.activity.recipient);
await app.reply(chatOrChannelId, replyChainId, 'Thread update!', { agenticIdentity });
```

For reactive threading (within a handler), see [Threading](#threading).
