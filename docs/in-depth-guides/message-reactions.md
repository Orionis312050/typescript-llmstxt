---
title: "Message Reactions"
sidebar_position: "6"
---

# Message Reactions

![Reaction emoji sent by bot](/screenshots/reaction-emoji.png)

Message reactions allow your agent to add or remove emoji reactions on messages in a Teams conversation, and to react to reactions added by users. This gives your agent a quick, low-friction way to acknowledge messages or signal status without sending a full text reply.

## Adding a Reaction

To add a reaction to a message, use the reactions client through the API client:

```typescript
app.on('message', async ({ activity, api, send }) => {
  await send("Hello! I'll react to your message.");

  // Add a reaction to the incoming message
  await api.conversations.addReaction(activity.conversation.id, activity.id, 'like');
});
```

## Removing a Reaction

You can also remove reactions that your agent has previously added:

```typescript
app.on('message', async ({ activity, api }) => {
  // First, add a reaction
  await api.conversations.addReaction(activity.conversation.id, activity.id, 'heart');

  // Wait a bit, then remove it
  await new Promise((resolve) => setTimeout(resolve, 2000));
  await api.conversations.deleteReaction(activity.conversation.id, activity.id, 'heart');
});
```

## Receiving Reactions

Your agent can also listen for reactions that users add or remove on messages your agent sent.
In 1:1 chats and channels, reactions on messages your agent did not send are not delivered.
The activity carries `reactionsAdded` and `reactionsRemoved` arrays.
Its `replyToId` is the id of the message that was reacted to.

```typescript
app.on('messageReaction', async ({ activity }) => {
  for (const reaction of activity.reactionsAdded ?? []) {
    console.log(`User added reaction: ${reaction.type}`);
  }

  for (const reaction of activity.reactionsRemoved ?? []) {
    console.log(`User removed reaction: ${reaction.type}`);
  }
});
```

A message sent from a reaction handler is placed relative to the reacted message.
In a 1:1 chat, when that message is inside a thread, the message is sent into the thread.
See [Responding to edits, reactions and card actions](../essentials/sending-messages.md#responding-to-edits-reactions-and-card-actions).

## Available Reaction Types

The SDK ships a small set of named reaction constants for the most common reactions.

Any string-valued reaction ID is accepted, so you can pass any reaction ID from the [Teams reactions reference page](https://learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/teams-reactions-reference).

- `'like'` — 👍
- `'heart'` — ❤️
- `'1f440_eyes'` — 👀
- `'2705_whiteheavycheckmark'` — ✅
- `'launch'` — 🚀
- `'1f4cc_pushpin'` — 📌

## Skin Tones

Reactions tagged **Diverse** in the Teams reactions reference support five skin-tone variants. Append `-tone1` through `-tone5` to the reaction ID to select a variant:

```text
1f44b_wavinghand-tone4
```

## Rate Limits

Agents are limited to **two reactions per second**. Calls that exceed this limit are throttled and return a `429` response. Implement exponential backoff and honor the `Retry-After` header.

## Best Practices

- **Use reactions sparingly.** Reactions work best when they signal something specific (acknowledgement, completion, an error). Reacting to every message creates noise.
- **Match the reaction to the context.** Different reactions carry different meanings. Choose one that aligns with your agent's purpose.
- **Replace, don't stack.** If your agent has already added a reaction and needs to update it, remove the existing reaction first, then add the new one. Adding a reaction that's already present is a no-op.
- **Handle errors.** Reaction operations can fail if the message no longer exists or the agent isn't a member of the conversation. Handle exceptions appropriately.

## Differences from Feedback

Message reactions are different from the [feedback](./feedback.md) feature:

- **Reactions** are quick emoji responses your agent adds to or removes from messages programmatically, and reactions users add to messages your agent can listen for.
- **Feedback** are interactive UI components (like/dislike buttons) that users can click to provide structured feedback on agent responses.

Use reactions when your agent wants to acknowledge or respond to a message with emoji. Use feedback when you want to collect user opinions about your agent's responses.
