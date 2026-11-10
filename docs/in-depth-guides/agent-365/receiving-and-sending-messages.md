---
title: "Receiving and sending messages"
sidebar_position: "3"
---

# Receiving and sending messages

Configure the blueprint client ID, credential, and tenant ID in your app's SDK credentials. Teams SDK uses them to acquire tokens for the Agentic User.

## Reactive messaging

An inbound Agent 365 activity includes the Agentic User in `activity.recipient`. The handler context automatically uses that identity for replies, sends, and API calls.

```typescript

const app = new App();

app.on('message', async ({ activity, reply }) => {
  const agenticIdentity = getAgenticIdentity(activity.recipient);
  if (!agenticIdentity?.agenticUserId) {
    throw new Error('The activity is not addressed to an Agentic User.');
  }

  await reply(
    `Hi! I'm an Agentic User, and my user ID is ${agenticIdentity.agenticUserId}. Nice to meet you!`
  );
});
```

:::note Who sends the reply?
The example sends the reply as the Agentic User from the inbound activity, not as the blueprint. The blueprint authenticates the reusable runtime, but it isn't the conversational sender.
:::

## Bot capabilities

Most Teams capabilities available to standard bots are also available to Agentic Users, including cards, reactions, and API calls. For inbound activities, the handler's API client uses the Agentic User identity automatically.

This example adds a reaction to an incoming message as the Agentic User:

```typescript
app.on('message', async ({ activity, api }) => {
  await api.conversations.addReaction(activity.conversation.id, activity.id, 'like');
});
```

For more reaction options, see [Message Reactions](../message-reactions.md).

## Proactive messaging

A proactive operation has no inbound activity that identifies the sender. Create an `AgenticIdentity` associated with the Agentic User, then pass it as the operation scope. The app helper supplies the required blueprint ID from app configuration; the Agentic App Instance and Agentic User IDs select the user-backed scope used in this example.

```mermaid
flowchart TD
  R[Reactive activity] --> RI[SDK reads Agentic User<br/>from recipient]
  RI --> RS[reply / send / API call]
  P[Proactive job] --> PI[App supplies Agentic User]
  PI --> PS[AgenticIdentity-scoped<br/>send / API call]
```

If your app will send proactive messages later, store the conversation ID, service URL, Agentic App Instance ID, Agentic User ID, and tenant ID in persistent storage when it receives an activity.

```typescript
const agenticIdentity = app.getAgenticIdentity({
  agenticAppId,
  agenticUserId,
});

await app.send(conversationId, 'Your scheduled update is ready.', {
  agenticIdentity,
});

const api = app.api.forAgenticIdentity(agenticIdentity);

await api.conversations.addReaction(conversationId, activityId, 'like');
```
