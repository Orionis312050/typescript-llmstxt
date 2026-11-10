---
title: "Listening To Activities"
sidebar_position: "1"
---

# Listening To Activities

An **Activity** is the Teams‑specific payload that flows between the user and your bot.
Where _events_ describe high‑level happenings inside your app, _activities_ are the raw Teams messages such as chat text, card actions, installs, or invoke calls.

The Teams SDK exposes a fluent router so you can subscribe to these activities with `app.on('<route>', …)`.

```mermaid
flowchart LR
    Teams["Teams"]:::less-interesting
    Server["App Server"]:::interesting
    ActivityRouter["Activity Router (app.on())"]:::interesting
    Handlers["Your Activity Handlers"]:::interesting

    Teams --> |Events| Server
    Server --> |Activity Event| ActivityRouter
    ActivityRouter --> |handler invoked| Handlers

    classDef interesting fill:#b1650f,stroke:#333,stroke-width:4px;
    classDef less-interesting fill:#666,stroke:#333,stroke-width:4px;
```

Here is an example of a basic message handler:

```typescript
app.on('message', async ({ activity, send }) => {
  await send(`You said: ${activity.text}`);
});
```

In the above example, the `activity` parameter is of type `MessageActivity`, which has a `text` property. You'll notice that the handler here does not return anything, but instead handles it by `send`ing a message back. For message activities, Teams does not expect your application to return anything (though it's usually a good idea to send some sort of friendly acknowledgment!).

[Other activity types](./on-activity/activity-type-reference.md) have different properties and different required results. For a given handler, the SDK will automatically determine the type of `activity` and also enforce the correct return type.

## Slash Commands

Slash commands are manifest-declared commands users run from the compose box. To enable slash commands, set `supportsTargetedMessages: true` in your app manifest (v1.29 or later) under the `bots` section. You can opt in with an explicit command list by declaring specific commands using `commandLists` with `triggers: ["slash"]`, which Teams shows in the slash menu when a user types `/`. Without a command list, users can still invoke your agent via `/agent-name` and provide free-form input.

```json
{
  "bots": [
    {
      "botId": "{{BOT_ID}}",
      "scopes": ["personal", "team", "groupChat"],
      "supportsTargetedMessages": true,
      "commandLists": [
        {
          "scopes": ["team", "groupChat"],
          "triggers": ["slash"],
          "commands": [
            { "title": "Review", "description": "Review a document" }
          ]
        }
      ]
    }
  ]
}
```

When a user sends a slash command, it appears as a private message visible only to them. Your agent can reply privately or, when appropriate, share a response with the broader group or channel.

Slash commands arrive as normal message activities with the targeted flag set on the activity's recipient object.

```typescript
app.on('message', async ({ activity, send, next }) => {
  if (activity.recipient?.isTargeted) {
    await send(`Received slash command: ${activity.text}`);
    return;
  }

  await next();
});
```

## Activity handler chaining

The `on` activity handlers follow a [middleware](https://www.patterns.dev/vanilla/mediator-pattern/) pattern similar to how `express` middlewares work. This means that for each activity handler, a `next` function is passed in which can be called to pass control to the next handler. This allows you to build a chain of handlers that can process the same activity in different ways.

```typescript
app.on('message', async ({ next }) => {
  console.log('global logger');
  next(); // pass control onward
});
```

```typescript
app.on('message', async ({ activity, next }) => {
  if (activity.text === '/help') {
    await send('Here are all the ways I can help you...');
    return;
  }

  // Conditionally pass control to the next handler
  next();
});
```

```typescript
app.on('message', async ({ activity }) => {
  // Fallthrough to the final handler
  await send(`Hello! you said ${activity.text}`);
});
```

## Activity Reference

For a list of supported activities that your application can listen to, see the [activity reference](./on-activity/activity-type-reference.md).
