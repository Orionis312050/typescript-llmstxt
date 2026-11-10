---
title: "State Management"
sidebar_position: "10"
---

# State Management

:::info[SDK 2.1]
Built-in state management is a **Teams SDK 2.1** feature.
:::

The Teams SDK provides built-in, per-turn state for storing conversation and user data across activities. State is loaded before each activity handler runs, saved automatically after the turn, and exposed through `ctx.state`.

## Setup

State is disabled by default. Enable it with the `state: true` app option:

```typescript

const app = new App({
  state: true,
});
```

Without a dedicated provider, state uses the Teams SDK's process-local, in-memory `LocalStorage` implementation. This is useful for local development, but values are lost when the process restarts and aren't shared across instances.

:::note
Registering an OAuth flow with `addOAuthFlow()` automatically enables state so pending sign-ins can be associated with the correct flow. Set `state: false` explicitly to fall back to process-local in-memory maps.
:::

## Reading and writing state

Use `ctx.state.conversation` and `ctx.state.user` in an activity handler. Values must be JSON-serializable.

```typescript
app.on('message', async (ctx) => {
  if (!ctx.state) {
    throw new Error('Turn state is not enabled.');
  }

  const count = (ctx.state.conversation.get<number>('messageCount') ?? 0) + 1;
  ctx.state.conversation.set('messageCount', count);

  if (ctx.state.user && ctx.activity.text?.startsWith('my name is ')) {
    ctx.state.user.set(
      'name',
      ctx.activity.text.slice('my name is '.length).trim()
    );
  }

  const name = ctx.state.user?.get<string>('name') ?? 'there';
  await ctx.reply(`Hello, ${name}. Message #${count}.`);
});
```

Use `has()` to check for a value, `delete()` to remove one, and `clear()` to remove all values. If you mutate an object or array returned by `get()`, call `set()` with the updated value so the scope is marked for persistence.

Conversation state is shared by everyone in the current conversation. User state is scoped to the current sender **within that conversation** and is unavailable when an activity has no usable sender ID.

The SDK writes a scope to storage only when it changes during the current activity. After all handlers for that activity finish, the SDK saves those changes and closes the turn state. Read or update state only while handling the activity. Don't capture it for timers or background tasks because accessing it after the turn ends throws `TurnStateSealedError`.

## Clearing state

Remove a value with `delete()`, or clear a scope with `clear()`. To remove both scopes from the backing store:

```typescript
await ctx.state.delete();
```

Values written after `delete()` are saved normally at the end of the current turn.

## Scaling to distributed state

For production or multi-instance deployments, implement the Teams SDK's `IStorage<string, string>` contract with a shared, durable backend and pass it through `state.storage`. The state API used by handlers doesn't change:

```typescript

function createApp(durableStorage: IStorage<string, string>): App {
  return new App({
    state: {
      storage: durableStorage,
      keyPrefix: 'my-app',
    },
  });
}
```

The SDK serializes each scope as a JSON string and replaces the complete scope on save, so concurrent turns use last-writer-wins semantics. Configure expiry, retries, and other storage-specific behavior on your storage implementation.
