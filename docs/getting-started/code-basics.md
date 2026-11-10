---
title: "Code Basics"
sidebar_position: "2"
---

# Code Basics

After following the guidance in [the quickstart](./quickstart.md) to create your first Teams application, let's review its structure and key components. This knowledge can help you build more complex applications as you progress.

## Project Structure

When you create a new Teams application, it generates a directory with this basic structure:

```
quote-agent/
├── src/
│   └── index.ts      # Main application code
```

- **src/**: Contains the main application code. The `index.ts` file is the entry point for your application.

## Core Components

Let's break down the simple application from the [quickstart](./quickstart.md) into its core components.

### The App Class

The heart of an application is the `App` class. This class handles all incoming activities and manages the application's lifecycle. It also acts as a way to host your application service.

```typescript title="src/index.ts"

const app = new App();
```

The app configuration includes a variety of options that allow you to customize its behavior, including controlling the underlying server, authentication, and other settings.

### Plugins

Plugins are a core part of the Teams SDK. They allow you to hook into various lifecycles of the application. The lifecycles include server events (start, stop, initialize, etc.), and also Teams Activity events (onActivity, onActivitySent, etc.)

### Message Handling

Teams applications respond to various types of activities. The most basic is handling messages:

```typescript title="src/index.ts"
app.on('message', async ({ send, activity }) => {
  await send({ type: 'typing' });
  await send(`you said "${activity.text}"`);
});
```

This code:

1. Listens for all incoming messages using `app.on('message')`.
2. Sends a typing indicator, which renders as an animated ellipsis (…) in the chat.
3. Responds by echoing back the received message.

:::info
Type safety is a core tenet of this version of the SDK. You can change the activity `name` to a different supported value, and the type system will automatically adjust the type of activity to match the new value.
:::

### Application Lifecycle

Your application starts when you run:

```typescript title="src/index.ts"
await app.start();
```

This code initializes your application server and, when configured for Teams, also authenticates it to be ready for sending and receiving messages.

## Next Steps

Now that you understand the basic structure of your Teams application, you're ready to [run it in Teams](running-in-teams). You'll use the Teams Developer CLI to register your bot and sideload it into Teams.

After that, you can:

- Add more activity handlers for different types of interactions. See [Listening to Activities](../essentials/listening-to-activities.md) for more details.
- Integrate with external services using the [API Client](../essentials/teams-api-client.md).
- Add interactive [cards](../in-depth-guides/adaptive-cards.md) and [dialogs](../in-depth-guides/dialogs.md).

Continue on to the next page to learn about these advanced features.

## Other Resources

- [Essentials](../essentials.md)
- [Teams concepts](../teams/user-authentication/sso-troubleshooting.md)
- [Teams developer tools](../developer-tools/exposing-your-local-app-with-a-tunnel.md)
