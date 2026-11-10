---
title: "Static Pages"
sidebar_position: "1"
---

# Static Pages

The `App` class lets you host web apps in the agent. This is useful for an efficient inner loop — you can build, deploy, and sideload both an agent and a Tab app inside Teams in a single step. It's also useful in production, since it makes it straightforward to host a simple experience such as an agent configuration page or a Dialog.

To host a static tab web app, call the `app.tab()` function and provide an app name and a path to a folder containing an `index.html` file to be served up.

```typescript
app.tab('myApp', path.resolve('dist/client'));
```

This registers a route that is hosted at `http://localhost:PORT/tabs/myApp` or `https://BOT_DOMAIN/tabs/myApp`.

## Additional resources

- For more details about Tab apps, see the [Tabs](../tabs.md) in-depth guide.
- For an example of hosting a Dialog, see the [Creating Dialogs](../dialogs/creating-dialogs.md) in-depth guide.
