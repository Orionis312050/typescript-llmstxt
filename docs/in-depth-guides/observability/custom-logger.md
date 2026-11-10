---
title: "🧃 Custom Logger"
sidebar_position: "2"
---

# 🧃 Custom Logger

The `App` will provide a default logger, but you can also provide your own.
The default `Logger` instance will be set to `ConsoleLogger` from the `@microsoft/teams.common` package.

```typescript

// initialize app with custom console logger
// set to debug log level
const app = new App({
  logger: new ConsoleLogger('echo', { level: 'debug' }),
});

app.on('message', async ({ send, activity, log }) => {
  log.debug(activity);
  await send({ type: 'typing' });
  await send(`you said "${activity.text}"`);
});

(async () => {
  await app.start();
})();
```

## Log Levels

Available levels, from most to least severe: `error`, `warn`, `info`, `debug`, `trace`. The default is `info`. Setting `level: 'debug'` emits `error`, `warn`, `info`, and `debug` — but not `trace`.

## Filtering by Logger Name

Each logger is created with a name. You can filter which loggers emit output by providing a name pattern, using `*` as a wildcard.

Each logger has a name. The `pattern` option filters which loggers emit, with `*` as a wildcard. Prefix a pattern with `-` to exclude. Combine with commas.

```typescript
new ConsoleLogger('my-app', { pattern: '@teams*' });        // only SDK loggers
new ConsoleLogger('my-app', { pattern: '*,-@teams/http*' }); // everything except HTTP
```

## Environment Variables

The TypeScript SDK's `ConsoleLogger` reads two environment variables at construction:

| Variable     | Purpose                          | Example                 |
| ------------ | -------------------------------- | ----------------------- |
| `LOG_LEVEL`  | Minimum severity                 | `LOG_LEVEL=debug`       |
| `LOG`        | Logger name pattern (wildcards)  | `LOG=@teams*`           |

Env vars override options passed to the constructor. If you do not pass a logger to `App`, the SDK creates a default `ConsoleLogger` named `@teams/app` — so `LOG_LEVEL=debug` alone is enough to enable debug output.

> **Gotcha:** `LOG` is a name filter, not a toggle. Setting `LOG` to a pattern that doesn't match the default `@teams/app` (for example `LOG=my-app*`) silences the default logger. If in doubt, unset `LOG` to match everything.

## Child Loggers

Call `log.child('scope')` on an existing logger to get a scoped logger. Its name is `parent/scope`, and pattern/level are inherited.

```typescript
app.on('message', async ({ log }) => {
  const msgLog = log.child('message-handler');
  msgLog.debug('processing'); // logged as "@teams/app/message-handler"
});
```
