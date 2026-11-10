---
title: "Handling Tool Calls"
sidebar_position: "3"
---

# Handling Tool Calls

An interactive widget can call tools on your app.
When it does, the Teams client sends your app an `htmlwidget/calltool` invoke activity.
The SDK gives you a typed handler for these invokes and a typed response shape, so you never parse the raw invoke body.

:::warning Preview
HTML Widgets are in preview and the API surface may change. HTML Widgets are planned for public preview on 11/16.
:::

## The flow

1. Your app sends a widget whose HTML calls `tools/call` over the MCP Apps protocol.
2. The Teams client forwards that call to your app as an `htmlwidget/calltool` invoke.
3. Your handler inspects the tool name and arguments, does its work, and returns a result.
4. The Teams client delivers the result back to the widget, which re-renders.

## Registering a handler

Register a handler for the `widget.callTool` event. The activity value is typed, so you get the tool `name` and `arguments` directly.

```typescript

app.on('widget.callTool', async ({ activity }) => {
  const { name, arguments: args } = activity.value;

  let callToolResult: IMcpUiCallToolResult;
  switch (name) {
    case 'getTime':
      callToolResult = {
        content: [{ type: 'text', text: new Date().toLocaleTimeString() }],
        structuredContent: { time: new Date().toISOString() },
        isError: false,
      };
      break;
    default:
      callToolResult = {
        content: [{ type: 'text', text: `Unknown tool: ${name}` }],
        isError: true,
      };
      break;
  }

  return buildCallToolResponse(callToolResult);
});
```

## Returning a result

A tool call result carries a list of content blocks, optional structured content the widget can render from, and an `isError` flag.
Teams currently renders `text` content; the other MCP content types (`image`, `audio`, `resource`) are defined for forward compatibility.

Return `buildCallToolResponse(result)`, which produces an `IHtmlWidgetCallToolResponse` with `responseType` set to `htmlwidget/calltoolresult` and the result as `callToolResult`.
The result's `content` is an array of content blocks; `structuredContent` holds data the widget can render from; `isError` signals a failure.
`buildCallToolResponse` always sets all three fields, defaulting to `[]`, `{}` and `false`, because an omitted field can reach the widget as `null`.
Use `buildCallToolResult(result)` when you need only the result without the response wrapper.

```typescript

const callToolResult: IMcpUiCallToolResult = {
  content: [{ type: 'text', text: 'Refreshed!' }],
  structuredContent: { counter: 1, lastAction: 'refresh' },
  isError: false,
};

return buildCallToolResponse(callToolResult);
```

## Client-supplied call context

The Teams client adds an `__msteams__` entry to the `arguments` of every call, carrying the `mcpUrl` and `correlationId` from the widget payload when they were set. The client overwrites any value the widget supplied, and the SDK passes `arguments` to your handler untouched. Ignore `__msteams__` when you read your tool's arguments, validate the other arguments as untrusted input, never throw from the handler, and do not authorize from `activity.from`.

The `IHtmlWidgetCallToolContext` type describes the `__msteams__` entry, with optional `mcpUrl` and `correlationId` fields.

## Throttling in large contexts

In large or busy contexts the Teams platform may throttle widget tool calls, so your app can receive fewer `htmlwidget/calltool` invokes than the widget issued. This throttling happens on the platform before the invoke reaches your app, so there is no error or `429` status for your handler to catch and nothing for the SDK to retry. Design your tool handlers so that a dropped call is not fatal: keep them idempotent, and let the widget reflect its own state rather than assuming every call reaches your app.
