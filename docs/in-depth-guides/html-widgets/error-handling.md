---
title: "Error Handling"
sidebar_position: "6"
---

# Error Handling

A widget can fail at four distinct points: when your app sends the message, when the payload is malformed, when a tool call runs, and when the Teams client renders the HTML. Each surfaces differently, and only some are reported back to your app.

:::warning Preview
HTML Widgets are in preview and the API surface may change. HTML Widgets are planned for public preview on 11/16.
:::

## Sending a widget

A widget travels inside a normal message, so sending one is a normal message send. If the Bot Connector rejects the request you get a standard Bot Framework HTTP error, such as `400` for a malformed activity, `401` or `403` for authentication problems, or `429` when throttled. See [HTTP status codes](https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-api-reference?view=azure-bot-service-4.0#http-status-codes) in the Bot Framework Connector reference for the full list. Handle these the way you handle any other message send failure.

## Payload validation

The build helpers (`buildHtmlWidgetMessage` and `buildHtmlWidgetMarkdown`) validate the payload before producing the message. They throw an `Error` when:

- `name` is empty,
- `html` is empty, or
- `domain` is not a valid `https://` URL.

Catching this at send time surfaces problems that would otherwise become a silent client-side "Couldn't load widget" failure with no signal back to your app.

```typescript

let message;
try {
  message = buildHtmlWidgetMessage({
    name: 'Simple Widget',
    html: '<div>Hello from a widget</div>',
    domain: 'https://teams.microsoft.com',
  });
} catch (err) {
  // Thrown when name or html is empty, or domain is not a valid https:// URL.
  console.error(`Invalid widget payload: ${(err as Error).message}`);
  return;
}

await send(message);
```

## Tool call errors

When a widget calls one of your tools, return a result with `isError: true` to signal a failure the widget can render. This is an application-level error carried in a successful invoke response, not a transport failure, so the widget receives it and can show its own error state. See [Handling tool calls](./handling-tool-calls.md) for the full handler.

```typescript

app.on('widget.callTool', async ({ activity }) => {
  const { name } = activity.value;

  return buildCallToolResponse({
    content: [{ type: 'text', text: `Unknown tool: ${name}` }],
    isError: true,
  });
});
```

Because the widget owns its own protocol code, how it reports an error result, and what it does when a request never gets a response, is up to your widget. Handle an error response in place of a result, and decide how long to wait before giving up on a request that gets no response.

## Render errors are not reported back

If the Teams client cannot render the widget HTML it shows a "Couldn't load widget" message inline in the chat. This is a client-side failure: it is **not** delivered to your app as an activity, invoke, or callback. Your app has no way to detect it after the message is sent.

Because of this, treat rendering as fire-and-forget and prevent problems before they ship:

- Let the build helpers validate the payload, as shown above, rather than assembling the fenced code block by hand.
- Use the [security policy](./security-policy.md) helpers so external references are not blocked at render time.
- Test the widget end to end in a real client before shipping.
