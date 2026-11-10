---
title: "Security Policy"
sidebar_position: "4"
---

# Security Policy

A widget's `securityPolicy` declares which origins the widget is allowed to reach for network requests, static resources, nested iframes, and base URIs.
The Teams client enforces this policy through the browser's Content Security Policy.
The SDK adds two helpers to make working with the policy easier.

:::warning Preview
HTML Widgets are in preview and the API surface may change. HTML Widgets are planned for public preview on 11/16.
:::

:::note Not a security boundary
These helpers are development-time convenience aids for catching policy mistakes early, not a security boundary in themselves. Treat them as a supplement to, not a replacement for, the Teams client's runtime enforcement and your own review and hardening of the HTML your app sends.
:::

## Declaring a policy

A security policy is a set of allow-lists you attach to a widget payload under `securityPolicy`. Each field lists the origins the widget may use for one kind of request:

| Field | Covers |
|-------|--------|
| `connectDomains` | Network requests the widget makes (`fetch`, `XMLHttpRequest`, WebSocket). |
| `resourceDomains` | Static resources the widget loads (scripts, styles, fonts, images). |
| `frameDomains` | Origins the widget may embed in nested iframes. |
| `baseUriDomains` | Values the widget may set for its document base URI. |

Attach the policy to the payload you send. When you omit `securityPolicy`, the SDK applies a restrictive default that allows only same-origin and `data:` resources.

```typescript

// Attach a securityPolicy to the payload. The build helper carries it through
// to the widget block; omit it and the SDK applies a restrictive default.
const message = buildHtmlWidgetMessage({
  name: 'Chart Widget',
  html: '<div id="chart"></div>',
  domain: 'https://teams.microsoft.com',
  securityPolicy: {
    connectDomains: ['https://api.contoso.com'],
    resourceDomains: ["'self'", 'data:', 'https://cdn.contoso.com'],
    frameDomains: [],
    baseUriDomains: [],
  },
});

await send(message);
```

## Validating a policy

The helper `validateSecurityPolicy` scans your HTML for external references and returns a warning for any origin that is not covered by the appropriate policy field.
This is a development-time audit tool, not a security boundary; the browser's CSP enforcement is the real protection.
Run it while developing to catch a missing `connectDomains` or `resourceDomains` entry before you ship, so the widget does not silently fail to load a font or reach an API.

Call it with your HTML and the policy you intend to send behind a development-only flag so it never runs in production. It returns a collection of warnings; an empty result means every external reference in that HTML is covered. It does not cover resources that those references load in turn, such as the font files a linked stylesheet requests.

```typescript

const html = `
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Roboto">
  <div style="font-family: Roboto, sans-serif;">Validation demo</div>`;

const policy = {
  connectDomains: [],
  resourceDomains: ["'self'", 'data:'],
  frameDomains: [],
  baseUriDomains: [],
};

// Run the audit only in development so it never executes in production.
const isDevelopment = process.env.NODE_ENV !== 'production';
if (isDevelopment) {
  const warnings = validateSecurityPolicy(html, policy);
  for (const w of warnings) {
    console.warn(`${w.source}: ${w.url} is not in ${w.policyField}`);
  }
}

// In development the audit above would warn that this HTML loads the Roboto
// stylesheet from fonts.googleapis.com, so you would add that origin to resourceDomains.
// The audit scans only the HTML you pass it and cannot see inside the stylesheet,
// so it does not report the font files that stylesheet loads from fonts.gstatic.com.
// Add that origin to resourceDomains yourself.
```

Each warning tells you the offending URL, where it appeared in the HTML, and which policy field it belongs in, so you can correct the policy before sending.
