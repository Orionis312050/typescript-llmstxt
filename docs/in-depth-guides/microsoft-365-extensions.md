---
title: "Microsoft 365 Extensions"
sidebar_position: "8"
---

# Microsoft 365 Extensions

Microsoft 365 Extensions let you extend an existing Teams SDK application to channels such as Direct Line and Email without rewriting any Teams-specific capabilities.

Register your existing Teams SDK handlers on the `App` returned by `useTeamsSdk` (TypeScript), `use_teams_sdk` (Python), or `AddTeamsSdk` (C#).
Microsoft 365 Extensions work with your existing Teams SDK configuration and handlers.
This app is a bridge where Microsoft Agents SDK `AgentApplication` acts the multichannel host.

Matching Teams activities continue through your Teams SDK handlers, while non-Teams activities and unmatched Teams activities flow through the Agents SDK application.

## How to Use It

Install the Microsoft 365 Extensions package alongside the Microsoft Agents SDK:

```bash
npm install @microsoft/agents-hosting @microsoft/teams.m365extensions
```

Add an Agents SDK host to your existing Teams SDK application, then replace your current `App` initialization with `useTeamsSdk`:

```typescript

  AgentApplication,
  CloudAdapter,
  getAuthConfigWithDefaults,
  MemoryStorage,
  type TurnState,
} from '@microsoft/agents-hosting';

const authConfig = getAuthConfigWithDefaults({
  clientId: process.env.CLIENT_ID,
  clientSecret: process.env.CLIENT_SECRET,
  tenantId: process.env.TENANT_ID,
});

const adapter = new CloudAdapter(authConfig);
const agentApp = new AgentApplication({
  storage: new MemoryStorage(),
  adapter,
});

const app = useTeamsSdk(agentApp, adapter.connectionManager);
```

## Key Benefits

### Extend Your Teams App to Other Channels

Through this extension, you can connect your existing Teams SDK app to other channels such as Direct Line and Email without rewriting your Teams-specific logic.

| Activity | Route |
| --- | --- |
| Teams activity matching a Teams SDK handler | Teams SDK |
| Teams activity without a matching Teams SDK handler | Agents SDK |
| Direct Line / Direct Line activity | Agents SDK |
| Email activity | Agents SDK |

#### Example Message Handler
You can use the Agents SDK message handler to identify the incoming channel and confirm how each activity was routed.
The incoming activity is recognized as Teams if the `channelId="msteams"`.

```typescript
AGENT_SDK_APP.onMessage(command('channel'), async (context: TurnContext) => {
  const via = isTeamsChannel(context.activity)
    ? 'Teams turn with no matching teams.ts route → fell through'
    : 'non-Teams channel → passed straight through';
  await context.sendActivity(`[Agent SDK] channelId=${context.activity.channelId} (${via})`);
});
```

#### Direct Line

![The channel command in Direct Line returning an Agents SDK response that identifies Direct Line as a non-Teams channel that passed straight through.](/screenshots/m365-extensions-web-chat.png)

#### Email

![The channel command sent through Email returning an Agents SDK response that identifies Email as a non-Teams channel that passed straight through.](/screenshots/m365-extensions-email.png)

### Access Rich Teams Capabilities from Agents SDK

When an Agents SDK handler receives a Teams activity, it can use the Teams SDK API client to access Teams-specific capabilities, such as message reactions:

```typescript

agentApp.onMessage(/^agents sdk react$/i, async (context) => {
  if (!isTeamsChannel(context.activity)) {
    await context.sendActivity('Message reactions are only available in Teams.');
    return;
  }

  const response = await context.sendActivity('Adding then removing a reaction.');
  const api = app.api.fromServiceUrl({
    serviceUrl: context.activity.serviceUrl!,
  });
  const conversationId = context.activity.conversation!.id;

  await api.conversations.addReaction(conversationId, response!.id!, 'like');
  await new Promise((resolve) => setTimeout(resolve, 2000));
  await api.conversations.deleteReaction(conversationId, response!.id!, 'like');
});
```

For a full example, the [TypeScript Microsoft 365 Extensions sample](https://github.com/microsoft/teams.ts/tree/main/examples/m365extensions) demonstrates a single application running across Teams, Web Chat, and Email.

## How It Works

Microsoft 365 Extensions adds a lightweight middleware bridge to your existing Teams SDK application. Aside from app initialization, your Teams-specific code stays the same.

1. **Add the bridge.** Call `useTeamsSdk` (TypeScript), `use_teams_sdk` (Python), or `AddTeamsSdk` (C#) to connect your app to Agents SDK.
2. **Keep your Teams logic.** Existing handlers and API client calls continue handling Teams activities.
3. **Add more channels.** Direct Line, Email, and unmatched Teams activities flow to the Agents SDK.
