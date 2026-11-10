---
title: "Bot-to-Bot Communication with A2A"
sidebar_position: "4"
---

# Bot-to-Bot Communication with A2A

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
For .NET, implement A2A with SDK 2.1 plus an external A2A protocol library. Use Teams SDK handlers for inbound user turns and proactive messaging, and use A2A messages for bot-to-bot handoff coordination.
:::

Agents are typically designed to interact either with people (chatbots) or with systems (tools, APIs, MCP servers). [Agent2Agent](https://a2a-protocol.org/) (A2A) introduces a third interaction model: agents communicating directly with other agents as peers — each with its own model, capabilities, and human audience.

This guide walks through a **handoff** between two Teams bots, **Alice** and **Bob**, each backed by its own LLM agent. A user DMs one bot; its agent reads the peer's capability description and decides whether to answer directly or hand the user off. On handoff, the receiving bot **proactively opens a 1:1 chat** with the user and greets them with the context that came across — so the conversation continues seamlessly in the new chat.

Both bots run the **same code**, differentiated entirely by environment variables (name, description, self/peer URLs). They use [`@a2a-js/sdk`](https://www.npmjs.com/package/@a2a-js/sdk) for the protocol and the OpenAI SDK for the LLM agent.

Full source: [examples/a2a](https://github.com/microsoft/teams.ts/tree/main/examples/a2a).

## Advertising capabilities with an Agent Card

Every A2A server publishes an `AgentCard` — a small machine-readable document describing who the agent is and what it can do. Peers fetch this card to learn about each other; their LLMs then read the `description` field to decide *when* to hand off a user.

```typescript

function buildAgentCard(config: Config): AgentCard {
  return {
    name: config.name,
    description: config.description,
    url: `${config.selfUrl.replace(/\/+$/, '')}/a2a`,
    version: '1.0.0',
    protocolVersion: '0.3.0',
    preferredTransport: 'JSONRPC',
    capabilities: {},
    defaultInputModes: ['application/json'],
    defaultOutputModes: ['text/plain'],
    skills: [{
      id: 'handoff',
      name: 'Handoff',
      description: `Accepts handoffs of users from peer bots. Specialty: ${config.description}`,
      tags: ['a2a', 'teams', 'handoff'],
    }],
  };
}
```

The `description` is the most important knob in this sample — it's the natural-language summary another bot's LLM uses to decide whether *this* bot is the right peer for a given user. Tweak it to match the persona and expertise you want each bot to advertise.

## The handoff message contract

A handoff carries everything the receiving bot needs to reach the user proactively: their **`aadObjectId`** (the tenant-wide identity both bots share — the Teams MRI one bot sees isn't valid against the other), the **`tenantId`**, the **`serviceUrl`**, and a **`summary`** of the conversation so the peer can pick up cold.

```typescript
export type HandoffMessage = {
  readonly kind: 'handoff';
  readonly from: string;
  readonly userName: string;
  readonly aadObjectId: string;
  readonly tenantId: string;
  readonly serviceUrl: string;
  readonly summary: string;
};

export function isHandoffMessage(value: unknown): value is HandoffMessage {
  if (!value || typeof value !== 'object') return false;
  const v = value as Record<string, unknown>;
  return (
    v.kind === 'handoff' &&
    typeof v.aadObjectId === 'string' && v.aadObjectId.length > 0 &&
    typeof v.tenantId === 'string' && v.tenantId.length > 0 &&
    typeof v.serviceUrl === 'string' && v.serviceUrl.length > 0 &&
    typeof v.summary === 'string'
  );
}
```

A type guard validates the inbound `DataPart` before the receiving bot acts on it.

## LLM-driven handoff

Routing is not a hard-coded rule — the LLM decides. Each bot exposes a single `handoff_to_peer` tool to its agent, and the agent's instructions include the live `AgentCard.description` of the peer. When a question fits the peer's expertise better than its own, the model calls the tool.

```typescript

function buildHandoffTool(): RunnableToolFunction<{ summary: string }> {
  return {
    type: 'function',
    function: {
      name: 'handoff_to_peer',
      description:
        'Hand off the current user to your peer when their expertise is a better fit. ' +
        'Pass a concise summary so the peer can pick up cold. The peer will message the user directly.',
      parameters: {
        type: 'object',
        properties: { summary: { type: 'string' } },
        required: ['summary'],
        additionalProperties: false,
      },
      function: async (args: { summary: string }) => {
        const identity = TURN_STORAGE.getStore();
        if (!identity) {
          // Called from a handoff greeting (no identity) — guard against ping-pong.
          return 'handoff_to_peer is unavailable in this context.';
        }
        const payload: HandoffMessage = { kind: 'handoff', from: config.name, ...identity, summary: args.summary };
        await a2aClient.sendHandoff(payload);
        return 'Handoff confirmed. The peer will message the user directly.';
      },
      parse: (raw: string) => JSON.parse(raw) as { summary: string },
    },
  };
}
```

The identity is held in an `AsyncLocalStorage` (`TURN_STORAGE`) for the duration of the turn. The agent's system prompt embeds the peer's live `AgentCard.description` so the model knows what the peer specializes in:

```typescript
const instructions = [
  `You are ${config.name}, a Teams bot. Your specialty: ${config.description}.`,
  'You have one peer:',
  `- ${config.peerName}: ${peerDescription}`,
  `- If the user's question fits ${config.peerName}'s specialty better than your own, call handoff_to_peer with a clear summary. Then briefly tell the user you're handing them over.`,
  '- Otherwise, answer directly.',
].join('\n');
```

The identity needed to build the handoff is captured from the inbound Teams activity and stashed for the duration of the turn, so the tool callback can reach it without threading it through every call. A handoff *greeting* runs with no identity set — the tool guards against that to prevent a ping-pong.

## Sending a handoff over A2A

The outbound side resolves the peer's `AgentCard` once (so the agent can read its live description into the tool), then ships the handoff as a `DataPart`.

```typescript

export class A2APeerClient {
  private cachedClient?: Client;

  async sendHandoff(payload: HandoffMessage): Promise<void> {
    if (!this.cachedClient) await this.getPeerCard();
    const params: MessageSendParams = {
      message: {
        kind: 'message',
        role: 'user',
        messageId: crypto.randomUUID(),
        parts: [{ kind: 'data', data: payload as unknown as Record<string, unknown> }],
      },
    };
    await this.cachedClient!.sendMessage(params);
  }

  // getPeerCard() resolves the peer's AgentCard once via the well-known endpoint
  // and constructs the underlying A2A client, then caches both.
}
```

## Receiving a handoff

The inbound side implements the A2A protocol's executor interface. For each inbound message it pulls the handoff out of the `DataPart`, opens a fresh 1:1 with the user against **their** `serviceUrl`, asks the agent to seed that conversation's history with the handoff context and produce a greeting, sends the greeting proactively, and acks back so the sender's call resolves.

```typescript

export class HandoffAgentExecutor implements AgentExecutor {
  execute = async (ctx: RequestContext, bus: ExecutionEventBus): Promise<void> => {
    const handoff = this.extractHandoff(ctx);
    if (!handoff) {
      this.publishText(bus, ctx, 'Unsupported or incomplete handoff message.');
      bus.finished();
      return;
    }
    // 1. Open a 1:1 with the user against THEIR serviceUrl.
    const newConvId = await this.openDmWithUser(handoff);
    // 2. Seed history with the handoff context + greeting.
    const greeting = await this.agent.greetWithHandoff(newConvId, handoff);
    // 3. Send the greeting proactively.
    await this.app.send(newConvId, greeting);
    // 4. Ack so the sender's sendMessage resolves.
    this.publishText(bus, ctx, `Handoff received and ${handoff.userName} contacted directly.`);
    bus.finished();
  };

  private async openDmWithUser(handoff: HandoffMessage): Promise<string> {
    const api = new TeamsApiClient(handoff.serviceUrl, this.app.api.http);
    const conv = await api.conversations.create({
      tenantId: handoff.tenantId,
      members: [{ id: handoff.aadObjectId, role: 'user', name: handoff.userName }],
    });
    if (!conv.id) throw new Error('CreateConversation returned no id.');
    return conv.id;
  }
}
```

`greetWithHandoff` runs the LLM with the handoff summary as a synthetic turn and leaves it in the per-conversation history, so subsequent user replies continue naturally.

Because the greeting turn is left in the per-conversation history, when the user replies in their new DM the agent picks up coherently.

## Wiring A2A into your Teams app

The Teams bot and A2A server run in the same process and share one HTTP surface: `/api/messages` for Teams, `/a2a` for inbound handoffs, and `/.well-known/agent-card.json` for the AgentCard.

```typescript

const expressApp = express();
const app = new App({ httpServerAdapter: new ExpressAdapter(expressApp) });

const a2aHandler = new DefaultRequestHandler(
  buildAgentCard(config),
  new InMemoryTaskStore(),
  new HandoffAgentExecutor(app, agent, config, log)
);
expressApp.use('/.well-known/agent-card.json', agentCardHandler({ agentCardProvider: a2aHandler }));
expressApp.use('/a2a', jsonRpcHandler({ requestHandler: a2aHandler, userBuilder: UserBuilder.noAuthentication }));

// Register /api/messages without starting an internal server — we own the http.Server.
await app.initialize();
http.createServer(expressApp).listen(Number(process.env.PORT) || 3978);
```

Each bot needs its own Teams app registration (so DMs route to the right bot) and its own port. The sample runs Alice on `3978` and Bob on `3979`; their peer URLs point at each other.

## Putting it all together

With both bots running and installed for the user, DM Alice with a question outside her specialty and watch the round-trip: Alice's LLM calls `handoff_to_peer`, Bob receives it over A2A, opens a new 1:1 with the user, and greets them with an answer already in hand. The bots are symmetric — the same flow runs the other way from Bob to Alice.

<img
  src=A2AImgUrl
  alt="Animated screenshot of the end-to-end A2A handoff: a user DMs Alice, Alice hands off to Bob, and Bob opens a new chat greeting the user with context."
  style={ width: '100%', maxWidth: 900 }
/>

:::warning
This sample configures no authenticator on the A2A endpoint, so any caller can post a handoff. For production, validate the caller's identity (a bearer token or mTLS) before opening a conversation with someone they named.
:::
