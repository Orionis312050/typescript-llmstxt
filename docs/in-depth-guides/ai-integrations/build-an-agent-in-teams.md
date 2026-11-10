---
title: "Build an agent in Teams"
sidebar_position: "1"
---

# Build an agent in Teams

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
In .NET, use SDK 2.1 for AI integrations. Keep `Microsoft.Teams.Apps` focused on Teams activity routing/streaming, and compose AI behavior with external agent/model tooling.
:::

This guide walks through building a Teams agent with the [OpenAI SDK](https://github.com/openai/openai-node) against Azure OpenAI. The TypeScript SDK stays agnostic about the intelligence layer — you bring the model client and the tool-call loop, and the Teams SDK handles activity routing, streaming, and Teams-native affordances like Adaptive Cards and feedback controls.

The agent loop here is driven by the OpenAI SDK's `runTools()` helper, which auto-executes each tool's `function` callback and feeds the result back to the model until it produces final text — so you don't hand-roll the tool-dispatch loop yourself.

:::note
This sample is bound to the OpenAI chat-completions wire protocol — Azure OpenAI and vanilla OpenAI both work; non-OpenAI providers do not.
:::

Full source: [examples/ai-mcp](https://github.com/microsoft/teams.ts/tree/main/examples/ai-mcp).

## Defining the agent

An agent is composed of three core elements: a **client** (model backend), **instructions** (system prompt), and **tools** (capabilities beyond text generation). A minimal setup starts with just a chat-enabled agent:

The "agent" is the model client plus a system prompt. The OpenAI SDK's `AzureOpenAI` client is the model backend, and `runTools()` (shown below) is the loop that drives instructions and tools together.

```typescript

const client = new AzureOpenAI({
  endpoint: process.env.AZURE_OPENAI_ENDPOINT!,
  apiKey: process.env.AZURE_OPENAI_API_KEY!,
  deployment: process.env.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME!,
  apiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-10-21',
});

const SYSTEM_PROMPT = 'You are a helpful Teams assistant.';
```

## Adding a local tool

Tools extend the agent with executable capabilities. They are regular functions the model can decide to invoke. Anything that runs in your process — database lookups, business logic, or Teams-specific actions like attaching an Adaptive Card to the reply — belongs here.

A good example is **clarification**: when a request is ambiguous, the agent asks the user to pick between interpretations instead of guessing. The tool builds an Adaptive Card and stashes it in a per-turn bucket the handler inspects after the run completes; the user's choice comes back as the next turn.

Tools are declared as `RunnableToolFunction`s — the OpenAI SDK runs each tool's `function` callback during the tool loop. The callback pushes the card into a per-turn bucket the handler inspects after the run completes, and returns a short placeholder string.

```typescript

type ClarificationArgs = { question: string; options: string[] };

export const CLARIFICATION_VERB = 'clarification';
export const CLARIFICATION_INPUT_ID = 'clarificationChoice';

function buildClarificationTool(pendingCards: AdaptiveCard[]): RunnableToolFunction {
  return {
    type: 'function',
    function: {
      name: 'request_clarification',
      description: 'Show an Adaptive Card asking the user to clarify their request when ambiguous.',
      parameters: {
        type: 'object',
        properties: {
          question: { type: 'string', description: 'The clarification question to ask the user.' },
          options: {
            type: 'array',
            items: { type: 'string' },
            description: '2-4 candidate interpretations the user can pick between.',
          },
        },
        required: ['question', 'options'],
        additionalProperties: false,
      },
      function: async (args: ClarificationArgs) => {
        pendingCards.push(buildClarificationCard(args));
        return 'Clarification card attached.';
      },
      parse: (raw: string) => JSON.parse(raw) as ClarificationArgs,
    },
  };
}

function buildClarificationCard(args: ClarificationArgs): AdaptiveCard {
  return new AdaptiveCard(
    new TextBlock(args.question, { weight: 'Bolder', size: 'Medium', wrap: true }),
    new ChoiceSetInput(...args.options.map((opt) => ({ title: opt, value: opt })))
      .withId(CLARIFICATION_INPUT_ID)
      .withIsRequired(true)
  ).withActions(
    new ExecuteAction({ title: 'Submit' })
      .withData(new SubmitData(CLARIFICATION_VERB))
      .withAssociatedInputs('auto')
  );
}
```

See [clarification cards](./enhance-the-teams-experience.md#clarification-cards) for how the user's choice flows back in.

<img
  src=ClarificationCardImgUrl
  alt="Screenshot of a clarification Adaptive Card in a Teams chat, asking the user to pick between candidate interpretations of an ambiguous question."
  style={ width: '100%', maxWidth: 700 }
/>

## Adding remote MCP tools

Remote tools are exposed via [MCP](https://modelcontextprotocol.io/introduction) servers and live behind a network boundary. The agent discovers their schemas at runtime and invokes them over HTTP. From the model's perspective, they behave like any other tool.

Connect to the MCP server once at startup, list its tools, and wrap each one as a `RunnableToolFunction`. The callback invokes the server and returns the result text to the model.

```typescript

const client = new Client({ name: 'ai-mcp-sample', version: '0.0.0' });
await client.connect(new StreamableHTTPClientTransport(new URL('https://learn.microsoft.com/api/mcp')));

const { tools } = await client.listTools();

const mcpTools: RunnableToolFunction>[] = tools.map((tool) => ({
  type: 'function',
  function: {
    name: tool.name,
    description: tool.description ?? '',
    parameters: (tool.inputSchema as Record<string, unknown>) ?? { type: 'object' },
    function: async (args: Record<string, unknown>) => {
      const result = await client.callTool({ name: tool.name, arguments: args });
      return stringifyResult(result.content);
    },
    parse: (raw: string) => JSON.parse(raw) as Record<string, unknown>,
  },
}));
```

## Running the agent in Teams

Integrate with Teams by forwarding incoming messages to the agent and streaming the response back to the chat interface chunk by chunk.

`runTools()` sends the request with your tool definitions, auto-invokes any tool the model calls, re-prompts with the result, and repeats until the model produces final text. `content` events fire for each text delta — forward them straight to the Teams stream.

```typescript
const runner = client.chat.completions.runTools({
  model: deployment,
  messages: history,
  tools: [clarificationTool, ...mcpTools],
  stream: true,
});

runner.on('content', (delta: string) => stream.emit(delta));
await runner.done();
```

## Per-conversation memory

By default, each run starts with no history — the model only sees the current message. This works for one-shot interactions, but is insufficient for multi-turn conversations where users refer back to earlier context. Keep a per-conversation buffer and reuse it across turns:

Keep one `ChatCompletionMessageParam[]` per Teams conversation. After each run, sync the runner's view — it includes the system, user, and every tool-call / tool-result / assistant message added during the loop — back into your map so the next turn sees the full prior context.

```typescript

const histories = new Map<string, ChatCompletionMessageParam[]>();

function getOrCreateHistory(convId: string): ChatCompletionMessageParam[] {
  let history = histories.get(convId);
  if (!history) {
    history = [{ role: 'system', content: SYSTEM_PROMPT }];
    histories.set(convId, history);
  }
  return history;
}

// after runner.done():
const ran = runner.messages as ChatCompletionMessageParam[];
history.splice(0, history.length, ...ran);
```

In production, push conversation history into Redis, Cosmos DB, or whatever you already use for state.

## Grounding responses with citations

When a tool returns search results, you usually want the model to cite its sources. The pattern: intercept each tool result, assign every source a stable 1-based index, and hand that index back to the model so it can reference it inline as `[1]`, `[2]`, and so on. The collected citations are attached to the final reply in [Enhancing the Teams Experience](./enhance-the-teams-experience.md#citations).

The extraction lives in a small `CitationCollector`. Each MCP tool callback feeds its raw result into `tryExtract`, which parses the search payload and assigns every source a stable 1-based position. The same collector instance is captured by every tool call on a turn.

```typescript
type CitationEntry = { position: number; url: string; title: string; snippet: string };

export class CitationCollector {
  private readonly entries = new Map<string, CitationEntry>();

  tryExtract(result: string): void {
    let doc: any;
    try {
      doc = JSON.parse(result);
    } catch {
      return; // non-JSON tool results are ignored
    }
    for (const item of doc?.results ?? []) {
      const url = item.contentUrl ?? item.link;
      if (!url || this.entries.has(url)) continue;
      const snippet = (item.content ?? item.description ?? '').slice(0, 160);
      this.entries.set(url, {
        position: this.entries.size + 1,
        url,
        title: item.title ?? '',
        snippet,
      });
    }
  }
}
```

To wire it in, call `citations.tryExtract(text)` inside each MCP tool's callback before returning the result. The collected entries are attached to the final reply in [Enhancing the Teams Experience](./enhance-the-teams-experience.md#citations).

For Teams-specific enhancements — continue to [Enhancing the Teams Experience](./enhance-the-teams-experience.md).
