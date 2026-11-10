---
title: "AI Integrations"
sidebar_position: "1"
---

# AI Integrations

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
For .NET, AI-specific libraries and older patterns from SDK 2.0 are deprecated. In SDK 2.1, use `Microsoft.Teams.Apps` as the Teams transport/runtime layer and integrate your preferred AI stack (for example Microsoft Agent Framework, OpenAI SDK, MCP, and A2A) in your handlers.
:::

Microsoft Teams SDK provides the platform and conversational interface for your app while remaining agnostic to the underlying intelligence. You can choose any AI framework, model, or protocol that suits your scenario and integrate it into your message handlers. The samples below walk through a few common ways to do that, from a single bot reasoning with an agent to multiple bots collaborating with each other.

- **[Build an agent in Teams](./ai-integrations/build-an-agent-in-teams.md)** — create an agent with the [OpenAI SDK](https://github.com/openai/openai-node) and Azure OpenAI, add a local clarification tool and remote MCP tool servers, stream responses into Teams, and preserve conversation history across turns.
- **[Enhancing the Teams Experience](./ai-integrations/enhance-the-teams-experience.md)** — build on the base integration with richer conversational features: clarification cards, suggested follow-up prompts, inline citations, and structured feedback handling.
- **[Exposing Teams to AI Agents (MCP)](./ai-integrations/exposing-teams-to-ai-agents-mcp.md)** — turn your bot into an [MCP](https://modelcontextprotocol.io/introduction) server so external agents can reach real users through Teams chat with tools like `find_user`, `notify`, `ask`, and `request_approval`. Useful for human-in-the-loop workflows.
- **[Bot-to-Bot with A2A](./ai-integrations/bot-to-bot-communication-with-a2a.md)** — two Teams bots, each backed by its own agent, hand users off to each other over the [Agent2Agent](https://a2a-protocol.org/) protocol, opening a proactive chat with the user so the conversation continues seamlessly.

All samples are available in the [`microsoft/teams.ts` examples](https://github.com/microsoft/teams.ts/tree/main/examples).
