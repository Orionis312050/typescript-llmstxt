---
title: "🔒 User Authentication"
sidebar_position: "7"
---

# 🔒 User Authentication

At times agents must access secured online resources on behalf of the user, such as checking email, checking on flight status, or placing an order. To enable this, the user must authenticate their identity and grant consent for the application to access these resources. This process results in the application receiving a token, which the application can then use to access the permitted resources on the user's behalf.

:::info
This is an advanced guide. It is highly recommended that you are familiar with [Teams Core Concepts](../teams/teams-core-concepts.md) before attempting this guide.
:::

:::warning
User authentication does not work with the developer tools setup. You have to run the app in Teams. Follow [Quickstart: Register your app](../get-started/quickstart-register-your-app.md) to register and sideload your bot.
:::

:::info
It is possible to authenticate the user into [other auth providers](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-identity-providers?view=azure-bot-service-4.0&tabs=adv2%2Cga2#other-identity-providers) like Facebook, Github, Google, Dropbox, and so on.
:::

## Project Setup

:::tip
If you're creating a new app, use the `graph` template. Skip this if you're adding auth to an existing app.
:::

Use your terminal to run the following command:

```sh
teams project new typescript oauth-app --template graph
```

This command:

1. Creates a new directory called `oauth-app`.
2. Bootstraps the graph agent template files into it under `oauth-app/src`.

### Set up the OAuth connection

User authentication requires an **Azure-managed bot** (Teams-managed bots don't support OAuth connections). If you registered with `--teams-managed`, migrate first:

```sh
teams app bot migrate <appId> --subscription <id> --resource-group <your-resource-group>
```

Then follow the [User Authentication Setup guide](/cli/guides/user-authentication-setup) to configure the AAD app, create the Azure Bot OAuth connection, and update the manifest. The guide covers both SSO (silent token exchange) and generic OAuth.

:::tip
If you'd rather have an AI coding assistant run the setup, install the [`teams-dev` skill](../developer-tools/agent-skills.md) and ask it to "set up SSO for my Teams bot".
:::

## Configure the OAuth connection

**SDK 2.0 (Legacy):**

```ts

const app = new App({
  oauth: {
    // The name of the auth connection to use.
    // It should be the same as the OAuth connection name defined in the Azure Bot configuration.
    defaultConnectionName: 'graph',
  },
});
```

**SDK 2.1 (current):**

Register the connections you plan to use. `addOAuthFlow` returns the object that owns one.

```ts

const app = new App();

// Or register up front: new App({ oauthFlows: ['graph'] })
const graph = app.addOAuthFlow('graph', {
  oauthCardText: 'Sign in to your account',
  signInButtonText: 'Sign in',
});
```

Both card options are optional. Use `app.getOAuthFlow('graph')` to get the flow again from another module.

:::note
Registering a flow enables per-turn state unless you set `state` yourself, so sign-in callbacks that don't name a connection can be traced back to the one that started them. `oauth.defaultConnectionName` can't be combined with registered flows.
:::

:::tip
Make sure you use the same name you used when creating the OAuth connection in the Azure Bot Service resource.
:::

:::note
In many templates, `graph` is the default name of the OAuth connection, but you can register the connection under any name — it just has to match the OAuth connection configured on the Azure Bot.
:::

## Signing In

Start sign-in from inside a route handler, for example when receiving the `/signin` message:

:::note
This uses the Single Sign-On (SSO) authentication flow. To learn more about all the available flows and their differences see the [official documentation](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-authentication?view=azure-bot-service-4.0).
:::

**SDK 2.0 (Legacy):**

```ts
app.message('/signin', async ({ signin, send }) => {
  if (await signin()) {
    await send('you are already signed in!');
  }
});
```

**SDK 2.1 (current):**

```ts
app.message('/signin', async (ctx) => {
  const token = await graph.signIn(ctx);
  if (token) {
    await ctx.send('you are already signed in!');
  }
});
```

## Handling sign-in completion

The sign-in method returns a token when the user already has one, and otherwise sends an OAuth card and returns nothing. The flow finishes on a later turn, so completion is delivered through a callback.

**SDK 2.0 (Legacy):**

```ts
app.event('signin', async ({ send, token }) => {
  await send(
    `Signed in using OAuth connection ${token.connectionName}. Please type **/whoami** to see your profile or **/signout** to sign out.`
  );
});
```

**SDK 2.1 (current):**

The callback is scoped to its connection, so there's no need to branch on the connection name.

```ts
graph.onSignInComplete(async (ctx, token) => {
  await ctx.send(
    `Signed in on ${token.connectionName}. Please type **/whoami** to see your profile or **/signout** to sign out.`
  );
});
```

:::note
Each flow holds a single completion callback — registering again replaces the previous one. The app-wide `signin` event still fires for every connection if you need one place to observe them all.
:::

## Start using the client

:::note
The default graph configuration requests the `User.ReadBasic.All` permission. It is possible to request other permissions by modifying the App Registration for the bot on Azure.
:::

From this point, you can query graph for the signed-in user, for example to reply to the `/whoami` message, or in any other route.

**SDK 2.0 (Legacy):**

Use the `isSignedIn` flag and the `userGraph` client.

```ts

app.message('/whoami', async ({ send, userGraph, signin }) => {
  if (!await signin()) {
    return;
  }
  const me = await userGraph.call(endpoints.me.get);
  await send(
    `you are signed in as "${me.displayName}" and your email is "${me.mail || me.userPrincipalName}"`
  );
});

app.on('message', async ({ send, activity, signin }) => {
  if (await signin()) {
    await send(
      `You said: "${activity.text}". Please type **/whoami** to see your profile or **/signout** to sign out.`
    );
  } else {
    await send(`You said: "${activity.text}". Please type **/signin** to sign in.`);
  }
});
```

**SDK 2.1 (current):**

Build the client from the flow's token, so it always talks to the connection that owns it.

```ts

app.message('/whoami', async (ctx) => {
  const token = await graph.signIn(ctx);
  if (!token) return; // OAuth card sent — resumes on the callback turn

  const client = new GraphClient({ token: () => token }, { baseUrlRoot: app.graphBaseUrl });
  const me = await client.call(endpoints.me.get);
  await ctx.send(
    `you are signed in as "${me.displayName}" and your email is "${me.mail || me.userPrincipalName}"`
  );
});
```

## Signing Out

You can signout by calling the `signout` method, this will remove the token from the User Token service cache

**SDK 2.0 (Legacy):**

```ts
app.message('/signout', async ({ send, signout, isSignedIn }) => {
  if (!isSignedIn) return;
  await signout();
  await send('you have been signed out!');
});
```

**SDK 2.1 (current):**

```ts
app.message('/signout', async (ctx) => {
  await graph.signOut(ctx);
  await ctx.send('you have been signed out!');
});
```

## Using multiple OAuth connections

An agent may need more than one identity — for example Graph for the user's profile and GitHub for their repos. Each connection is registered separately, so tokens, callbacks and sign-out are scoped to the connection that owns them.

```ts
const app = new App();

const graph = app.addOAuthFlow('graph', { oauthCardText: 'Sign in with Microsoft' });
const github = app.addOAuthFlow('github', { oauthCardText: 'Sign in with GitHub' });

app.message('/graph', async (ctx) => {
  const token = await graph.signIn(ctx);
  if (token) await sendGraphProfile(ctx, token);
});

app.message('/github', async (ctx) => {
  const token = await github.signIn(ctx);
  if (token) await sendGitHubProfile(ctx, token);
});

graph.onSignInComplete(async (ctx, token) => sendGraphProfile(ctx, token.token));
github.onSignInComplete(async (ctx, token) => sendGitHubProfile(ctx, token.token));

app.message('/signout github', async (ctx) => {
  await github.signOut(ctx);
  await ctx.send('Signed out of GitHub.');
});
```

Signing out of one connection leaves the other signed in.

:::info
Each connection must exist on the Azure Bot resource under the name you register it with. Follow the [User Authentication Setup guide](/cli/guides/user-authentication-setup) once per connection.
:::

## Checking connection status

One call returns the status of every registered connection, so you don't have to probe them individually.

```ts
app.message('/status', async (ctx) => {
  const statuses = await ctx.getConnectionStatus();
  const lines = statuses.map(
    (status) => `- \`${status.connectionName}\`: ${status.hasToken ? 'signed in' : 'signed out'}`
  );
  await ctx.send(lines.join('\n'));
});
```

## Resuming Pending Messages After Sign-In

When a user isn't signed in and your message handler calls the sign-in method, an OAuth card is sent and the current turn ends. The sign-in completes on a *separate turn* — meaning the original message text is not available in the sign-in success context.

To avoid ignoring what the user originally asked, store the pending message before initiating sign-in, then retrieve and process it once sign-in succeeds:

**SDK 2.0 (Legacy):**

```ts
const pendingMessages = new Map<string, { text: string; activity: any }>();

app.on('message', async ({ signin, activity, send }) => {
  // signin() returns the token if already signed in, or undefined if OAuth card was sent
  const token = await signin({
    oauthCardText: 'To help with that, I need to sign you in first.',
  });

  if (!token) {
    // OAuth card sent — store the original message for later
    pendingMessages.set(activity.from.id, {
      text: activity.text,
      activity,
    });
    return;
  }

  // User is already signed in — process normally
  await processMessage(activity.text, { send });
});

app.event('signin', async ({ send, userGraph, activity }) => {
  const userId = activity.from.id;
  const pending = pendingMessages.get(userId);

  if (pending) {
    pendingMessages.delete(userId);
    await send('Successfully signed in! Processing your original request...');
    await processMessage(pending.text, { send, userGraph });
  } else {
    await send('You are now signed in!');
  }
});
```

**SDK 2.1 (current):**

```ts
const pendingMessages = new Map<string, string>();

app.on('message', async (ctx) => {
  // signIn() returns the token if already signed in, or undefined if an OAuth card was sent
  const token = await graph.signIn(ctx, {
    oauthCardText: 'To help with that, I need to sign you in first.',
  });

  if (!token) {
    pendingMessages.set(ctx.activity.from.id, ctx.activity.text);
    return;
  }

  await processMessage(ctx.activity.text, ctx, token);
});

graph.onSignInComplete(async (ctx, token) => {
  const userId = ctx.activity.from.id;
  const pending = pendingMessages.get(userId);

  if (!pending) {
    await ctx.send('You are now signed in!');
    return;
  }

  pendingMessages.delete(userId);
  await ctx.send('Successfully signed in! Processing your original request...');
  await processMessage(pending, ctx, token.token);
});
```

:::tip
For production apps, consider using a persistent store (database, Redis, etc.) instead of an in-memory map so pending messages survive restarts. You should also implement expiration or cleanup logic (e.g., a TTL) to discard stale entries when sign-in is cancelled, times out, or fails.
:::

## Handling Sign-In Failures

When using SSO, if the token exchange fails Teams sends a `signin/failure` invoke activity to your app. The SDK includes a built-in default handler that logs a warning with actionable troubleshooting guidance. You can optionally register your own handler to customize the behavior:

**SDK 2.0 (Legacy):**

```ts
app.on('signin.failure', async ({ activity, send }) => {
  const { code, message } = activity.value;
  console.log(`Sign-in failed: ${code} - ${message}`);
  await send('Sign-in failed.');
});
```

**SDK 2.1 (current):**

```ts
graph.onSignInFailure(async (ctx, failure) => {
  ctx.log.error(`Graph sign-in failed: ${failure?.code} - ${failure?.message}`);
  await ctx.send('Graph sign-in failed.');
});
```

`failure` is undefined for token-service and token-exchange failures, which carry no Teams payload.

:::tip
The most common failure codes are `installedappnotfound` (bot app not installed for the user) and `resourcematchfailed` (Token Exchange URL doesn't match the Application ID URI). See [SSO Setup - Troubleshooting](../teams/user-authentication/sso-setup.md#troubleshooting) for a full list of failure codes and troubleshooting steps.
:::

:::warning
A `signin/failure` activity doesn't name its connection. The SDK attributes it from the pending sign-in recorded when the flow started; if that can't be resolved, *every* registered connection's failure handler is notified. Check the connection first if your handler does something destructive.

Registering a flow also turns on per-turn state unless you configure `state` yourself, defaulting to local, in-process storage. Configure shared state storage before running more than one instance — otherwise pending sign-in attribution and token-exchange dedup don't cross processes, and sign-in fails intermittently.
:::

## Regional Configs
You may be building a regional bot that is deployed in a specific Azure region (such as West Europe, East US, etc.) rather than global. This is important for organizations that have data residency requirements or want to reduce latency by keeping data and authentication flows within a specific area.

These examples use West Europe, but follow the equivalent for other regions.

**Azure Portal:**

To configure a new regional bot in Azure, you must setup your resoures in the desired region. Your resource group must also be in the same region. 

1. Deploy a new App Registration in `westeurope`.
2. Deploy and link a new Enterprise Application (Service Principal) on Microsoft Entra in `westeurope`.
3. Deploy and link a new Azure Bot in `westeurope`.
4. In your App Registration, in the `Authentication (Preview)` tab, add a `Redirect URI` for the Platform Type `Web` to your regional endpoint (e.g., `https://europe.token.botframework.com/.auth/web/redirect`)

![Authentication Tab](/screenshots/regional-auth.png)

5. In your `.env` file (or wherever you set your environment variables), add your `OAUTH_URL`. For example:
`OAUTH_URL=https://europe.token.botframework.com`

**Agents Toolkit:**

To configure a new regional bot with ATK, you will need to make a few updates. Note that this assumes you have not yet deployed the bot previously.

1. In `azurebot.bicep`, replace all `global` occurrences to `westeurope`
2. In `manifest.json`, in `validDomains`, `*.botframework.com` should be replaced by `europe.token.botframework.com`
3. In `aad.manifest.json`, replace `https://token.botframework.com/.auth/web/redirect` with `https://europe.token.botframework.com/.auth/web/redirect`
4. In your `.env` file, add your `OAUTH_URL`. For example:
`OAUTH_URL=https://europe.token.botframework.com`

## Resources

[User Authentication Basics](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-authentication?view=azure-bot-service-4.0)
