---
title: "Creating Dialogs"
sidebar_position: "1"
---

# Creating Dialogs

:::tip
If you're not familiar with how to build Adaptive Cards, check out [the cards guide](../adaptive-cards.md). Understanding their basics is a prerequisite for this guide.
:::

## Entry Point

To open a dialog, add a button to your Adaptive Card using `OpenDialogData`. This sets up the `task/fetch` protocol and includes a `dialog_id` that the SDK uses to route to the correct handler.

```typescript

  AdaptiveCard,
  IAdaptiveCard,
  OpenDialogData,
  SubmitAction,
} from '@microsoft/teams.cards';
// ...

app.on('message', async ({ send }) => {
  await send({ type: 'typing' });

  const card: IAdaptiveCard = new AdaptiveCard({
    type: 'TextBlock',
    text: 'Select the examples you want to see!',
    size: 'Large',
    weight: 'Bolder',
  }).withActions(
    // OpenDialogData sets msteams.type = "task/fetch" and adds dialog_id for routing
    new SubmitAction()
      .withTitle('Simple form test')
      .withData(new OpenDialogData('simple_form')),
    new SubmitAction()
      .withTitle('Webpage Dialog')
      .withData(new OpenDialogData('webpage_dialog')),
    new SubmitAction()
      .withTitle('Multi-step Form')
      .withData(new OpenDialogData('multi_step_form'))
  );

  await send(new MessageActivityInput('Enter this form').addCard('adaptive', card));
});
```

## Handling Dialog Open Events

When a user clicks the button, Teams sends a `task/fetch` invoke to your app. Register a handler using `dialog.open.<dialog_id>` to handle a specific dialog, or `dialog.open` for a catch-all.

:::tip
Use sub-routes like `dialog.open.simple_form` to handle specific dialogs directly, instead of a single catch-all handler with if-else logic. This keeps each handler focused and avoids routing boilerplate.
:::

```typescript

// ...

// Handle a specific dialog by ID — no if-else needed
app.on('dialog.open.simple_form', async ({ activity }) => {
  const card: IAdaptiveCard = new AdaptiveCard()...

  return {
    task: {
      type: 'continue',
      value: {
        title: 'Title of Dialog',
        card: cardAttachment('adaptive', card),
      },
    },
  };
});
```

### Rendering A Card

You can render an Adaptive Card in a dialog by returning a card response.

```typescript

// ...

app.on('dialog.open.simple_form', async () => {
  const dialogCard = new AdaptiveCard(
    {
      type: 'TextBlock',
      text: 'This is a simple form',
      size: 'Large',
      weight: 'Bolder',
    },
    new TextInput()
      .withLabel('Name')
      .withIsRequired()
      .withId('name')
      .withPlaceholder('Enter your name')
  )
    // Use SubmitData to set the "action" field, which routes to dialog.submit.<action>
    .withActions(
      new SubmitAction().withTitle('Submit').withData(new SubmitData('simple_form'))
    );

  return {
    task: {
      type: 'continue',
      value: {
        title: 'Simple Form Dialog',
        card: cardAttachment('adaptive', dialogCard),
      },
    },
  };
});
```

:::info
The action type for submitting a dialog must be `Action.Submit`. This is a requirement of the Teams client. If you use a different action type, the dialog will not be submitted and the agent will not receive the submission event.
:::

### Rendering A Webpage

You can render a webpage in a dialog as well. There are some security requirements to be aware of:

1. The webpage must be hosted on a domain that is allow-listed as `validDomains` in the Teams app [manifest](../../teams/teams-manifest.md) for the agent
2. The webpage must also host the [teams-js client library](https://www.npmjs.com/package/@microsoft/teams-js). The reason for this is that for security purposes, the Teams client will not render arbitrary webpages. As such, the webpage must explicitly opt-in to being rendered in the Teams client. Setting up the teams-js client library handles this for you.

```typescript

// ...

app.on('dialog.open.webpage_dialog', async () => {
  return {
    task: {
      type: 'continue',
      value: {
        title: 'Webpage Dialog',
        // The webpage must be publicly accessible, use the teams-js client library,
        // and be registered in validDomains in the manifest.
        url: `${process.env['BOT_ENDPOINT']}/tabs/dialog-form`,
        width: 1000,
        height: 800,
      },
    },
  };
});
```

### Setting up Embedded Web Content

To serve web content for dialogs, you can use the `tab` method to host static webpages:

```typescript

// In your app setup (e.g., index.ts)
// Hosts a static webpage at /tabs/dialog-form
app.tab('dialog-form', path.join(__dirname, 'views', 'customform'));
```
