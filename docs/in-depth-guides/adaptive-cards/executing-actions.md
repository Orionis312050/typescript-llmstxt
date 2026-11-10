---
title: "Executing Actions"
sidebar_position: "2"
---

# Executing Actions

Adaptive Cards support interactive elements through **actions**—buttons, links, and input submission triggers that respond to user interaction.
You can use these to collect form input, trigger workflows, show task modules, open URLs, and more.

## Action Types

The Teams SDK supports several action types for different interaction patterns:

| Action Type               | Purpose                | Description                                                                  |
| ------------------------- | ---------------------- | ---------------------------------------------------------------------------- |
| `Action.Execute`          | Server‑side processing | Send data to your bot for processing. Best for forms & multi‑step workflows. |
| `Action.Submit`           | Simple data submission | Legacy action type. Prefer `Execute` for new projects.                       |
| `Action.OpenUrl`          | External navigation    | Open a URL in the user's browser.                                            |
| `Action.ShowCard`         | Progressive disclosure | Display a nested card when clicked.                                          |
| `Action.ToggleVisibility` | UI state management    | Show/hide card elements dynamically.                                         |

:::info
For complete reference, see the [official documentation](https://adaptivecards.microsoft.com/?topic=Action.Execute).
:::

## Creating Actions with the SDK

### Single Actions

The SDK provides builder helpers that abstract the underlying JSON. For example:

```typescript

// ...

new ExecuteAction({ title: 'Submit Feedback' })
  .withData({ action: 'submit_feedback' })
  .withAssociatedInputs('auto'),
```

### Action Sets

Group actions together using `ActionSet`:

```typescript

// ...

new ActionSet(
  new ExecuteAction({ title: 'Submit Feedback' })
    .withData({ action: 'submit_feedback' })
    .withAssociatedInputs('auto'),
  new OpenUrlAction('https://adaptivecards.microsoft.com').withTitle('Learn More')
);
```

### Raw JSON Alternative

Just like when building cards, if you prefer to work with raw JSON, you can do just that. You get type safety for free in TypeScript.

```typescript

// ...

{
  type: 'Action.OpenUrl',
  url: 'https://adaptivecards.microsoft.com',
  title: 'Learn More',
} as const satisfies IOpenUrlAction
```

## Working with Input Values

### Associating data with the cards

Sometimes you want to send a card and have it be associated with some data. Set the `data` value to be sent back to the client so you can associate it with a particular entity.

```typescript

  AdaptiveCard,
  TextInput,
  ToggleInput,
  ActionSet,
  ExecuteAction,
} from '@microsoft/teams.cards';
// ...

function editProfileCard() {
  const card = new AdaptiveCard(
    new TextInput({ id: 'name' }).withLabel('Name').withValue('John Doe'),
    new TextInput({ id: 'email', label: 'Email', value: 'john@contoso.com' }),
    new ToggleInput('Subscribe to newsletter').withId('subscribe').withValue('false'),
    new ActionSet(
      new ExecuteAction({ title: 'Save' })
        .withData({
          action: 'save_profile',
          entityId: '12345', // This will come back once the user submits
        })
        .withAssociatedInputs('auto')
    )
  );

  // Data received in handler
  /**
  {
    action: "save_profile",
    entityId: "12345",     // From action data
    name: "John Doe",      // From name input
    email: "john@doe.com", // From email input
    subscribe: "true"      // From toggle input (as string)
  }
  */

  return card;
}
```

### Input Validation

Input Controls provide ways for you to validate. More details can be found on the Adaptive Cards [documentation](https://adaptivecards.microsoft.com/?topic=input-validation).

```typescript

  AdaptiveCard,
  NumberInput,
  TextInput,
  ActionSet,
  ExecuteAction,
} from '@microsoft/teams.cards';
// ...

function createProfileCardInputValidation() {
  const ageInput = new NumberInput({ id: 'age' })
    .withLabel('Age')
    .withIsRequired(true)
    .withMin(0)
    .withMax(120);

  const nameInput = new TextInput({ id: 'name' })
    .withLabel('Name')
    .withIsRequired()
    .withErrorMessage('Name is required!'); // Custom error messages
  const card = new AdaptiveCard(
    nameInput,
    ageInput,
    new TextInput({ id: 'location' }).withLabel('Location'),
    new ActionSet(
      new ExecuteAction({ title: 'Save' })
        .withData({
          action: 'save_profile',
        })
        .withAssociatedInputs('auto') // All inputs should be validated
    )
  );

  return card;
}
```

## Routing & Handlers

### Using SubmitData

The SDK provides a `SubmitData` helper that sets the routing key for your action. This is the recommended way to wire up actions to specific handlers:

```typescript

// ...

new ExecuteAction({ title: 'Submit Feedback' })
  .withData(new SubmitData('submit_feedback'))
  .withAssociatedInputs('auto')

// You can also pass extra static data alongside the action name
new ExecuteAction({ title: 'Save' })
  .withData(new SubmitData('save_profile', { entityId: '12345' }))
  .withAssociatedInputs('auto')
```

`SubmitData` sets a reserved `action` key in the card's data payload. When the user clicks the button, the SDK router reads this key to dispatch to the matching handler.

### Action-Specific Handlers

Register handlers for specific actions. When you use `SubmitData` to set the action name on the card, the SDK routes directly to the matching handler:

```typescript

// ...

// 'submit_feedback' matches the identifier passed to SubmitData('submit_feedback')
app.on('card.action.submit_feedback', async ({ activity, send }) => {
  const data = activity.value.action.data;
  await send(`Feedback received: ${data.feedback}`);
  return {
    statusCode: 200,
    type: 'application/vnd.microsoft.activity.message',
    value: 'Action processed successfully',
  };
});

app.on('card.action.save_profile', async ({ activity, send }) => {
  const data = activity.value.action.data;
  await send(`Profile saved!\nName: ${data.name}\nEmail: ${data.email}`);
  return {
    statusCode: 200,
    type: 'application/vnd.microsoft.activity.message',
    value: 'Action processed successfully',
  };
});
```

The route name follows the pattern `card.action.<action-name>`, where `<action-name>` matches the value passed to `SubmitData`. This is cleaner than a catch-all with a switch statement, and scales better as you add more actions.

### Catch-All Handler

If you need to handle all card actions in one place, you can use the catch-all handler:

```typescript

  AdaptiveCardActionErrorResponse,
  AdaptiveCardActionMessageResponse,
} from '@microsoft/teams.api';

// ...

app.on('card.action', async ({ activity, send }) => {
  const data = activity.value?.action?.data;
  if (!data?.action) {
    return {
      statusCode: 400,
      type: 'application/vnd.microsoft.error',
      value: {
        code: 'BadRequest',
        message: 'No action specified',
        innerHttpError: {
          statusCode: 400,
          body: { error: 'No action specified' },
        },
      },
    } satisfies AdaptiveCardActionErrorResponse;
  }

  console.debug('Received action data:', data);

  switch (data.action) {
    case 'submit_feedback':
      await send(`Feedback received: ${data.feedback}`);
      break;

    case 'purchase_item':
      await send(`Purchase request received for game: ${data.choiceGameSingle}`);
      break;

    case 'save_profile':
      await send(
        `Profile saved!\nName: ${data.name}\nEmail: ${data.email}\nSubscribed: ${data.subscribe}`
      );
      break;

    default:
      return {
        statusCode: 400,
        type: 'application/vnd.microsoft.error',
        value: {
          code: 'BadRequest',
          message: 'Unknown action',
          innerHttpError: {
            statusCode: 400,
            body: { error: 'Unknown action' },
          },
        },
      } satisfies AdaptiveCardActionErrorResponse;
  }

  return {
    statusCode: 200,
    type: 'application/vnd.microsoft.activity.message',
    value: 'Action processed successfully',
  } satisfies AdaptiveCardActionMessageResponse;
});
```

:::note
The `data` values are not typed and come as `any`, so you will need to cast them to the correct type in this case.
:::

## Responding from a card action

A message sent from a card action handler is placed relative to the card that was clicked.
In a 1:1 chat, when the card is inside a thread, `send()` sends the message into that thread, and when the card is at the top level it sends the message at the top level.
This applies to `Action.Execute` and `Action.Submit`.
In a channel, the response to a card button that arrives as a `message` (`Action.Submit`) is sent through the thread's reply chain, and the SDK sends the response to any other card action to the inbound conversation.
See [Responding to edits, reactions and card actions](../../essentials/sending-messages.md#responding-to-edits-reactions-and-card-actions) for the rules.

## Dynamic Typeahead Search

An `Input.ChoiceSet` can fetch its choices dynamically as the user types, instead of listing them statically. Leave `choices` as an empty array and add a `choices.data` `Data.Query` that names a `dataset`. As the user types, Teams sends your bot an `application/search` invoke, and you return the matching results.

```typescript

const card: IAdaptiveCard = {
  type: 'AdaptiveCard',
  version: '1.5',
  body: [
    {
      type: 'Input.ChoiceSet',
      id: 'game',
      label: 'Game',
      placeholder: 'Search for a game',
      style: 'filtered',
      choices: [],
      'choices.data': { type: 'Data.Query', dataset: 'nintendoGames' },
    },
  ],
  actions: [
    { type: 'Action.Execute', title: 'Submit', data: { action: 'submit_game' } },
  ],
};
```

Register a handler for the `application/search` invoke to return results as ` title, value ` pairs, where `title` is the display text and `value` is what gets submitted when the choice is selected:

```typescript

// ...

const GAMES = ['Super Mario Odyssey', 'Metroid Dread', 'Splatoon 3'];

app.on('card.search', async ({ activity }) => {
  const query = activity.value.queryText?.toLowerCase() ?? '';
  const results = GAMES.filter((g) => g.toLowerCase().includes(query)).map(
    (g) => ({ title: g, value: g })
  );

  return {
    statusCode: 200,
    type: 'application/vnd.microsoft.search.searchResponse',
    value: { results },
  };
});
```

The activity value also carries the query options (`skip`/`top` for paging) and the originating `dataset`, so a single handler can branch on the dataset when several `Input.ChoiceSet` inputs share it.
