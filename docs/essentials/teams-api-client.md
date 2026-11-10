---
title: "Teams API Client"
sidebar_position: "6"
---

# Teams API Client

Teams has a number of areas that your application has access to via its API. These are all available via the `app.api` object. Here is a short summary of the different areas:

| Area            | Description                                                                                                                                                          |
| --------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `conversations` | Gives your application the ability to perform activities on conversations (send, update, delete messages, etc.), or create conversations (like 1:1 chat with a user) |
| `meetings`      | Gives your application access to meeting details and participant information via `getById` and `getParticipant`                                                       |
| `teams`         | Gives your application access to team or channel details                                                                                                             |

An instance of the API client is passed to handlers that can be used to fetch details:

## Example

In this example, we use the API client to fetch the members in a conversation. The `api` object is passed to the activity handler in this case.

```typescript
app.on('message', async ({ activity, api }) => {
  const members = await api.conversations.getMembers(activity.conversation.id);
});
```

## Proactive API

It's also possible to access the API client from outside a handler via the app instance. Here we have the same example as above, but we're access the API client via the app instance.

```typescript

const res = await app.api.graph.call(endpoints.chats.getAllMessages.get);
```

## Meetings Example

In this example, we use the API client to get a specific meeting participant's details, such as their role (e.g. Organizer) and whether they are currently in the meeting. Provide the user's AAD Object ID to specify which participant to look up. The `meetingId` and `tenantId` are available from the activity's channel data.

:::note
To retrieve **all** members of a meeting, use the conversations API as shown in the [example above](#example), since meetings are also conversations.
:::

```typescript
app.on('meetingStart', async ({ activity, api }) => {
  const meetingId = activity.channelData?.meeting?.id;
  const tenantId = activity.channelData?.tenant?.id;
  const userId = activity.from?.aadObjectId;

  if (meetingId && tenantId && userId) {
    const participant = await api.meetings.getParticipant(meetingId, userId, tenantId);
    // participant.meeting?.role — "Organizer", "Presenter", "Attendee"
    // participant.meeting?.inMeeting — true/false
  }
});
```

Visit [Meeting Events](../in-depth-guides/meeting-events.md) to learn more about meeting events.
