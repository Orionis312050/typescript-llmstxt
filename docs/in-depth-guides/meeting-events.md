---
title: "Meeting Events"
sidebar_position: "5"
---

# Meeting Events

Microsoft Teams provides meeting events that allow your application to respond to various meeting lifecycle changes. Your app can listen to events like when a meeting starts, meeting ends, and participant activities to create rich, interactive experiences.

## Overview

Meeting events enable your application to:
- Send notifications when meetings start or end
- Track participant activity (join/leave events)
- Display relevant information or cards based on meeting context
- Integrate with meeting workflows

## Configuring Your Bot

There are a few requirements in the Teams app manifest (`manifest.json`) to support these events.

1. The scopes section must include `team`, and `groupChat`

```json
bots": [
        {
            "botId": "",
            "scopes": [
                "team",
                "personal",
                "groupChat"
            ],
            "isNotificationOnly": false
        }
    ]
```

2. In the authorization section, make sure to specify the following resource-specific permissions:

```json
 "authorization":{
        "permissions":{
            "resourceSpecific":[
                {
                    "name":"OnlineMeetingParticipant.Read.Chat",
                    "type":"Application"
                },
                {
                    "name":"ChannelMeeting.ReadBasic.Group",
                    "type":"Application"
                },
                {
                    "name":"OnlineMeeting.ReadBasic.Chat",
                    "type":"Application"
                }
                ]
            }
        }
```

3. In the Teams Developer Portal, for your `Bot`, make sure the `Meeting Event Subscriptions` are checked off. This enables you to receive the Meeting Participant events. For these events, you must create your Bot via TDP.

## Meeting Start Event

When a meeting starts, your app can handle the `meetingStart` event to send a notification or card to the meeting chat.

```typescript

const app = new App();

app.on('meetingStart', async ({ activity, send }) => {
  const meetingData = activity.value;
  const startTime = new Date(meetingData.StartTime).toLocaleString();

  const card = new AdaptiveCard(
    new TextBlock(`'${meetingData.Title}' has started at ${startTime}.`, {
      wrap: true,
      weight: 'Bolder'
    }),
    // Meetings held inside a channel have no join URL, so there is no action to offer.
    ...(meetingData.JoinUrl
      ? [new ActionSet(new OpenUrlAction(meetingData.JoinUrl).withTitle('Join the meeting'))]
      : [])
  );

  await send(card);
});
```

## Meeting End Event

When a meeting ends, your app can handle the `meetingEnd` event to send a summary or follow-up information.

```typescript

const app = new App();

app.on('meetingEnd', async ({ activity, send }) => {
  const meetingData = activity.value;
  const endTime = new Date(meetingData.EndTime).toLocaleString();

  const card = new AdaptiveCard(
    new TextBlock(`'${meetingData.Title}' has ended at ${endTime}.`, {
      wrap: true,
      weight: 'Bolder'
    })
  );

  await send(card);
});
```

## Participant Join Event

When a participant joins a meeting, your app can handle the `meetingParticipantJoin` event to welcome them or display their role.

```typescript

const app = new App();

app.on('meetingParticipantJoin', async ({ activity, send }) => {
  const meetingData = activity.value;
  const member = meetingData.members[0].user.name;
  const role = meetingData.members[0].meeting.role;

  const card = new AdaptiveCard(
    new TextBlock(`${member} has joined the meeting as ${role}.`, {
      wrap: true,
      weight: 'Bolder'
    })
  );

  await send(card);
});
```

## Participant Leave Event

When a participant leaves a meeting, your app can handle the `meetingParticipantLeave` event to notify others.

```typescript

const app = new App();

app.on('meetingParticipantLeave', async ({ activity, send }) => {
  const meetingData = activity.value;
  const member = meetingData.members[0].user.name;

  const card = new AdaptiveCard(
    new TextBlock(`${member} has left the meeting.`, {
      wrap: true,
      weight: 'Bolder'
    })
  );

  await send(card);
});
```
