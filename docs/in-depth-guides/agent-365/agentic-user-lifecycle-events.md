---
title: "Agentic User lifecycle events"
sidebar_position: "4"
---

# Agentic User lifecycle events

Agent 365 sends lifecycle events when an Agentic User is created or when its identity, manager, availability, deletion status, or workload onboarding changes. These events let your app keep its state and behavior synchronized with Microsoft 365.

For example, your app can initialize per-agent storage after creation, update ownership when the manager changes, pause scheduled work when the Agentic User is disabled, clean up access after deletion, and wait for Teams onboarding before sending messages.

## Supported events

| Event                         | Typical use                                           |
| ----------------------------- | ----------------------------------------------------- |
| Agentic User identity created | Initialize per-agent configuration and state          |
| Identity property updated     | Refresh mail, alias, or user principal name data      |
| Manager updated               | Update ownership, approval, or escalation routing     |
| Enabled                       | Resume work for the Agentic User                      |
| Disabled                      | Stop scheduled or proactive work                      |
| Deleted                       | Disable access and begin cleanup                      |
| Undeleted                     | Restore state after the Agentic User is restored      |
| Workload onboarding updated   | Wait for Teams onboarding to succeed before messaging |

The event names in activity payloads use `AgenticUser*`, such as `AgenticUserIdentityCreated` and `AgenticUserDisabled`.

## Register handlers

Register a general handler when you need common processing for every lifecycle event. Register typed handlers when you need the payload for a specific variant.

```typescript
app.on('agentLifecycle', async (ctx) => {
  ctx.log.info('Lifecycle event received', {
    valueType: ctx.activity.valueType,
    agenticUserId: ctx.activity.value.agenticUserId,
  });

  // Continue so the matching typed handler can also run.
  await ctx.next();
});

app.on('agenticUserIdentityCreated', ({ activity, log }) => {
  log.info('Agentic User created', { agenticUserId: activity.value.agenticUserId });
});

app.on('agenticUserIdentityUpdated', ({ activity, log }) => {
  log.info('Agentic User identity updated', {
    updatedProperty: activity.value.updatedProperty,
  });
});

app.on('agenticUserManagerUpdated', ({ activity, log }) => {
  log.info('Agentic User manager updated', { manager: activity.value.manager });
});

app.on('agenticUserEnabled', ({ activity, log }) => {
  log.info('Agentic User enabled', { agenticUserId: activity.value.agenticUserId });
});

app.on('agenticUserDisabled', ({ activity, log }) => {
  log.info('Agentic User disabled', { agenticUserId: activity.value.agenticUserId });
});

app.on('agenticUserDeleted', ({ activity, log }) => {
  log.info('Agentic User deleted', { deletionReason: activity.value.deletionReason });
});

app.on('agenticUserUndeleted', ({ activity, log }) => {
  log.info('Agentic User restored', { agenticUserId: activity.value.agenticUserId });
});

app.on('agenticUserWorkloadOnboardingUpdated', ({ activity, log }) => {
  log.info('Agentic User workload onboarding updated', {
    workloadName: activity.value.workloadName,
    workloadOnboardingState: activity.value.workloadOnboardingState,
  });
});
```
