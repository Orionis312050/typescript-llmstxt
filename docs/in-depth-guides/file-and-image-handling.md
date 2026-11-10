---
title: "File and Image Handling"
sidebar_position: "1"
---

# File and Image Handling

This section covers how your app receives files and inline images and reads their bytes, and how it sends images back.

:::info[Preview]
The high-level file API is being rolled out incrementally. Today it supports **receiving** files attached to messages via the TypeScript, Python, and C# SDKs: in personal (1:1) chats for every agent, and soon in group chats and channels for agentic users. Receiving inline images through the same API is planned for version 2.2.0.
:::

## Files vs. attachments

In personal scope, `activity.attachments` carries every non-text part: files the user attached, inline images, Adaptive Cards, `@mentions`, link previews, and other content the client or the platform adds. A **file** is the specific subset of attachments that describes a document the user attached, one whose `contentType` is `file.download.info`.

The `ctx.files` accessor is the attached-file view over that raw array. It maps each attached file and inline image to a lazy handle you can download from, and skips everything else. The original `activity.attachments` array is always available when you need the raw payload or a non-file attachment.

An **inline image** is an image attachment with an authenticated download URL, commonly paired with an HTML attachment that describes where the image appeared in the composed message. Inline images are listed by the high-level file API alongside attached files, marked with an inline image attachment style.

Sending works similarly: an image you send is an attachment whose content type is an image MIME type, which is why it renders inline rather than arriving as a document.

## In this section

- **[Receiving Files](./file-handling/receiving-files.md)** — read files a user attaches to a message, stream or download their bytes, and handle unsupported scopes, expired links, and agentic users.
- **[Receiving Inline Images](./file-handling/receiving-inline-images.md)** — list images pasted into a message, download their bytes, and handle image formats that cannot be downloaded.
- **[Sending Inline Images](./file-handling/sending-inline-images.md)** — send an image that renders directly in the conversation, from a hosted URL or as base64-encoded bytes, within the platform's size and format limits.

## Resources

- [Files in Teams bots](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/bots-filesv4)
- [Picture messages](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-messages#use-picture-messages)
