---
title: "Receiving Inline Images"
sidebar_position: "3"
---

# Receiving Inline Images

When a user pastes an image into the Teams compose box, the image does not arrive inside `activity.text`. Teams sends it as an attachment on the inbound message activity, and `ctx.files` lists it alongside the files the user attached, as a lazy handle you can download from in the same way as an attached file.

:::info[Coming Soon]
Receiving inline images through `ctx.files` is planned for the next SDK release, version 2.2.0.
:::

## List pasted images

Each listed file has an `attachmentStyle` that says how it was attached: `'inlineImage'` for an image pasted into the message, or `'file'` for a file attachment. Pass the style as a filter to `list()` or `first()` to get only pasted images:

```typescript
app.on('message', async ({ files, send }) => {
  const pastedImages = await files.list({ attachmentStyle: 'inlineImage' });

  if (pastedImages.length === 0) {
    await send('Paste an image into your message and I will read it.');
    return;
  }

  await send(`You pasted ${pastedImages.length} image(s).`);
});
```

Without a filter, both styles are returned in the order they arrived. Like any listing, this makes no network requests.

## Download a pasted image

A pasted image downloads through the same methods as an attached file, described in [Receiving Files](./receiving-files.md#read-a-file):

```typescript
const image = await files.first({ attachmentStyle: 'inlineImage' });

if (image) {
  const downloaded = await image.download();
  const base64 = Buffer.from(downloaded.bytes).toString('base64');

  // Pass `downloaded.bytes` or `base64`, with `downloaded.contentType`, to your image-processing or model client.
}
```

## How pasted images are named and typed

A pasted image has no file name, so the SDK names it `image-<attachmentId>`. Its listed metadata comes from a format hint Teams sends with the message:

- `name` gets an extension only for PNG, JPEG, GIF, or WebP, and `extension` is taken from it, so it is absent when the name has none.
- `contentType` is the matching MIME type for those four formats, otherwise `application/octet-stream`.
- `uniqueId` is always unset, because a pasted image is not stored in OneDrive or SharePoint.
- `contentUrl` is the image's Bot API location.

On the downloaded file, `contentType` is detected from the bytes, and the extension of `filename` follows it, so the downloaded file name can differ from the listed name. If the bytes are not PNG, JPEG, GIF, or WebP, the listed type is used.

:::note
Teams saves a pasted GIF as a static PNG. It is listed as `.gif` and `image/gif` but downloads as a PNG.
:::

## Formats that cannot be downloaded

Some image formats, such as BMP and AVIF, cannot be converted by the Teams service, so Teams shows a placeholder, and they cannot be downloaded. Downloading one of these fails with `InlineImageTransformationError`. The same image sent as a file attachment downloads normally, so asking the user to resend it that way is the recovery:

```typescript

try {
  const downloaded = await image.download();
  // ...
} catch (err) {
  if (err instanceof InlineImageTransformationError) {
    await send('That pasted image cannot be downloaded. Try sending it as a file attachment instead.');
  } else {
    throw err;
  }
}
```

If the Bot API refuses the token presented for the image, the download fails with `InlineImageAccessError` with a `401` or `403` status. Both errors share the common base error described in [Handle errors](./receiving-files.md#handle-errors).

## How the image is read

A pasted image is read from the Bot API with the same token the app uses to reply in that conversation, not through Microsoft Graph. This works for traditional bots and for agentic users, which read as their own identity. Inline images list and download in personal chats, group chats, and channels. Outside personal chats, only messages that @mention the agent reach it.

The token is sent only to a URL under the activity's `serviceUrl`. An inline image hosted anywhere else is skipped by `list()` with a warning in the log.

## Access the raw attachments

An inbound message with a pasted image typically contains two related attachments:

```json
{
  "attachments": [
    {
      "contentType": "image/png",
      "contentUrl": "https://.../v3/attachments/{id}/views/original"
    },
    {
      "contentType": "text/html",
      "content": "<p><img src=\"https://.../objects/{id}/views/imgo\" ...></p>"
    }
  ]
}
```

The listed file's `raw` holds only the `image/*` attachment. The `text/html` attachment, which preserves where the image appeared among the message's text and other content, stays in `ctx.activity.attachments`.
