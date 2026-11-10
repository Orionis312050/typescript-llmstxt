---
title: "Receiving Files"
sidebar_position: "2"
---

# Receiving Files

Teams users can send files to agents by attaching them to chat messages. When users send files to an agent in 1:1 chat, or to an agentic user in a group chat or channel, they arrive as part of the context of the received message activity. `ctx.files` provides lazy access to the metadata and bytes of each attached file.

An attached file is not delivered inline with the message. Teams stores the file in OneDrive or SharePoint and hands your agent a small metadata attachment describing it, including a short-lived, pre-authorized URL in personal scopes for traditional bots, while agentic users must make a Graph call via the `contentUrl`. Reading the contents means fetching the bytes from either URL over the network, so the download itself is deferred until explicitly called.

This API covers file downloads in **personal (1:1) chats** for every agent. Agentic users can also download files in **group chats** and **channels** starting in version 2.2.0; traditional bots don't receive files there. See [Conversation scope support](#conversation-scope-support).

## Enable file support in your manifest for traditional bots

Teams only offers users the option to attach a file in 1:1 with a bot when your app manifest opts in. Set `supportsFiles` to `true` on your bot entry; without it the compose box shows no attach affordance for your agent, no file reaches your handler, and `ctx.files.list()` is always empty:

```json
"bots": [
  {
    "botId": "${{BOT_ID}}",
    "scopes": ["personal"],
    "supportsFiles": true
  }
]
```

## Enable file support on the blueprint for agentic users

Agentic users need a Graph file permission consented on the agent's blueprint by an administrator in order to access files. See [inheritable permissions](https://learn.microsoft.com/entra/agent-id/concept-inheritable-permissions) for how an agent acquires those scopes.

## Access attached files

The accessor `ctx.files` is available on every activity context, but only inbound message activities carry attachments, so it is empty everywhere else. Use `list()` to get all files on the current message, including pasted images, or `first()` for the common single-file case. Both hand back the file's metadata rather than its contents, so listing is cheap even when the attached files are large:

```typescript
app.on('message', async ({ files, send }) => {
  const attached = await files.list();

  if (attached.length === 0) {
    await send('Send me a file and I will read it.');
    return;
  }

  await send(`You sent ${attached.length} file(s): ${attached.map((f) => f.name).join(', ')}`);
});
```

Listing with `list()` returns the files in the same order they arrived. It never throws: activities with no attachments (or non-message activities) return an empty list, and malformed file entries are skipped rather than failing the whole batch.

```typescript
const file = await files.first();

if (file) {
  await send(`Reading ${file.name}...`);
}
```

Images the user pasted into the message are listed too, with `attachmentStyle` set to `'inlineImage'` (see [Receiving Inline Images](./receiving-inline-images.md)). Files sent as file attachments have the `'file'` style. To list only one style, pass it as a filter to `list()` or `first()`:

```typescript
// Only files sent as file attachments
const attachedFiles = await files.list({ attachmentStyle: 'file' });

// Only images pasted into the message
const pastedImages = await files.list({ attachmentStyle: 'inlineImage' });
```

:::note
Starting in 2.2.0, `list()` and `first()` return pasted images when no filter is given. To keep the previous behavior of returning only file attachments, pass `'file'` as the filter.
:::

### File metadata

Each `IIncomingFile` carries metadata describing the file, populated from what the platform reports:

| Property | Description |
|---|---|
| `uniqueId` | The OneDrive/SharePoint drive-item id, when the platform provides it. Present only for files backed by ODSP storage. |
| `name` | File name including its extension (e.g. `report.pdf`). |
| `extension` | File extension without the dot (e.g. `pdf`), from the platform-supplied `fileType`. Absent when the platform omits it. |
| `attachmentStyle` | How the file was attached: `'file'` when sent as a file attachment, or `'inlineImage'` when pasted into the message. |
| `contentType` | The file's MIME type, when the source provides one. Always unset for file attachments received from a bot activity: the `file.download.info` attachment carries no MIME type, only the extension surfaced as `extension`. To learn the type of the bytes you actually received, read `contentType` on the downloaded file, which is resolved from the download response. For an inline image, see [Receiving Inline Images](./receiving-inline-images.md#how-pasted-images-are-named-and-typed). |
| `scope` | The conversation scope the file arrived in (`personal`, `groupChat`, or `channel`). |
| `source` | Where the SDK found the file. Currently always `botActivity`. |
| `contentUrl` | A browsable link to the file in OneDrive/SharePoint, when known. Not a fetchable download URL. For an inline image, it is the image's Bot API location. |
| `raw` | The original wire attachment (the metadata object, not the bytes) — see [Access the raw attachment](#access-the-raw-attachment). |

## Read a file

The most common path is `download()`, which fetches the whole file and buffers it into an in-memory copy you own:

```typescript
const file = await files.first({ attachmentStyle: 'file' });

if (file) {
  const downloaded = await file.download();

  await send(`Downloaded ${downloaded.filename} (${downloaded.bytes.length} bytes, ${downloaded.contentType})`);
}
```

For text files, `text()` is a shortcut for `download()` followed by decoding:

```typescript
const file = await files.first({ attachmentStyle: 'file' });

if (file) {
  const contents = await file.text();
  await send(`The file starts with: ${contents.slice(0, 100)}`);
}
```

:::note
Decoding with `text()` reads the downloaded bytes as UTF-8 by default, and is lossy: bytes that are not valid for the encoding become the replacement character (`U+FFFD`) instead of throwing. To decode a different encoding, pass an encoding name (e.g. `file.text('latin1')`). For binary files, read the raw bytes instead.
:::

To save a file to disk without holding it all in memory, use `saveAs()`, which streams straight to the path:

```typescript
await file.saveAs('./downloads/report.pdf');
```

The path is resolved relative to the process working directory, overwrites an existing file, and does not create missing parent directories.

:::warning
A file's `name` comes off the wire and is set by the uploader, so treat it as untrusted input rather than a destination. It may contain `/` or `\`, a drive or UNC prefix, or `..` segments, and no sanitization is performed on the developer's behalf. Therefore, no assumptions are made related to the destination like illegal characters based on operating system or naming collisions.
:::

For large files or streaming pipelines, `stream()` hands you the raw byte stream so you can process it as it arrives:

```typescript
const stream = await file.stream();

for await (const chunk of stream) {
  // process each chunk (e.g. pipe to a parser)
}
```

An `IIncomingFile` handle holds no cached bytes: each call to `download()`, `text()`, `stream()`, or `saveAs()` fetches the file again. This matters because a file's download link is short-lived: the URL embeds its own pre-authorized `tempauth` credential, which can expire between reads. To read the same file several ways, call `download()` once and reuse the downloaded file it returns, which is a point-in-time copy whose readers never re-fetch:

```typescript
const downloaded = await file.download();

const text = downloaded.text();          // decode as UTF-8
const buffer = downloaded.arrayBuffer(); // the raw bytes
await downloaded.saveAs('./copy.bin');   // write to disk, no re-fetch
```

| Property / method | Description |
|---|---|
| `bytes` | The buffered file bytes as a `Uint8Array`. |
| `contentType` | MIME type resolved from the download response header, or the incoming file's metadata type if the response omits one. Falls back to `application/octet-stream` when neither provides a type, so this is never empty. |
| `filename` | The resolved file name. |
| `sourceUrl` | The URL the bytes were fetched from. |
| `text(encoding?)` | Decode the bytes as UTF-8 (or a given encoding). Lossy; never throws. |
| `arrayBuffer()` | Return the bytes as an `ArrayBuffer`. |
| `saveAs(path)` | Write the buffered bytes to a local path (no re-fetch). |

## Handle errors

Byte retrieval can fail for a few well-defined reasons. All of them share a common base error, so a single check catches any of them and a new one can be added later without callers changing:

```typescript

  FileAccessError,
  FileCredentialError,
  FileError,
  FileScopeNotSupportedError,
  FileUrlExpiredError,
  InlineImageAccessError,
  InlineImageTransformationError,
} from '@microsoft/teams.apps';
```

| Error | Raised when | Carries |
|---|---|---|
| **File URL expired** | The short-lived, pre-authorized download link lapsed before the bytes were fetched. | `reason`: <br/> **first fetch**, the link had already lapsed so no bytes were retrieved; <br /> **re-read**, an earlier download succeeded, but a later read through the same handle came after expiry. Avoid this by downloading once and reusing the downloaded file. |
| **No Graph credential** | The Graph read an agentic user's file needs was ruled out before the request, because no credential was available. | `actor`<sup><a href="#actor">2</a></sup>: the identity a credential was being resolved for. <br /> `cause`: why acquiring the token failed, when it failed rather than simply returning nothing. |
| **File access denied** | The Graph read an agentic user's file needs was refused by the storage service. | `status`: `401` when the token itself was rejected, `403` when the identity lacks the grant<sup><a href="#access-denied">1</a></sup>. <br/> `actor`<sup><a href="#actor">2</a></sup>: the identity that was refused. <br /> `details`: what the service said, if anything. |
| **Inline image could not be converted** | The Teams service could not convert a pasted image for download. See [Formats that cannot be downloaded](./receiving-inline-images.md#formats-that-cannot-be-downloaded). | `status`: `415`. <br /> `code`: the service's error code, such as `TransformationFailed`, when present. <br /> `details`: what the service said, if anything. <br /> `actor`<sup><a href="#actor">2</a></sup>: the identity the read was attempted as. |
| **Inline image access denied** | The Bot API refused the token presented for a pasted image. | `status`: `401` when the token itself was rejected, `403` when the identity is not allowed to read the image. <br /> `actor`<sup><a href="#actor">2</a></sup>: the identity that was refused. <br /> `details`: what the service said, if anything. |
| **File scope not supported** | Download was called on a file from a conversation scope this API does not yet support, meaning anything other than personal, group chat, or channel for an agentic user. A traditional bot receives files only in personal (1:1) chats. | `scope`: the conversation scope. |

<sup id="access-denied">1</sup> A **`403`** deliberately collapses causes the SDK cannot differentiate: an unconsented permission, a file the identity was never granted, and an item that does not exist all arrive the same way. There is no separate "file not found" case to branch on; `details` carries the service's own text for diagnosis. In a standard channel, a file picked from OneDrive or SharePoint, rather than uploaded, can be refused this way, because Teams does not grant the agent access to the picked item (see [Group chats and channels](#group-chats-and-channels)). A `401` is a different problem, which is why the status is reported rather than folded in: the token was rejected, so the fix is the token rather than the sharing.

<sup id="actor">2</sup> `actor` names which identity the read was attempted as. The SDK chooses it, and today there are two:

- **Agentic user**, chosen whenever the inbound activity carries an agentic identity. An agentic user reads as itself rather than as an app. Its Graph scopes come from its blueprint's inheritable permissions set by a tenant administrator or from a direct grant by a user with permissions to share it.
- **App**, chosen for every other activity. Graph file retrieval is supported for agentic users, which read as their own identity; an app identity and user-delegated permissions may be used, but are not supported through the SDK at this time.

The actor is absent when the failure came before any identity was selected, which is the case when the app has no Graph route configured at all.

A transport or service failure, such as a Graph 5xx, is not one of these and surfaces as an ordinary error.

You may add error handling based on the type of error: 

```typescript
try {
  const downloaded = await file.download();
  // ...
} catch (err) {
  if (err instanceof FileUrlExpiredError && err.reason === 'firstFetch') {
    await send('That file link has expired before it could be read.');
  } else if (err instanceof FileCredentialError) {
    await send(`Could not read that file as ${err.actor ?? 'the identity used'}: no Graph credential was available.`);
  } else if (err instanceof FileAccessError) {
    await send(`Could not read that file as ${err.actor ?? 'the identity used'}: the service returned ${err.status}.`);
  } else if (err instanceof InlineImageTransformationError) {
    await send('That pasted image cannot be downloaded. Try sending it as a file attachment instead.');
  } else if (err instanceof InlineImageAccessError) {
    await send(`Could not read that pasted image as ${err.actor ?? 'the identity used'}: the service returned ${err.status}.`);
  } else if (err instanceof FileScopeNotSupportedError) {
    await send(`Downloading files from ${err.scope} conversations is not supported yet.`);
  } else if (err instanceof FileError) {
    // Any future inbound-file failure lands here rather than escaping unhandled.
    await send('That file could not be read.');
  }
}
```

## Access the raw attachment

Every file the accessor returns retains its original wire attachment on `raw`, for diagnostics or for the rare case where you need a field the typed surface does not expose:

```typescript
const file = await files.first();

if (file) {
  // `raw` is the untyped wire attachment: the escape hatch when you need a
  // field the typed surface does not expose. Here we log the whole attachment
  // to inspect exactly what the platform sent.
  log.debug('raw file attachment', file.raw);
}
```

Attachments that are **not** files or inline images (Adaptive Cards, mentions, link previews), and malformed file entries the accessor skips, are not returned by `list()` and therefore have no `IIncomingFile`. Reach those directly through `ctx.activity.attachments`. The HTML attachment that accompanies a pasted image is also only in that array; see [Receiving Inline Images](./receiving-inline-images.md).

## Conversation scope support

Listing with `list()` surfaces the files Teams delivers on the incoming activity. What gets delivered and downloaded depends on the conversation scope and on whether your agent is a traditional bot or an agentic user.

### Personal chats

Downloading in personal (1:1) chats depends on the URL the platform sends, not on who is reading. A traditional bot receives a short-lived, pre-authorized download link, which the SDK fetches directly with no token attached.

An **agentic user** receives no pre-authorized link. Its attachment data carries only a `contentUrl` locating the item in OneDrive or SharePoint, so the SDK resolves the bytes through Microsoft Graph's `/shares` endpoint using the agent's own credential. The call behavior is unchanged from above; only the downloading route underneath differs.

If an agentic user has no usable Graph credential, the read fails with `FileCredentialError`; if the storage service refuses the read, for example because the file isn't shared with the agent, it fails with `FileAccessError` (see [Handle errors](#handle-errors)). Both name the identity involved rather than reporting the file as missing.

### Group chats and channels

:::info[Coming Soon]
Receiving files as an agentic user in group chats and channels is planned for the next SDK release, version 2.2.0.
:::

In **group chats** and **channels**, the attachment has the same shape, carrying only a `contentUrl`. The SDK reads it through Graph `/shares` using the agent's own credential. Your handler code is the same too: call `list()` and then `download()`, and the SDK chooses the agent's credential.

Outside personal chats, only messages that @mention the agent reach it, so a file posted without an @mention is never delivered.

A file can be attached by uploading it from the device or by picking it from OneDrive or SharePoint. Whether each kind downloads depends on the conversation:

- In **group chats**, both kinds download.
- In **private channels**, both kinds download. A picked file is delivered as a copy in the channel's site.
- In **non-private channels**, files uploaded from the device download, but a picked file can be refused, because Teams does not grant the agent access to the picked item. The refusal surfaces as `FileAccessError` with a `403` status (see [Handle errors](#handle-errors)).

The SDK reads each file as the agent, so a download succeeds only when the agent already has access to that file. A sharing link that doesn't name the agent, such as an "Anyone" link or an organization-wide link, doesn't give it that access.

Files attached to a reply in a channel thread follow the same rules as files on a new post.

A traditional bot cannot receive files in group chats or channels: the attached file isn't delivered to it on the activity, so `list()` doesn't surface it.

## Next steps

- See [Files in Teams bots](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/bots-filesv4) for the platform-side file model.
