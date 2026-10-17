# QuillPen embedded editor protocol (QuillPost ⇄ QuillPen)

QuillPost can edit a message body in a complete, embedded copy of QuillPen when QuillPen is
installed. The editor runs inside QuillPost's own window (an iframe), so the person never leaves
the message. Both apps implement this file exactly.

## 1. Locating QuillPen

| Platform | Where QuillPost finds QuillPen's web build | Detection |
|---|---|---|
| Android | QuillPen app (applicationId `io.github.socrtwo.quillpen`) exports a read-only ContentProvider, authority `io.github.socrtwo.quillpen.embed`, serving its built web files (`dist/`, i.e. the APK's `public/` assets) by path: `content://io.github.socrtwo.quillpen.embed/index.html`. QuillPost's WebView intercepts requests to `https://quillpen.embed/<path>` (`shouldInterceptRequest`) and answers them from that provider with `ContentResolver.openInputStream`, setting the MIME type from the extension. | `PackageManager.getPackageInfo("io.github.socrtwo.quillpen")` (manifest `<queries><package android:name="io.github.socrtwo.quillpen"/></queries>`) and the provider answering `index.html`. |
| Desktop (Windows/macOS/Linux) | QuillPen's installer ships its web build unpacked in `resources/embed/` (electron-builder `extraResources`), next to `app.asar`. QuillPost desktop looks in the standard install locations (Windows: `%LOCALAPPDATA%\Programs\QuillPen\resources\embed`, `%PROGRAMFILES%\QuillPen\resources\embed`; macOS: `/Applications/QuillPen.app/Contents/Resources/embed`, `~/Applications/...`; Linux: `/opt/QuillPen/resources/embed`, AppImage users set the folder) or a folder chosen in Settings, and serves it from its local server at `/quillpen-embed/<path>`. | `index.html` exists in that folder. |
| iOS | Not available (iOS apps cannot read each other's files). The option is hidden. | — |
| Web | A QuillPen web deployment URL entered in Settings (for example a self-hosted copy). Same-origin or CORS is not needed: only `postMessage` crosses the frame boundary. | URL set and `index.html` loads. |

## 2. Starting the editor

QuillPost loads `<base>/index.html?embed=compose&host=quillpost&v=1` in an `<iframe>` that fills the
compose body area (sandbox: `allow-scripts allow-same-origin allow-popups allow-modals allow-downloads`).

In embed mode QuillPen:
- shows its full ribbon, QAT, AI tab and editing features, but hides File/Backstage commands that
  open, save or close files, the document tabs, and window chrome;
- shows a slim bar with **Done** (send the result back) and **Cancel**;
- never writes drafts, recent files or autosave to its own storage for embedded documents;
- uses the iframe's size (no fixed heights) and respects the same safe-area variables.

## 3. Messages (window.postMessage, JSON objects)

All messages have `type` starting with `quill-embed:` and `v: 1`. Each side checks `event.source`
is the expected window; QuillPost also checks `event.origin` against the base it loaded.

| Direction | type | Fields | Meaning |
|---|---|---|---|
| Pen → Post | `quill-embed:ready` | `app: "QuillPen"`, `version` | Editor loaded and listening. Post then sends `load`. |
| Post → Pen | `quill-embed:load` | `html` (message body HTML), `subject?`, `theme?: "light"\|"dark"`, `lang?` | Replace the editor content. |
| Pen → Post | `quill-embed:changed` | `dirty: true` | Content changed (debounced, at most every 2 s). |
| Post → Pen | `quill-embed:request` | `id` | Ask for the current content without closing (Send pressed). |
| Pen → Post | `quill-embed:content` | `id`, `html`, `text` | Reply to `request`. |
| Pen → Post | `quill-embed:done` | `html`, `text` | Done pressed: Post replaces the body and closes the editor. |
| Pen → Post | `quill-embed:cancel` | — | Cancel pressed: Post closes the editor and keeps the old body. |

`html` is a fragment suitable for an email body: inline styles only (no `<style>` blocks that
depend on classes, no scripts), images as `data:` URIs or `cid:` references unchanged.

## 4. QuillPost user interface

- Settings ▸ Compose: **Edit messages in QuillPen** (Off / Ask / Always), shown only when QuillPen
  is detected; on desktop and web, a field to choose the QuillPen folder or URL.
- Compose window: a **QuillPen** button (also a command, so it can be placed on the QAT) that opens
  the current body in the embedded editor; the built-in editor stays available.
- If QuillPen does not answer `ready` within 8 s, QuillPost shows an error and falls back to the
  built-in editor.
