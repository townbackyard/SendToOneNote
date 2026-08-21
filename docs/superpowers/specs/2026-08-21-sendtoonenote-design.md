# SendToOneNote — Design Spec

Date: 2026-08-21
Status: Draft for review

## Problem

Classic Outlook for Windows has a "Send to OneNote" button: pick a OneNote section from a dialog, and the email becomes a OneNote page — subject as page title, body as editable HTML. New Outlook's replacement (a built-in web add-in) is unreliable (it has vanished from the Apps menu for many users, including the project owner, with no reinstall path) and can never work for Gmail/IMAP accounts, because new Outlook does not support add-ins on non-Microsoft accounts.

## Solution overview

**SendToOneNote** is a small open-source Windows tray app. The user drags an email out of new Outlook (any account type — Exchange/M365, Outlook.com, Gmail, any IMAP) into a watched folder. New Outlook saves it as a `.eml` file. The app detects it, pops a section-picker dialog (search box, recent sections, notebook tree), and creates a OneNote page in the chosen section via the Microsoft Graph OneNote API: subject as title, a From/To/Date header block, then the full HTML body with images embedded.

Email content flows entirely from the local `.eml` file, so the *source* account needs no authentication. Only the *destination* — the Microsoft account whose OneDrive holds the notebooks — signs in.

### Goals

- Replicate the classic Send-to-OneNote workflow (per-email section picker, HTML page, subject as title) for new Outlook, all account types.
- Zero-friction for non-technical users: download, run, sign in once.
- Open source (MIT) on GitHub; useful beyond the project owner.

### Non-goals (v1)

- Embedding file attachments (attachment names are listed in the header block; files are not embedded). — v2 candidate.
- Right-click "Send to OneNote" shell handler for `.eml` files. — v2 candidate.
- Non-`.eml` file types in the drop folder (ignored in v1). — v2 candidate.
- Code signing (ships unsigned; README documents the SmartScreen unblock).
- Non-Windows platforms; localization.

### v2 candidates

Deliberately deferred; captured here so v1 design doesn't preclude them:

- **Include-attachments option**: a yes/no setting (and/or per-save checkbox in the picker) that embeds file attachments on the page instead of only listing their names.
- **Auto-send to the same section**: an "always send to the selected location" mode mirroring the classic dialog's checkbox — skips the picker and files straight to the remembered section, with an obvious way to turn it back off (tray menu + settings).
- **Other drop-folder file types**: `.txt` creates a page with the file's text content; `.png`/`.jpg` create a page with the image embedded in the body (not as an attachment). Filename becomes the page title; the picker flow is identical.
- **Right-click "Send to OneNote"** shell handler for `.eml` files.

## Architecture

Single .NET 10 C# WPF application, tray-resident. Projects:

- `SendToOneNote` — WPF app: tray icon, picker window, settings window, toast notifications.
- `SendToOneNote.Core` — class library, no UI dependencies: watcher, parser, page builder, Graph client, auth. All logic that unit tests target.
- `SendToOneNote.Tests` — xUnit tests against `Core`, using committed synthetic fixtures.

Components in `Core`:

| Component | Responsibility |
|---|---|
| `FolderWatcher` | `FileSystemWatcher` on the drop folder; debounce + retry-until-unlocked before handing off (new Outlook may still be writing the file). |
| `EmlParser` | MimeKit-based. Extracts headers, best body (HTML preferred, plain text fallback), inline CID images, attachment names. |
| `PageBuilder` | Produces the XHTML page: `<title>` = subject, header table (From, To, Cc, Sent date, attachment names), body. Resolves `cid:` references to MIME parts; downloads remote images; recompresses only when a request would exceed API limits. |
| `OneNoteClient` | Graph calls: list notebooks/section groups/sections; create page (multipart); append remaining images via PATCH when they exceed per-request part limits. |
| `AuthManager` | MSAL.NET public client with WAM broker; encrypted token cache; silent after first interactive sign-in. |
| `SectionCache` | Persists the notebook/section tree and recent-section list locally so the picker opens instantly; refreshes in the background on each popup and via manual refresh. |

## User flow

1. One-time setup: run the app → sign in with the Microsoft account that owns the notebooks → choose (or accept default) drop folder, e.g. `Documents\SendToOneNote Drop`. The app offers to pin the folder to Explorer Quick Access and to start with Windows (Startup-folder shortcut — created only with the user's consent).
2. Daily use: drag an email from new Outlook into the folder → picker pops to foreground → type to filter sections or pick from recents/tree → Enter.
3. Result: toast "Saved to <Section>" with a link that opens the page in OneNote. The `.eml` is deleted on success.
4. Failure: `.eml` moves to a `Failed` subfolder; toast states the reason. Dragging the file back into the drop folder retries.

## Picker dialog

Modeled on the classic dialog (see problem statement):

- Search box with type-to-filter across all sections, matching section, section group, and notebook names; results shown as `Section (Notebook » Group)`.
- Recent sections (most recently used first, persisted, ~10 entries) listed above the tree.
- Expandable tree: notebooks → section groups (nested) → sections.
- Keyboard-first: opens focused on the search box; arrow keys navigate; Enter confirms; Esc cancels (leaves the `.eml` in the folder untouched).
- Data comes from `SectionCache`; a subtle indicator shows when a background refresh is running. Graph call: `GET /me/onenote/notebooks?$expand=sections,sectionGroups($expand=sections)` with `$select` trimming, following `@odata.nextLink` paging.

## Page creation

- `POST /me/onenote/sections/{id}/pages` as `multipart/form-data`: a required `Presentation` XHTML part plus binary image parts referenced as `<img src="name:partN"/>`.
- Header block: single-row-per-field table (From, To, Cc when present, Sent; attachment names when present) styled within OneNote's supported inline CSS.
- Plain-text-only emails render as paragraphs preserving line breaks (the classic behavior), not `<pre>`, unless the text is columnar (receipt-style alignment detected by run-length of spaces) — then `<pre>`.
- Images:
  - Inline `cid:` images resolve to their MIME parts and are embedded.
  - Remote (`http(s)`) images are downloaded (bounded parallelism, per-image timeout ≈10 s) and embedded so pages survive link rot. On download failure, the original URL is left in place — the page never fails because of one image.
  - Graph limits respected: ≤4 MB per request, limited binary parts per request. The page is created with the first batch; remaining images are appended with `PATCH /me/onenote/pages/{id}/content` in batches. An image is recompressed/downscaled only if it alone would exceed the request cap.
- Known fidelity caveat (documented in README): the Graph OneNote API sanitizes HTML — scripts/forms/external CSS removed, limited inline CSS, tables without rowspan/colspan. Content, links, and images survive; complex marketing layouts simplify.

## Authentication & app registration

- MSAL.NET `PublicClientApplication`, `https://login.microsoftonline.com/common`, WAM broker on Windows, encrypted token cache (MSAL extensions). Delegated scopes: `User.Read`, `Notes.ReadWrite` (+ `offline_access`).
- Entra registration (owner's tenant, one-time, manual): public client (no secret), **multi-tenant + personal Microsoft accounts** (`AzureADandPersonalMicrosoftAccount`) so both work/school and consumer OneDrive notebooks work.
- The registration's client ID ships as the app default (public identifier, safe in an OSS repo). Config override `ClientId` supports orgs that block third-party apps (bring-your-own registration, README walkthrough provided).
- **Publisher verification (launch checklist):** TownBackyard has a Microsoft Partner account, so the registration will be publisher-verified before release — associate the Partner (MPN) ID in Entra → app registration → Branding & properties, with `townbackyard.com` set and verified as publisher domain. Users then see a verified-publisher checkmark instead of "unverified", and default consent policies in many tenants stop blocking the app. Prerequisites noted for the manual steps doc: valid Partner Center account, verified MPN ID, matching verified domain.
- Corporate reality (documented in README): some tenants still require one-time admin consent for `Notes.ReadWrite` regardless of verification. Outs: admin consent URL, or BYO client ID.
- Privacy statement (README): email content is read from the local `.eml` and sent only to Microsoft Graph under the user's own sign-in. Nothing is sent to the project, its owner, or any third party. No telemetry.

## Configuration

JSON at `%APPDATA%\SendToOneNote\settings.json`: drop folder path, client ID override, delete-vs-archive on success (default delete), recent sections, cached tree. No secrets (tokens live in the MSAL encrypted cache).

## Error handling

- File locked/still syncing: retry with backoff before parse; give up after ~30 s with a toast.
- Parse or Graph failure: move `.eml` to `Failed\`, toast with a one-line reason; details to a rolling log file at `%APPDATA%\SendToOneNote\logs\`.
- Offline / auth expired: picker still opens from cache; save fails with a clear toast (file stays in the drop folder for retry after reconnect/re-sign-in).
- Non-`.eml` files dropped into the folder: ignored, logged once (v1; see v2 candidates for `.txt`/image support).

## Open source & distribution

- MIT license. Public GitHub repo at `https://github.com/townbackyard/SendToOneNote` (creation and every push explicitly confirmed by the owner; nothing published before that).
- Fixtures: synthetic `.eml` files modeled on real-world structure (multipart HTML with remote images, plain-text-only receipt, HTML with inline CID images) committed for tests. Real personal `.eml` files live in `fixtures/local/`, which is gitignored.
- README: what it is (with screenshot), download/install, first-run sign-in, corporate-consent guidance, BYO-registration walkthrough, SmartScreen unblock note, privacy statement, fidelity caveat, troubleshooting (drag-out regression workaround: click the message before dragging; "… → Save as" fallback).
- CI: GitHub Actions — build + tests on PR/push; release workflow publishes a self-contained single-file exe (win-x64) as a zip on tagged releases.

## Testing

- Unit: `EmlParser` and `PageBuilder` against the synthetic fixtures (HTML+remote images, plain-text receipt, inline-CID case, malformed/truncated `.eml`).
- Integration (local only, owner's account): create pages into a scratch "SendToOneNote Test" section; verify title, header block, images, plain-text rendering. Also validates the real personal fixtures in `fixtures/local/`.
- Manual E2E checklist before each release: drag from an M365 account and a Gmail account in new Outlook; image-heavy email; plain-text email; cancel from picker; failure path (airplane mode); fresh sign-in on a second Windows account.

## Manual steps performed by the owner (nothing automated by tooling)

1. Entra app registration + publisher verification (click-by-click doc to be written during implementation).
2. GitHub repository creation and pushes.
3. Installing/running the app, Startup shortcut, pinning the drop folder.

## Decisions log

- Watched-folder capture (drag from new Outlook; direct drop onto app windows is unsupported by new Outlook's architecture) — chosen over right-click-only.
- Download & embed images — chosen over leaving remote links (link rot).
- .NET/C# WPF — chosen over PowerShell+WPF.
- MIT — chosen over Apache-2.0/GPLv3.
- Ship owner's client ID with BYO override — chosen over BYO-only (non-technical-user goal).
- Support work/school AND personal Microsoft accounts in v1.
- Feasibility verified 2026-08-20/21 on the owner's machine: drag-out produces faithful `.eml` from both M365 and Gmail accounts (full headers, full HTML, remote image URLs intact; plain-text-only source exports faithfully).
