# Background — why this app exists

Research verified 2026-08-20/21 (sources: Microsoft Learn, Microsoft Support, Microsoft Q&A). These platform facts shape the design; if one changes, the design assumption built on it should be revisited.

## The gap

Classic Outlook's "Send to OneNote" ribbon button (COM add-in: section-picker dialog → editable HTML page, subject as title) has no reliable equivalent in **new Outlook for Windows**:

- New Outlook does not support COM/VSTO add-ins at all — only web add-ins.
- Microsoft's replacement is a built-in "Send to OneNote" **web add-in** (Apps icon on an open email). It officially exists ("available in both" per Microsoft's classic-vs-new comparison article), but since ~Nov 2025 it has repeatedly vanished from users' Apps menus with **no reinstall path** (confirmed by Microsoft moderators), had multiple 2025–2026 service incidents, and fails on OneDrive libraries with >5,000 OneNote items. The owner's machine is affected: the add-in is absent even on an M365 account.
- **Web add-ins cannot run on non-Microsoft accounts** (Gmail/IMAP/POP) in new Outlook — per Microsoft Learn's supported-accounts matrix. No shipped or announced change as of Aug 2026. An in-Outlook button for Gmail accounts is impossible; this is the constraint that rules out building an Outlook add-in.
- `me@onenote.com` email-in was retired 2025-03-26 (MC1011145). Power Automate's OneNote connector can't take a consumer-Gmail trigger (Google's 2020 connector policy) and has no send-time section picker anyway.

## Why capture is "drag to a folder"

- New Outlook saves a dragged message as a **.eml** file (classic saved .msg). Verified on the owner's machine 2026-08-20 for both M365 and Gmail accounts: full headers, complete HTML bodies, remote image URLs intact; a text-only source email exports faithfully as text-only.
- Dropping messages **directly onto other apps' windows is unsupported by design** (WebView2 architecture; Microsoft moderator statement). The supported pattern is drag-to-filesystem, then the file. Hence: watched folder, not a drop-target window.
- Known Outlook regression (Nov 2025, still open mid-2026): single-message drag-out can fail unless the message is clicked/selected first. "… → Save as" always works. Documented in the README troubleshooting section.

## Why the OneNote side is Graph (v1) / desktop OneNote when available (v1.1)

- The Microsoft Graph OneNote API (`/me/onenote/...`) is the supported programmatic surface for OneDrive/SharePoint-hosted notebooks, works for both work/school and personal Microsoft accounts with **delegated** auth (app-only auth was retired March 2025 — irrelevant here), and has no announced deprecation as of Aug 2026.
- Classic Send-to-OneNote used the local desktop OneNote COM API, which imported richer HTML. The Graph API sanitizes harder — fidelity is deliberately documented as "content, links, tables, images survive; fancy newsletter layout simplifies." v1.1 restores the desktop COM path as the primary backend when desktop OneNote is installed, keeping Graph as the fallback for machines without it (see the Backends section of `agents/knowledge/architecture.md`).

## Why the legacy button was fast

- Classic Outlook's "Send to OneNote" ribbon button never talked to a cloud API at all — it wrote the page via COM directly into desktop OneNote's **local cache** and let OneNote's own background sync push it to the cloud afterward. That's why it felt instant: the write is local disk, not a network round trip.
- v1's Graph-only design lost this: a page created through Graph has to sync back *down* into the owner's desktop OneNote before it's visible there ("Page Not Yet Available"), sometimes minutes on image-heavy emails, because Graph's own indexing and write throttles sit in between. v1.1's `DesktopOneNoteBackend` reproduces the classic button's approach — COM into the local cache — which is why it's the primary path when desktop OneNote is present.
- **Click-to-Run type-library fact (verified 2026-08-28, spike + live app):** Office installed via Click-to-Run registers the OneNote COM type library in a virtualized registry hive that other processes can't see. Late binding (`dynamic`, `Type.InvokeMember`) against `OneNote.Application` fails with `TYPE_E_LIBNOTREGISTERED`. Only vtable calls through a hand-declared `[ComImport]` dual interface (`IApplication`, IID `452AC71A-B655-4967-A208-A4CC39DD7949`, methods in IDL order) work — there is no NuGet interop assembly for OneNote. This is why `SendToOneNote.Core.Desktop.OneNoteInterop` declares the interface by hand instead of using `dynamic` or a generated PIA.

## Auth model

- Public-client OAuth: the shipped client ID is a public identifier (registered in the owner's tenant, multi-tenant + personal accounts). Users' tokens and email content flow directly between their machine and Microsoft — nothing transits the project or its owner.
- TownBackyard has a Microsoft Partner account; the app registration gets **publisher verification** (MPN ID association) before public release so corporate consent policies don't block it. Some tenants still require one-time admin consent for `Notes.ReadWrite` — that's their policy, not fixable in code; the README documents the admin-consent and bring-your-own-client-ID outs.
