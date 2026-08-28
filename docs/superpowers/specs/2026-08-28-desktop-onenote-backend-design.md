# SendToOneNote v1.1 — Desktop OneNote Backend (COM primary, Graph fallback)

Date: 2026-08-28
Status: Draft for review
Supersedes parts of: `2026-08-21-sendtoonenote-design.md` (the "Auth" and "Page creation" sections become Graph-fallback-only)

## Problem

After a week of daily use, v1's real pain is not failures — it is that pages created through the Graph API take a long time (sometimes minutes) to sync *down* into the owner's desktop OneNote ("Page Not Yet Available"), worst on image-heavy emails. Graph also imposes an upload budget (4 MB/request, part caps, post-create indexing lag, write throttles) that forced retry/poll/append machinery and raised the question of ranking "important" images.

Classic Outlook's Send-to-OneNote button never had these problems because it wrote the page into the **desktop OneNote app's local cache** via COM and let OneNote's own sync push it to the cloud later.

## Spike findings (2026-08-28, owner's machine, desktop OneNote M365 Click-to-Run)

- `OneNote.Application` activates in ~5 ms. **Late binding (`dynamic`, `Type.InvokeMember`) fails** with `TYPE_E_LIBNOTREGISTERED` — Click-to-Run registers the type library in a virtualized registry hive other processes can't see. **Early binding via a hand-declared `[ComImport]` dual interface works** (IID `452AC71A-B655-4967-A208-A4CC39DD7949`, OneNote 2013+ `IApplication`, methods declared in IDL/vtable order). No NuGet interop assembly exists for OneNote.
- Whole flow — hierarchy (17 notebooks incl. a local-only one), create page, write HTML, get hyperlink — in **151 ms**.
- `one:HTMLBlock` renders tables, bold/italic, links, bullets. Remote `<img>` URLs are **downloaded at import and embedded** (13,504 bytes stored for a remote logo; no reference kept) — no link rot. Base64 data-URI images work. The page synced to SharePoint within a minute.

## Goals

- Desktop OneNote present → saves are local and instant, **no Microsoft sign-in at all**, page visible in OneNote immediately, cloud sync handled by OneNote.
- Graph path retained unchanged as the fallback for machines without desktop OneNote; it gains image ranking so image-heavy emails stop failing there too.
- Tracking pixels and spacer images stripped on both paths (also stops sender read-receipts firing from the user's machine).

## Non-goals (v1.1)

- Jumping OneNote to the new page after save (`NavigateTo`) — toast-only, like the legacy button. Backlog if requested.
- A settings UI. Backend choice and diagnostics live in `settings.json`.
- OneNote for Windows 10 (UWP) — no COM API; those users get the Graph path.
- Per-save automatic Graph fallback when a COM save fails (Graph isn't signed in on the desktop path; failures go to `Failed\` like today).

## Architecture

### Backend abstraction (Core)

```csharp
public interface IOneNoteBackend
{
    string Name { get; }                                   // "desktop" | "graph"
    Task<NotebookTree> GetTreeAsync(CancellationToken ct);
    Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml,
        IReadOnlyList<ResolvedImage> images, CancellationToken ct);
}
```

- `GraphBackend` wraps today's `PagePlanner` + `OneNoteClient` unchanged (plus the image ranking below).
- `DesktopOneNoteBackend` is new. `SavePipeline` depends only on `IOneNoteBackend`; everything before the backend (watcher, parser, builder, resolver) is shared.

### Backend selection

`BackendSelector.Choose(settings)`: `settings.Backend` = `auto` (default) | `desktop` | `graph`. **`graph` forces the cloud path even when desktop OneNote is installed** (the owner-requested escape hatch); `desktop` forces COM; `auto` → desktop if `Type.GetTypeFromProgID("OneNote.Application")` resolves AND the QueryInterface for `IApplication` succeeds, else Graph. Selection is logged and shown in the tray tooltip ("SendToOneNote — desktop OneNote" / "— cloud (Graph)").

### First-run & tray on the desktop path

- First-run window: drop-folder choice + startup-shortcut checkbox only; the sign-in button and `MsalTokenProvider` are not constructed. A one-line note says "Saving to your desktop OneNote — no sign-in needed."
- Tray menu: "Open drop folder", "Exit"; "Sign in again" appears only when Graph is active.
- Recents are stored per backend (`RecentSectionIds` stays for Graph; new `RecentDesktopSectionIds`) because section IDs differ.
- No notebook-tree cache on the desktop path — `GetHierarchy` is ~100 ms; the picker calls it fresh each time (fresh notebooks, no stale-cache UX). The Graph cache path is unchanged.

## DesktopOneNoteBackend

### COM interface

`Interop/OneNoteInterop.cs`: `[ComImport, Guid("452AC71A-…"), InterfaceType(InterfaceIsDual)] interface IApplication` with methods in IDL order through `GetHyperlinkToObject` (the spike's declaration, verified against the live app). Enums as `int` constants: `HierarchyScope.hsSections = 3`, `hsPages = 4`; `XMLSchema.xs2013 = 2`; `NewPageStyle.npsDefault = 0`; `PageInfo.piBinaryData = 1`. `string` ↔ BSTR, `bool` ↔ VARIANT_BOOL, `DateTime` ↔ DATE by default marshaling.

### Threading

All COM calls go through one dedicated STA worker thread (`StaComWorker`: a `Thread` with `SetApartmentState(STA)` running a work queue). The RCW is created on that thread and never touched elsewhere. Calls are milliseconds, so the queue is effectively serial and instant; the WPF UI thread is never blocked.

### Hierarchy → NotebookTree

`GetHierarchy("", hsSections, xs2013)` XML (namespace `http://schemas.microsoft.com/office/onenote/2013/onenote`): `one:Notebook` → nested `one:SectionGroup` → `one:Section`. Mapping: `ID` → id, `name` → name; skip elements with `isRecycleBin="true"` or `isInRecycleBin="true"`; skip notebooks whose sections are all recycle bin. Parsed by a pure `HierarchyParser` (unit-tested with a synthetic XML fixture).

### Page creation

1. `CreateNewPage(sectionId, out pageId, npsDefault)`.
2. `UpdatePageContent(pageXml, DateTime.MinValue, xs2013, force: false)` where `pageXml` is:

```xml
<one:Page xmlns:one="…/2013/onenote" ID="{pageId}">
  <one:Title><one:OE><one:T><![CDATA[{subject}]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:HTMLBlock><one:Data><![CDATA[{html}]]></one:Data></one:HTMLBlock>
  </one:OEChildren></one:Outline>
</one:Page>
```

   `{html}` = the same `PageXhtmlBuilder` output as today (header table + body), after `ImageResolver` and a new **`DataUriInliner`** that rewrites each resolved `src="name:imgN"` to `src="data:{contentType};base64,…"`. Images that failed to download keep their URL — OneNote fetches and embeds them itself at import. CDATA safety: any `]]>` in the payload is split as `]]]]><![CDATA[>`. Built by a pure `OneNotePageXmlBuilder` (unit-tested).
3. `GetHyperlinkToObject(pageId, "", out link)` → `CreatedPage.ClientUrl` for the toast's click action. No `NavigateTo`.

### Errors

- ProgID missing / QI fails → not selected (Graph). If `Backend=desktop` is forced and unavailable → clear startup message.
- `COMException` during a save → normal failure path (Failed\ + sidecar + toast). Message maps common HRESULTs: `0x80042000`-range OneNote errors (e.g. section read-only, notebook not open) to plain language; anything else to "Desktop OneNote error 0x…".
- OneNote not running: out-of-process activation launches it (a few seconds on first call). The pipeline already serializes saves, so this cost is paid once.

## Image pipeline changes (both backends)

1. **Junk filter** in `ImageResolver` (after download, before parts are assigned): drop an image when any of: decoded dimensions ≤ 2×2; HTML `width`/`height` attributes both ≤ 2; URL path/query matches `(open|track|pixel|beacon|spacer|blank)\.(gif|png)` or contains `/o.gif`; a 1×1 GIF/PNG by bytes. Dropped images are removed from the HTML entirely (no broken box). Every decision is recorded for diagnostics.
2. **Graph-fallback ranking** in `PagePlanner` (desktop path ignores it — no budget): images sorted by rendered area (HTML width×height attrs, else decoded pixels); embed in rank order until the existing 30-part / 3.5 MB caps; per owner decision, **remaining minor images are dropped** (removed from the HTML), not left as links. Images already in the body's first ~10% of the HTML get a rank boost so logos/banners survive.
3. **Diagnostics dump** (setting `ImageDiagnostics`, default `false`; the owner enables it): per email, `<DropFolder>\Diagnostics\<eml-stem>-<yyyyMMdd-HHmmss>\images.csv` with columns `index,src,bytes,width,height,alt,source(cid|remote),decision(embedded|dropped-junk|dropped-minor|left-as-url),reason`, plus the downloaded image files named `imgN.<ext>`. Never contains email text.

## Settings additions

`Backend` (`auto`|`desktop`|`graph`, default `auto`), `ImageDiagnostics` (bool, default false), `RecentDesktopSectionIds` (list). Existing keys unchanged; existing Graph users' settings keep working.

## Privacy & repo notes

- Desktop path sends nothing to any service; OneNote's own sync is the only network activity. README's privacy section gains this paragraph and keeps the Graph statement for the fallback.
- Tracking-pixel stripping is a privacy feature; note it in README.
- The synthetic hierarchy fixture uses invented notebook names — never the owner's real hierarchy output.

## Testing

- Unit (CI-safe, no OneNote): `HierarchyParser` (nesting, recycle-bin filtering), `OneNotePageXmlBuilder` (title, HTMLBlock, CDATA escaping), `DataUriInliner`, junk filter, ranking under caps, diagnostics CSV writer, `BackendSelector` with an injected probe.
- Gated integration (`STN_INTEGRATION=1`, owner's machine): desktop smoke — hierarchy contains "SendToOneNote Test", create a page with a remote + base64 image, read back with `piBinaryData` and assert both images embedded.
- E2E checklist additions: fresh first-run shows no sign-in on this machine; image-heavy newsletter saves instantly and appears in OneNote immediately; `Backend=graph` override still works end-to-end; tracking pixel absent from the saved page.

## Migration / compatibility

- v1 users on Graph: `auto` switches them to desktop if OneNote is installed; their Graph recents remain if they force `graph`. Entra registration stays (fallback users). Token cache untouched.
- Spike project `D:\Projects\TownBackyard\SendToOneNote-spike-com` is deleted once the backend lands; its interface declaration is the seed for `OneNoteInterop.cs`.

## Decisions log

- COM primary, Graph fallback (owner, 2026-08-28) — instant local write; fixes the sync-down complaint at the root.
- No sign-in when desktop OneNote is present (owner).
- Toast-only after save; no `NavigateTo` (owner).
- Base64-inline resolved images; strip tracking pixels/spacers (owner) — deterministic, privacy-positive; OneNote's own fetch remains the fallback for undownloadable URLs.
- Graph fallback: embed important, drop minor; diagnostics dump next to the drop folder (owner) — kept in v1.1 scope.
- Early-bound hand-declared COM interface, no packages (spike) — late binding is impossible on Click-to-Run Office.
