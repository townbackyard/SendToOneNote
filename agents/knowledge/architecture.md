# Architecture

> Status: describes the v1 target per the implementation plan. Update as tasks land (plan: `docs/superpowers/plans/2026-08-21-sendtoonenote-v1.md`).

## Repository layout

```
SendToOneNote/
├── src/
│   ├── SendToOneNote.Core/       # All logic; no UI dependencies (net10.0-windows)
│   │   ├── Email/                # EmlParser (MimeKit), ParsedEmail, InlineImage
│   │   ├── Pages/                # PageXhtmlBuilder, ImageResolver (AngleSharp),
│   │   │                         #   PagePlanner, ImageShrinker (System.Drawing)
│   │   ├── OneNote/              # OneNoteClient (Graph), NotebookTree models
│   │   ├── Desktop/              # OneNoteInterop (COM IApplication), StaComWorker,
│   │   │                         #   DesktopOneNoteProbe, HierarchyParser, OneNotePageXmlBuilder
│   │   ├── Backends/             # IOneNoteBackend seam, GraphBackend, DesktopOneNoteBackend,
│   │   │                         #   BackendSelector
│   │   ├── Auth/                 # ITokenProvider, MsalTokenProvider (WAM broker)
│   │   ├── Picker/               # SectionPickerViewModel (UI-free)
│   │   ├── Storage/              # AppSettings, JsonFileStore (settings + tree cache)
│   │   └── Watch/                # DropFolderWatcher, FileReadiness
│   └── SendToOneNote/            # WPF tray app (H.NotifyIcon.Wpf)
│       ├── PickerWindow.xaml     # Classic-style section picker
│       ├── FirstRunWindow.xaml   # Sign-in + folder choice + startup shortcut
│       ├── SavePipeline.cs       # Orchestration: eml path → OneNote page
│       └── TrayContext.cs        # Tray icon, watcher wiring, notifications
├── tests/SendToOneNote.Tests/    # xUnit against Core
├── fixtures/
│   ├── synthetic/                # Committed test .emls (fabricated content)
│   └── local/                    # GITIGNORED — owner's real emails
├── docs/
│   ├── superpowers/specs|plans/  # Design spec + implementation plan
│   ├── entra-app-registration.md # Click-by-click registration + publisher verification
│   └── e2e-checklist.md          # Pre-release manual checklist
├── agents/                       # These docs
├── .github/workflows/            # build.yml (CI), release.yml (tag → win-x64 zip)
├── SendToOneNote.slnx
├── AGENTS.md · CLAUDE.md · README.md · LICENSE (MIT)
```

## The save pipeline

One email flows: `DropFolderWatcher` (readiness-polls the .eml until Outlook releases it) → `EmlParser` (headers, best body, inline CID images, attachment names) → `PickerWindow` over `SectionPickerViewModel` (per-backend recents; type-to-filter; cached notebook tree on Graph, fetched fresh on desktop) → `PageXhtmlBuilder` (title, header table, body; plain-text→paragraphs or `<pre>` when columnar) → `ImageResolver.ResolveWithReportAsync` (normalizes email HTML to well-formed XHTML via AngleSharp; drops tracking pixels/spacers before download; embeds cid:/remote images as `name:imgN` parts, failures keep the original URL; records a per-image `ImageDecision`) → `IOneNoteBackend.CreatePageAsync` (see Backends below — `GraphBackend` ranks/caps/drops via `PagePlanner` then calls `OneNoteClient`; `DesktopOneNoteBackend` base64-inlines every resolved image and writes via COM) → toast with the returned `ClientUrl`; .eml deleted on success or moved to `Failed\`.

The source email account never authenticates — content comes entirely from the local .eml. Only the destination Microsoft account (whose OneDrive holds the notebooks) signs in, and only on the Graph path.

## Backends

`IOneNoteBackend` (`Core/Backends/IOneNoteBackend.cs`) is the seam `SavePipeline` depends on: `Name` (`"desktop"` | `"graph"`), `GetTreeAsync`, `CreatePageAsync(sectionId, pageXhtml, images)`. Two implementations:

- `GraphBackend` — thin wrapper: `PagePlanner.Plan` (rank/cap/drop, see the Graph constraints below) then `OneNoteClient` over HTTP.
- `DesktopOneNoteBackend` — writes into the local desktop OneNote app via COM: `CreateNewPage` → `UpdatePageContent` (a `one:HTMLBlock` page built by `OneNotePageXmlBuilder`, images base64-inlined by `DataUriInliner`) → `GetHyperlinkToObject` for the toast link. On an RPC-disconnect HRESULT (OneNote exited/restarted) it drops and recreates the RCW so the next call re-activates OneNote instead of failing identically forever; a malformed-XML or read-only-section HRESULT leaves the cached RCW in place. Cancellation is honored at the queue boundary (before a queued job starts), not mid-COM-call.

`BackendSelector.Choose(settings.Backend, desktopAvailableProbe)`: `"graph"` forces Graph even when desktop OneNote is present; `"desktop"` forces COM and throws a startup error if unavailable; `"auto"` (default) picks desktop when `DesktopOneNoteProbe.IsAvailable()` succeeds, else Graph. `TrayContext` runs the probe on the `StaComWorker` and shows the choice in the tray tooltip; "Sign in again" is only added to the tray menu when Graph is active.

**StaComWorker rule:** every COM call — probe, hierarchy fetch, page creation, and the smoke test's own read-back — runs through `StaComWorker`, one dedicated STA thread with a serial work queue; the `IApplication` RCW is created on that thread and never touched from another. See `agents/rules/csharp.md`.

Recents and the tree cache are per backend: `AppSettings.RecentSectionIds` (Graph) vs `RecentDesktopSectionIds` (desktop), because section IDs differ between the two. The Graph tree is cached in `cache.json` with a background refresh; the desktop tree has no cache — `GetHierarchy` over COM is ~100 ms, so `SavePipeline` fetches it fresh on every save.

## Graph OneNote API constraints (load-bearing)

- ≤4 MB per request → `PagePlanner.MaxRequestBytes = 3_500_000` safety cap — this is the real binding constraint.
- Binary parts per request: Microsoft documents ~6 multipart sections, but the live service accepted 20 in one POST (verified 2026-08-21); `MaxBinaryPartsPerRequest = 30`. Most emails save every embedded image in one atomic request within these caps.
- Graph path never appends any more: images are ranked by area (first three boosted) and embedded in one request up to the caps; the rest are dropped (removed from the XHTML) per the v1.1 spec. OneNoteClient keeps its append/poll support (tested, unused by the planner).
- Input must be well-formed UTF-8 XHTML; the service strips scripts/forms/complex CSS, tables lose rowspan/colspan. Email layouts simplify — that's the API, not a bug.
- Section-group nesting deeper than one level requires recursive `GET /me/onenote/sectionGroups/{id}/sectionGroups` calls; `$expand` only goes one level down.
- All GETs follow `@odata.nextLink` paging.

## App data (runtime)

`%APPDATA%\SendToOneNote\`: `settings.json` (drop folder, ClientIdOverride, DeleteOnSuccess, `Backend`, `ImageDiagnostics`, recent section ids per backend), `cache.json` (Graph notebook tree only — the desktop tree isn't cached), `msal_cache.bin` (encrypted tokens, Graph only), `logs\stn-YYYYMMDD.log`. Diagnostics dumps (when `ImageDiagnostics` is on) live under `<DropFolder>\Diagnostics\`, not `%APPDATA%`.

## Testing model

- Unit tests: Core only, synthetic fixtures, stubbed `HttpMessageHandler` — no network, no auth.
- `LocalFixtureTests`: run only when `fixtures/local/` exists (owner's machine); skip on CI.
- `IntegrationSmokeTests`: gated on `STN_INTEGRATION=1`; real sign-in, creates pages in a section named "SendToOneNote Test" via Graph.
- `DesktopIntegrationSmokeTests`: gated on `STN_INTEGRATION=1`; requires desktop OneNote and the same "SendToOneNote Test" section; creates a page with a remote + inline image via `DesktopOneNoteBackend` and reads it back with `GetPageContent(piBinaryData)` to assert both are stored embedded.
- WPF layer has no automated UI tests; its logic lives in Core view-models, and the E2E checklist (`docs/e2e-checklist.md`) covers the rest.
