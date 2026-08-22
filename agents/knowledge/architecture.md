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

One email flows: `DropFolderWatcher` (readiness-polls the .eml until Outlook releases it) → `EmlParser` (headers, best body, inline CID images, attachment names) → `PickerWindow` over `SectionPickerViewModel` (cached notebook tree; recents; type-to-filter) → `PageXhtmlBuilder` (title, header table, body; plain-text→paragraphs or `<pre>` when columnar) → `ImageResolver` (normalizes email HTML to well-formed XHTML via AngleSharp; embeds cid:/remote images as `name:imgN` parts, failures keep the original URL) → `PagePlanner` (splits into Graph-sized requests; shrinks/drops oversized images) → `OneNoteClient` (POST create + PATCH appends) → toast with `onenote:` link; .eml deleted on success or moved to `Failed\`.

The source email account never authenticates — content comes entirely from the local .eml. Only the destination Microsoft account (whose OneDrive holds the notebooks) signs in.

## Graph OneNote API constraints (load-bearing)

- ≤4 MB per request → `PagePlanner.MaxRequestBytes = 3_500_000` safety cap.
- ≤5 binary parts per request besides the `Presentation`/`Commands` part.
- Overflow images: their `<img>` becomes `<div data-id="slot-imgN">&#160;</div>` in the create request (the nbsp is load-bearing — OneNote prunes EMPTY elements even with a data-id, error 20149), then PATCH `[{target:"#slot-imgN", action:"append", content:"<img src=\"name:imgN\"/>"}]` with the binary parts (`append`, not `replace` — Graph rejects replace on div targets, error 20141). A just-created page can 404 (error 20102) until indexed, so appends retry up to 5× with linear backoff.
- Input must be well-formed UTF-8 XHTML; the service strips scripts/forms/complex CSS, tables lose rowspan/colspan. Email layouts simplify — that's the API, not a bug.
- Section-group nesting deeper than one level requires recursive `GET /me/onenote/sectionGroups/{id}/sectionGroups` calls; `$expand` only goes one level down.
- All GETs follow `@odata.nextLink` paging.

## App data (runtime)

`%APPDATA%\SendToOneNote\`: `settings.json` (drop folder, ClientIdOverride, DeleteOnSuccess, recent section ids), `cache.json` (notebook tree), `msal_cache.bin` (encrypted tokens), `logs\stn-YYYYMMDD.log`.

## Testing model

- Unit tests: Core only, synthetic fixtures, stubbed `HttpMessageHandler` — no network, no auth.
- `LocalFixtureTests`: run only when `fixtures/local/` exists (owner's machine); skip on CI.
- `IntegrationSmokeTests`: gated on `STN_INTEGRATION=1`; real sign-in, creates pages in a section named "SendToOneNote Test".
- WPF layer has no automated UI tests; its logic lives in Core view-models, and the E2E checklist (`docs/e2e-checklist.md`) covers the rest.
