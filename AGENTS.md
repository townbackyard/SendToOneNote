# SendToOneNote — AI Agent Guide

SendToOneNote is a **.NET 10 WPF Windows tray app** (MIT, open source) that brings classic Outlook's "Send to OneNote" back for new Outlook — including Gmail/IMAP accounts where Outlook add-ins can't run. Drag an email (.eml) into a watched folder → section-picker dialog → the email becomes a OneNote page (subject as title, From/To/Date block, HTML body with embedded images) — written straight into desktop OneNote via COM when it's installed (no sign-in), or via the Microsoft Graph OneNote API otherwise.

This file is the orientation for AI coding agents. Detailed knowledge lives in `agents/`.

## Quick Reference

| Command | Description |
|---------|-------------|
| `dotnet build SendToOneNote.slnx` | Build all three projects |
| `dotnet test` | Run unit tests (synthetic fixtures only; no network, no sign-in) |
| `dotnet run --project src/SendToOneNote` | Run the tray app |
| `STN_INTEGRATION=1 dotnet test --filter IntegrationSmokeTests` | Real Graph API smoke test (owner's machine only; needs sign-in + a "SendToOneNote Test" section) |
| `STN_INTEGRATION=1 dotnet test --filter DesktopIntegrationSmokeTests` | Real desktop OneNote smoke test (owner's machine only; needs desktop OneNote + a "SendToOneNote Test" section) |

No local infrastructure required for unit tests. The real Graph API is touched only by the gated integration test and manual E2E.

## Project Knowledge

Detailed docs in `agents/`:

### Knowledge
- [architecture.md](agents/knowledge/architecture.md) — Repo layout, the three projects, the save pipeline (watcher → parser → builder → resolver → backend seam), the desktop/Graph backends, Graph API constraints, app-data locations.
- [background.md](agents/knowledge/background.md) — Why this app exists: the new-Outlook add-in gap, the Gmail restriction, why capture is a watched folder, verified research findings with dates.

### Rules
- [csharp.md](agents/rules/csharp.md) — .NET 10 conventions, project boundaries (Core has no UI), dependency policy, TDD expectations.
- [privacy.md](agents/rules/privacy.md) — No telemetry, fixtures/local is radioactive, what's public by design (client ID) vs never committed.

## Stack Constraints

- **.NET 10**, TFM `net10.0-windows`, Windows-only, WPF. Solution file is `SendToOneNote.slnx` (XML solution format — do not add a classic `.sln`).
- **Approved dependencies:** MimeKit, AngleSharp, Microsoft.Identity.Client (+ .Broker, + .Extensions.Msal), H.NotifyIcon.Wpf, System.Drawing.Common, xUnit (+ Xunit.SkippableFact). Adding any other package is a decision, not a convenience — surface it as an Open Question first.
- **System.Drawing.Common is deliberate** (Windows-only desktop app is its supported scenario; isolated behind `ImageShrinker`). Don't propose SkiaSharp/ImageSharp migrations without a reason.
- **Graph OneNote API hard limits:** ≤4 MB per request (3.5 MB safety cap in `PagePlanner`) — the binding constraint; up to 30 binary parts per request (docs claim ~6, live service verified higher); aggressive XHTML sanitization. These shape the page-creation design; don't "simplify" them away.
- **Delegated scopes exactly:** `User.Read`, `Notes.ReadWrite`. Authority `https://login.microsoftonline.com/common`. Never request broader scopes.
- **`Backend` setting** (`auto`|`desktop`|`graph`, default `auto`): `auto` uses desktop OneNote via COM when `DesktopOneNoteProbe.IsAvailable()` succeeds, else falls back to Graph; `desktop`/`graph` force one path. Desktop mode never signs in and never talks to Graph — see `agents/knowledge/architecture.md`'s Backends section.

## Planning & Scope Rules

- v1 spec: `docs/superpowers/specs/2026-08-21-sendtoonenote-design.md`. v1 plan: `docs/superpowers/plans/2026-08-21-sendtoonenote-v1.md`. GitHub issues #1–#13 map 1:1 to v1 plan tasks; #16–#20 are v2 backlog.
- v1.1 spec: `docs/superpowers/specs/2026-08-28-desktop-onenote-backend-design.md`. v1.1 plan: `docs/superpowers/plans/2026-08-28-desktop-onenote-backend.md`. GitHub issues #21–#30 map 1:1 to its Tasks 1–10 (#31 is the owner's manual steps).
- **The active plan defines scope.** v2 candidates (attachments option, auto-send mode, .txt/image drops, right-click handler) are explicitly out of v1 — don't implement them incidentally.
- If you discover related work that *could* be done but isn't in the plan, surface it as an **Open Question** — never fold it into a task silently.

## Keeping the Agent Docs Current

Update `agents/` docs in the same change that makes them stale.

**Update triggers:**
- `agents/knowledge/architecture.md` — a project or pipeline component is added/removed/renamed, a Graph endpoint or limit changes, app-data file locations change, build/run commands change.
- `agents/knowledge/background.md` — a Microsoft platform fact this app depends on changes (add-in support for Gmail accounts, drag-out behavior, OneNote API status).
- `agents/rules/csharp.md` — a project-wide convention or the dependency list changes.
- `agents/rules/privacy.md` — anything about data handling, scopes, or what ships in the repo changes.
- `AGENTS.md` (this file) — a new file is added under `agents/`.

**Not triggers:** bug fixes that don't change a documented name/shape, new private methods, test additions, internal refactors.

When unsure, ask — a stale doc is worse than no doc.
