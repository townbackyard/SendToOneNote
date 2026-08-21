# C# conventions

## Target framework & solution

- All three projects target **`net10.0-windows`**, pinned once in `Directory.Build.props` (don't re-declare per-project).
- `<Nullable>enable</Nullable>` everywhere. Honor nullability; don't suppress with `!` unless the alternative is clearly worse — and justify it.
- Solution file is **`SendToOneNote.slnx`** (XML format). Never add a classic `.sln` alongside it.

## Project boundaries

- **`SendToOneNote.Core` contains no UI.** No WPF types, no windows, no `Application.Current`, no dialogs. If a Core component needs a decision from the user, it exposes data/events and the WPF layer asks. This is what keeps the logic testable — protect it.
- The WPF project contains only: windows/XAML, tray wiring, `SavePipeline` orchestration, and logging setup. New logic goes in Core with tests.

## Serialization

- **System.Text.Json everywhere.** There is no Newtonsoft in this project — don't introduce it (MimeKit/AngleSharp/MSAL don't need it). Graph payloads are built/parsed with `JsonSerializer`/`JsonDocument` directly; no Graph SDK.

## Records and immutability

- Data shapes (`ParsedEmail`, `ResolvedImage`, `PagePlan`, notebook models…) are `sealed record` types with positional syntax. Keep them dumb — no behavior beyond trivial derived properties.
- Mutable state lives only where mutation is the point: `AppSettings`, view-model internals.

## Async

- Everything that touches the filesystem-over-network, HTTP, or MSAL is async; no `.Result`/`.Wait()`.
- `CancellationToken ct = default` as the last parameter on public async APIs in Core.
- Fire-and-forget (`_ = SomethingAsync()`) only for genuinely optional background work (e.g., the tree-cache refresh), and the method must swallow-and-log, never throw unobserved.

## Dependencies

Approved list (see AGENTS.md): MimeKit, AngleSharp, MSAL trio, H.NotifyIcon.Wpf, System.Drawing.Common, xUnit + SkippableFact. Anything else is an Open Question first. In particular:

- **System.Drawing.Common stays** — Windows-only desktop apps are its supported scenario since .NET 6; it's isolated behind `ImageShrinker` (one file) if a swap is ever warranted.
- No logging frameworks (the tiny `FileLog` is deliberate), no DI containers (composition happens in `TrayContext` by hand — the object graph is ~6 objects), no MVVM frameworks.

## Testing (TDD)

- New Core behavior: failing test first, then implementation. Tests use the synthetic fixtures and `StubHttpHandler` — never the network, never real auth.
- Timing-sensitive watcher tests may need generous waits on CI, but never add sleeps to *implementation* code to make tests pass.
- `fixtures/local/` tests must skip cleanly when the folder is absent (CI, contributors' machines).

## Error handling

- Typed exceptions at boundaries: `EmlParseException`, `AuthRequiredException`, `OneNoteApiException(StatusCode)`. The pipeline maps them to user-facing one-liners; details go to the log file, not the toast.
- A failure must never delete or corrupt the user's .eml — the file either stays put or moves to `Failed\`.
- Partial degradation beats total failure: an image that won't download keeps its URL; an undecodable oversized image is dropped with a visible "[image omitted]" note; the page still gets created.
