# Embed Email Attachments (v2, issue #16) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Every ordinary file attachment on a dropped email becomes a real embedded file object on the OneNote page, under the header block, on both the desktop (COM) and Graph backends; oversize or over-budget files are skipped with a visible note; attached messages are listed by name only; `IncludeAttachments` (default on) and `MaxAttachmentBytes` (default 25 MB) live in settings.json.

**Architecture:** `EmlParser` now captures attachment bytes. A pure `AttachmentPlanner` applies the setting and size cap. `PageXhtmlBuilder` emits Graph-canonical `<object data-attachment=… data="name:attN">` markup after the header table. The backend seam takes a new `PageContent(Xhtml, Images, Attachments)` record. `PagePlanner` places attachments as extra binary parts after images under the existing 30-part / 3.5 MB caps. `DesktopOneNoteBackend` strips the object markup, writes each file to a per-save temp folder, and adds a `one:InsertedFile` outline before the HTML block; the temp folder is deleted in a `finally`.

**Tech Stack:** .NET 10 / C# / WPF (`net10.0-windows`), MimeKit (existing), AngleSharp (existing), COM interop via the existing hand-declared `IApplication`, xUnit + Xunit.SkippableFact (existing). **No new packages.**

**Spec:** `docs/superpowers/specs/2026-09-15-attachments-design.md` (binding). Backend seam and page XML background: `docs/superpowers/specs/2026-08-28-desktop-onenote-backend-design.md`.

## Global Constraints

- .NET 10, TFM `net10.0-windows`; `SendToOneNote.Core` contains no UI types. Nullable enabled everywhere.
- **No new NuGet packages.**
- **All COM calls run on the single `StaComWorker` thread**, including the temp-file write that precedes `UpdatePageContent`.
- Settings: `IncludeAttachments` bool default `true`; `MaxAttachmentBytes` long default `26_214_400`. Per-attachment cap on both backends.
- Graph caps unchanged and binding: `PagePlanner.MaxBinaryPartsPerRequest = 30`, `MaxRequestBytes = 3_500_000`. Images are placed first; attachments fill what remains, in email order; the rest are dropped with a note.
- Attachment part names are `att1…attN` (1-based, email order). Image part names stay `img0…`.
- The `<object>` element in the page XHTML is always written with an explicit end tag (`<object …></object>`), never self-closing, because AngleSharp's HTML parser would otherwise swallow the rest of the page into it.
- Attached messages (`message/rfc822`) are never embedded; they appear in the header "Attachments" row and get a grey note.
- Temp files: `%TEMP%\SendToOneNote\<Guid>\<sanitised name>`, created only when there is at least one attachment, deleted in a `finally` after `UpdatePageContent` (success or failure).
- Fixtures are synthetic (`example.com` / `example.net` / `example.org`, invented content). `fixtures/local/` is gitignored — never commit, quote, or copy anything from it; local-fixture tests assert counts only.
- Tests touching real OneNote or Graph are gated on `STN_INTEGRATION=1` and must SKIP elsewhere.
- Commit style: `feat:` / `fix:` / `test:` / `docs:` prefixes, present tense; end every commit with the session's `Co-Authored-By` and `Claude-Session` trailers.
- Work on `main`; push after each task; CI (`.github/workflows/build.yml`) must stay green. Verify with `dotnet test --nologo -v q` (baseline at plan time: 98 passed, 3 skipped).
- Update `agents/` docs in the same change that makes them stale (Task 8 does the bulk; earlier tasks do not touch docs).

## Context for a fresh session (read before Task 1)

**Repo state at plan time:** `main` @ `ebb9abf`, CI green. GitHub issues #34–#41 map 1:1 to Tasks 1–8; close each when its task lands. Read `AGENTS.md`, `agents/knowledge/architecture.md`, `agents/rules/csharp.md`, and the spec above first.

**Existing signatures this plan builds on (exact):**

```csharp
// SendToOneNote.Core.Email
public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);
public sealed record ParsedEmail(string Subject, string From, string To, string? Cc, DateTimeOffset? SentDate,
    string? HtmlBody, string? TextBody, IReadOnlyList<InlineImage> InlineImages, IReadOnlyList<string> AttachmentNames);
public static class EmlParser { public static ParsedEmail Parse(Stream emlStream); }

// SendToOneNote.Core.Pages
public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data, int Width = 0, int Height = 0);
public static class PageXhtmlBuilder { public static string Build(ParsedEmail email); }
public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts, IReadOnlyList<AppendPlan> Appends)
{ public IReadOnlyList<string> DroppedPartNames { get; init; } = []; }
public static class PagePlanner { public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images); }
public static class DataUriInliner { public static string Inline(string xhtml, IReadOnlyList<ResolvedImage> images); }
public sealed class ImageResolver { public Task<ImageResolution> ResolveWithReportAsync(string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default); }
public sealed record ImageResolution(string Xhtml, IReadOnlyList<ResolvedImage> Images, IReadOnlyList<ImageDecision> Decisions);

// SendToOneNote.Core.Backends
public interface IOneNoteBackend
{
    string Name { get; }
    Task<NotebookTree> GetTreeAsync(CancellationToken ct = default);
    Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml, IReadOnlyList<ResolvedImage> images, CancellationToken ct = default);
}
public sealed class GraphBackend(OneNoteClient client) : IOneNoteBackend;
public sealed class DesktopOneNoteBackend(StaComWorker worker, Func<IApplication>? applicationFactory = null) : IOneNoteBackend;

// SendToOneNote.Core.Desktop
public static class OneNotePageXmlBuilder
{
    public static string Build(string pageId, string title, string html);
    public static string CdataEscape(string s);
    public static string ExtractTitle(string pageXhtml);
}
public static class OneNoteConstants { public const string ProgId; public const int NpsDefault, Xs2013, PiBinaryData; public static readonly string Namespace2013; }

// SendToOneNote.Core.OneNote
public sealed record CreatedPage(string Id, string? ClientUrl, string? WebUrl);
public sealed class OneNoteClient { public Task<CreatedPage> CreatePageAsync(string sectionId, PagePlan plan, CancellationToken ct = default); }

// SendToOneNote.Core.Storage
public sealed class AppSettings { string? DropFolder; string? ClientIdOverride; bool DeleteOnSuccess = true; List<string> RecentSectionIds; string Backend = "auto"; bool ImageDiagnostics; List<string> RecentDesktopSectionIds; }
```

**Test helpers that already exist:** `Fixtures.Open(name)` opens `fixtures/synthetic/<name>`; `FakeOneNoteApplication : IApplication` records `CreatedPages`, `UpdatedPageXml`, and throws `ThrowOnUpdate` from `UpdatePageContent`; `StubHttpHandler(Func<HttpRequestMessage, HttpResponseMessage>)` records `Requests`; `StaComWorker` is `IDisposable` with `Task<T> RunAsync(Func<T>)`.

## File map

| File | Responsibility | Task |
|---|---|---|
| `src/SendToOneNote.Core/Email/ParsedEmail.cs` | `EmailAttachment`; `ParsedEmail` gains `Attachments`, `AttachedMessageNames`; `AttachmentNames` becomes derived | 1 |
| `src/SendToOneNote.Core/Email/EmlParser.cs` | Decode attachment bytes; list attached messages | 1 |
| `fixtures/synthetic/pdf-attachment.eml`, `attached-message.eml` | New synthetic fixtures | 1 |
| `src/SendToOneNote.Core/Storage/AppSettings.cs` | Two new settings | 2 |
| `src/SendToOneNote.Core/Pages/AttachmentPlanner.cs` | `PlannedAttachment`, `OmittedAttachment`, `AttachmentPlan`, `AttachmentPlanner.Plan` | 2 |
| `src/SendToOneNote.Core/Pages/AttachmentMarkup.cs` | The one place that knows the `<object>` shape, the omitted-note shape, the size formatter, and the regex that finds/strips objects | 3 |
| `src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs` | `Build(email, plan)` emits the attachment block after the header table | 3 |
| `src/SendToOneNote.Core/Pages/PageContent.cs` | `PageContent` record | 4 |
| `src/SendToOneNote.Core/Backends/IOneNoteBackend.cs`, `GraphBackend.cs`, `DesktopOneNoteBackend.cs` | Seam takes `PageContent` | 4, 6 |
| `src/SendToOneNote.Core/Pages/PagePlanner.cs` | `Plan(PageContent)`; attachments as parts under the caps | 4, 5 |
| `src/SendToOneNote.Core/Desktop/AttachmentTempFolder.cs` | Per-save temp folder: sanitise names, write bytes, delete on dispose | 6 |
| `src/SendToOneNote.Core/Desktop/OneNotePageXmlBuilder.cs` | `InsertedFile` record; four-argument `Build` with the file outline | 6 |
| `src/SendToOneNote/SavePipeline.cs` | Planner → builder → resolver → `PageContent` → backend; logs omissions | 4 (mechanical), 7 |
| `tests/SendToOneNote.Tests/*` | Tests per task; both gated smoke tests gain a PDF | 1–7 |
| `agents/knowledge/architecture.md`, `agents/rules/privacy.md`, `README.md`, `docs/e2e-checklist.md`, `AGENTS.md` | Docs | 8 |

---

### Task 1: Attachment bytes in the data model and parser

**Files:**
- Modify: `src/SendToOneNote.Core/Email/ParsedEmail.cs`
- Modify: `src/SendToOneNote.Core/Email/EmlParser.cs`
- Create: `fixtures/synthetic/pdf-attachment.eml`
- Create: `fixtures/synthetic/attached-message.eml`
- Modify: `tests/SendToOneNote.Tests/Fixtures.cs` (add the two names to `FixtureExists`)
- Modify: `tests/SendToOneNote.Tests/EmlParserTests.cs`
- Modify: `tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs` (only the `Email(...)` helper — the constructor shape changes)
- Modify: `tests/SendToOneNote.Tests/LocalFixtureTests.cs`

**Interfaces:**
- Produces:
  ```csharp
  public sealed record EmailAttachment(string FileName, string ContentType, byte[] Data);
  public sealed record ParsedEmail(string Subject, string From, string To, string? Cc, DateTimeOffset? SentDate,
      string? HtmlBody, string? TextBody, IReadOnlyList<InlineImage> InlineImages,
      IReadOnlyList<EmailAttachment> Attachments, IReadOnlyList<string> AttachedMessageNames)
  { public IReadOnlyList<string> AttachmentNames { get; } }   // files then messages
  ```

- [ ] **Step 1: Create the two synthetic fixtures**

`fixtures/synthetic/pdf-attachment.eml` (the base64 is a 218-byte hand-written PDF skeleton: catalog, one empty 200×100 page, no xref; invented content only):

```
From: Billing <billing@example.com>
To: pat@example.net
Subject: Your invoice INV-0001
Date: Tue, 15 Sep 2026 10:00:00 +0000
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="mix1"

--mix1
Content-Type: text/html; charset="utf-8"
Content-ID: <BODY0001@example>

<html><body><p>Thanks for your order. The invoice is attached.</p></body></html>
--mix1
Content-Type: application/pdf; name="invoice.pdf"
Content-Disposition: attachment; filename="invoice.pdf"
Content-ID: <ATTACH0001@example>
Content-Transfer-Encoding: base64

JVBERi0xLjQKMSAwIG9iago8PCAvVHlwZSAvQ2F0YWxvZyAvUGFnZXMgMiAwIFIgPj4KZW5kb2Jq
CjIgMCBvYmoKPDwgL1R5cGUgL1BhZ2VzIC9LaWRzIFszIDAgUl0gL0NvdW50IDEgPj4KZW5kb2Jq
CjMgMCBvYmoKPDwgL1R5cGUgL1BhZ2UgL1BhcmVudCAyIDAgUiAvTWVkaWFCb3ggWzAgMCAyMDAg
MTAwXSA+PgplbmRvYmoKdHJhaWxlcgo8PCAvUm9vdCAxIDAgUiA+PgolJUVPRgo=
--mix1--
```

`fixtures/synthetic/attached-message.eml`:

```
From: Sender <sender@example.com>
To: pat@example.net
Subject: Forwarding the note
Date: Tue, 15 Sep 2026 10:30:00 +0000
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="mix2"

--mix2
Content-Type: text/plain; charset="utf-8"

See the attached message.
--mix2
Content-Type: message/rfc822; name="Original note.eml"
Content-Disposition: attachment; filename="Original note.eml"

From: Someone <someone@example.org>
To: sender@example.com
Subject: Original note about the picnic
Date: Mon, 14 Sep 2026 15:00:00 +0000
MIME-Version: 1.0
Content-Type: text/plain; charset="utf-8"

Bring napkins.
--mix2--
```

Add both names to the `[InlineData]` list of `FixtureTests.FixtureExists` in `tests/SendToOneNote.Tests/Fixtures.cs`.

- [ ] **Step 2: Write the failing parser tests**

Append to `tests/SendToOneNote.Tests/EmlParserTests.cs`:

```csharp
    [Fact]
    public void CapturesAttachmentBytesAndContentType()
    {
        var e = EmlParser.Parse(Fixtures.Open("pdf-attachment.eml"));
        var a = Assert.Single(e.Attachments);
        Assert.Equal("invoice.pdf", a.FileName);
        Assert.Equal("application/pdf", a.ContentType);
        Assert.True(a.Data.Length > 100);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(a.Data, 0, 4)); // decoded, not base64 text
        Assert.Empty(e.AttachedMessageNames);
        Assert.Equal(["invoice.pdf"], e.AttachmentNames);
        Assert.Empty(e.InlineImages);
    }

    [Fact]
    public void AttachedMessageIsListedByNameNotEmbedded()
    {
        var e = EmlParser.Parse(Fixtures.Open("attached-message.eml"));
        Assert.Empty(e.Attachments);
        Assert.Equal(["Original note about the picnic"], e.AttachedMessageNames);
        Assert.Equal(["Original note about the picnic"], e.AttachmentNames);
    }

    [Fact]
    public void InlineCidImageIsNotAnAttachment()
    {
        var e = EmlParser.Parse(Fixtures.Open("inline-cid-image.eml"));
        Assert.Empty(e.Attachments);
        Assert.Empty(e.AttachedMessageNames);
    }

    [Fact]
    public void AttachmentWithoutContentTypeFallsBackToOctetStream()
    {
        var raw = """
            From: sender@example.com
            To: recipient@example.com
            Subject: Blob
            MIME-Version: 1.0
            Content-Type: multipart/mixed; boundary="mixed9"

            --mixed9
            Content-Type: text/plain

            body
            --mixed9
            Content-Disposition: attachment; filename="data.bin"
            Content-Transfer-Encoding: base64

            AQID
            --mixed9--
            """;
        var e = EmlParser.Parse(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(raw)));
        var a = Assert.Single(e.Attachments);
        Assert.Equal("data.bin", a.FileName);
        Assert.Equal("application/octet-stream", a.ContentType);
        Assert.Equal(new byte[] { 1, 2, 3 }, a.Data);
    }
```

Update the helper at the top of `tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs` so the file still compiles (the constructor gains a parameter; `attachments` here are file names):

```csharp
    private static ParsedEmail Email(string? html = null, string? text = null,
        string subject = "S & T <test>", IReadOnlyList<string>? attachments = null) =>
        new(subject, "a@b.c", "d@e.f", null,
            new DateTimeOffset(2026, 8, 20, 18, 0, 35, TimeSpan.Zero),
            html, text, [],
            (attachments ?? []).Select(n => new EmailAttachment(n, "application/octet-stream", [1, 2, 3])).ToList(),
            []);
```

Append to `tests/SendToOneNote.Tests/LocalFixtureTests.cs` (counts only — never content):

```csharp
    [SkippableTheory]
    [MemberData(nameof(LocalEmls))]
    public void RealAttachmentsCarryBytes(string? path)
    {
        Skip.If(path is null, "fixtures/local not present (CI or fresh clone)");
        using var s = File.OpenRead(path!);
        var e = EmlParser.Parse(s);
        Assert.All(e.Attachments, a =>
        {
            Assert.False(string.IsNullOrWhiteSpace(a.FileName));
            Assert.False(string.IsNullOrWhiteSpace(a.ContentType));
            Assert.True(a.Data.Length > 0);
        });
        Assert.Equal(e.Attachments.Count + e.AttachedMessageNames.Count, e.AttachmentNames.Count);
    }
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet build SendToOneNote.slnx --nologo -v q`
Expected: compile errors — `EmailAttachment`, `Attachments`, `AttachedMessageNames` do not exist.

- [ ] **Step 4: Implement the model**

Replace the record in `src/SendToOneNote.Core/Email/ParsedEmail.cs`:

```csharp
namespace SendToOneNote.Core.Email;

public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);

/// <summary>A file the sender attached (not an inline cid: image). Bytes are decoded.</summary>
public sealed record EmailAttachment(string FileName, string ContentType, byte[] Data);

public sealed record ParsedEmail(
    string Subject,
    string From,
    string To,
    string? Cc,
    DateTimeOffset? SentDate,
    string? HtmlBody,
    string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<EmailAttachment> Attachments,
    IReadOnlyList<string> AttachedMessageNames)
{
    /// <summary>Header-row list: file attachments first, then attached messages (never embedded).</summary>
    public IReadOnlyList<string> AttachmentNames =>
        [.. Attachments.Select(a => a.FileName), .. AttachedMessageNames];
}

public sealed class EmlParseException(string message, Exception? inner = null)
    : Exception(message, inner);
```

- [ ] **Step 5: Implement the parser change**

In `src/SendToOneNote.Core/Email/EmlParser.cs`, replace the `var attachments = msg.Attachments.OfType<MimePart>() … .ToList();` statement with:

```csharp
            var attachments = new List<EmailAttachment>();
            var messageNames = new List<string>();
            foreach (var entity in msg.Attachments)
            {
                switch (entity)
                {
                    case MimePart p:
                        if (p.ContentId is not null && referencedCids.Contains(p.ContentId.Trim('<', '>')))
                            continue; // inline image referenced by the body, captured above
                        byte[] data = [];
                        if (p.Content is not null)
                        {
                            using var ms = new MemoryStream();
                            p.Content.DecodeTo(ms);
                            data = ms.ToArray();
                        }
                        var mime = p.ContentType?.MimeType;
                        attachments.Add(new EmailAttachment(
                            p.FileName ?? "attachment",
                            string.IsNullOrWhiteSpace(mime) ? "application/octet-stream" : mime,
                            data));
                        break;
                    case MessagePart mp:
                        // Nested .eml: listed by name, never embedded (no useful handler for new-Outlook users).
                        var subject = mp.Message?.Subject;
                        messageNames.Add(!string.IsNullOrWhiteSpace(subject) ? subject
                            : mp.ContentDisposition?.FileName ?? "attached message");
                        break;
                }
            }
```

and change the constructor call's last two arguments to `Attachments: attachments, AttachedMessageNames: messageNames`.

Note for the implementer: `MimeMessage.Attachments` yields every entity whose `Content-Disposition` is `attachment`, including `MessagePart`. If the message-fixture test fails because MimeKit does not surface the `MessagePart`, do not change the fixture's disposition header; instead also scan `msg.BodyParts.OfType<MessagePart>()` and report the finding in your summary.

- [ ] **Step 6: Run the tests to verify they pass**

Run: `dotnet test --nologo -v q`
Expected: all pass (102 + the new local-fixture theory on the owner's machine; on CI the local theories skip).

- [ ] **Step 7: Commit**

```bash
git add fixtures/synthetic/pdf-attachment.eml fixtures/synthetic/attached-message.eml src/SendToOneNote.Core/Email tests/SendToOneNote.Tests/Fixtures.cs tests/SendToOneNote.Tests/EmlParserTests.cs tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs tests/SendToOneNote.Tests/LocalFixtureTests.cs
git commit -m "feat: capture attachment bytes; list attached messages by name (#16)"
git push
```

---

### Task 2: Settings and AttachmentPlanner

**Files:**
- Modify: `src/SendToOneNote.Core/Storage/AppSettings.cs`
- Create: `src/SendToOneNote.Core/Pages/AttachmentPlanner.cs`
- Modify: `tests/SendToOneNote.Tests/StorageTests.cs`
- Create: `tests/SendToOneNote.Tests/AttachmentPlannerTests.cs`

**Interfaces:**
- Consumes: `ParsedEmail`, `EmailAttachment` (Task 1).
- Produces:
  ```csharp
  // AppSettings
  public bool IncludeAttachments { get; set; } = true;
  public long MaxAttachmentBytes { get; set; } = 26_214_400;
  // SendToOneNote.Core.Pages
  public sealed record PlannedAttachment(string PartName, EmailAttachment Attachment);
  public sealed record OmittedAttachment(string FileName, long Bytes, string Reason);
  public sealed record AttachmentPlan(IReadOnlyList<PlannedAttachment> Embedded, IReadOnlyList<OmittedAttachment> Omitted)
  { public static AttachmentPlan Empty { get; } }
  public static class AttachmentPlanner
  {
      public const string OverSizeLimit = "over the size limit";
      public static AttachmentPlan Plan(ParsedEmail email, bool includeAttachments, long maxAttachmentBytes);
  }
  ```

- [ ] **Step 1: Write the failing tests**

Create `tests/SendToOneNote.Tests/AttachmentPlannerTests.cs`:

```csharp
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class AttachmentPlannerTests
{
    private static ParsedEmail Email(params EmailAttachment[] attachments) =>
        new("s", "a@b.c", "d@e.f", null, null, "<p>hi</p>", null, [], attachments, ["Fwd: nested"]);

    private static EmailAttachment File(string name, int bytes) =>
        new(name, "application/octet-stream", new byte[bytes]);

    [Fact]
    public void DisabledYieldsEmptyPlan()
    {
        var plan = AttachmentPlanner.Plan(Email(File("a.pdf", 10)), includeAttachments: false, maxAttachmentBytes: 1000);
        Assert.Same(AttachmentPlan.Empty, plan);
        Assert.Empty(plan.Embedded);
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void NumbersEmbeddedAttachmentsInEmailOrder()
    {
        var plan = AttachmentPlanner.Plan(Email(File("a.pdf", 10), File("b.docx", 20)), true, 1000);
        Assert.Equal(["att1", "att2"], plan.Embedded.Select(p => p.PartName));
        Assert.Equal(["a.pdf", "b.docx"], plan.Embedded.Select(p => p.Attachment.FileName));
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void ExactlyAtCapIsEmbeddedOneByteOverIsOmitted()
    {
        var plan = AttachmentPlanner.Plan(Email(File("ok.pdf", 100), File("big.pdf", 101)), true, 100);
        Assert.Equal(["att1"], plan.Embedded.Select(p => p.PartName));
        var omitted = Assert.Single(plan.Omitted);
        Assert.Equal("big.pdf", omitted.FileName);
        Assert.Equal(101, omitted.Bytes);
        Assert.Equal(AttachmentPlanner.OverSizeLimit, omitted.Reason);
    }

    [Fact]
    public void NumberingSkipsOmittedFiles()
    {
        // att numbers count only embedded files, so parts are contiguous.
        var plan = AttachmentPlanner.Plan(Email(File("big.pdf", 500), File("small.pdf", 5)), true, 100);
        Assert.Equal(["att1"], plan.Embedded.Select(p => p.PartName));
        Assert.Equal("small.pdf", plan.Embedded[0].Attachment.FileName);
    }
}
```

Extend `NewSettingsHaveSafeDefaultsAndRoundTrip` in `tests/SendToOneNote.Tests/StorageTests.cs`: after `Assert.Empty(s.RecentDesktopSectionIds);` add

```csharp
        Assert.True(s.IncludeAttachments);
        Assert.Equal(26_214_400, s.MaxAttachmentBytes);
```

after `s.RecentDesktopSectionIds.Add("{S1}");` add `s.IncludeAttachments = false; s.MaxAttachmentBytes = 1000;` and at the end add

```csharp
        Assert.False(s2.IncludeAttachments);
        Assert.Equal(1000, s2.MaxAttachmentBytes);
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet build SendToOneNote.slnx --nologo -v q`
Expected: compile errors for `AttachmentPlanner`, `IncludeAttachments`.

- [ ] **Step 3: Implement**

Add to `src/SendToOneNote.Core/Storage/AppSettings.cs` after `RecentDesktopSectionIds`:

```csharp
    /// <summary>Embed file attachments on the page (default on). Off = names only, as before v2.</summary>
    public bool IncludeAttachments { get; set; } = true;
    /// <summary>Per-attachment cap; larger files are skipped with a note on the page. Default 25 MB.</summary>
    public long MaxAttachmentBytes { get; set; } = 26_214_400;
```

Create `src/SendToOneNote.Core/Pages/AttachmentPlanner.cs`:

```csharp
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

/// <summary>An attachment that will be embedded; PartName is "att1", "att2", … (Graph part name / desktop file index).</summary>
public sealed record PlannedAttachment(string PartName, EmailAttachment Attachment);

public sealed record OmittedAttachment(string FileName, long Bytes, string Reason);

public sealed record AttachmentPlan(IReadOnlyList<PlannedAttachment> Embedded, IReadOnlyList<OmittedAttachment> Omitted)
{
    public static AttachmentPlan Empty { get; } = new([], []);
}

/// <summary>Applies the IncludeAttachments switch and the per-file size cap. Pure; backend-agnostic.</summary>
public static class AttachmentPlanner
{
    public const string OverSizeLimit = "over the size limit";

    public static AttachmentPlan Plan(ParsedEmail email, bool includeAttachments, long maxAttachmentBytes)
    {
        if (!includeAttachments) return AttachmentPlan.Empty;

        var embedded = new List<PlannedAttachment>();
        var omitted = new List<OmittedAttachment>();
        foreach (var a in email.Attachments)
        {
            if (a.Data.LongLength > maxAttachmentBytes)
                omitted.Add(new OmittedAttachment(a.FileName, a.Data.LongLength, OverSizeLimit));
            else
                embedded.Add(new PlannedAttachment($"att{embedded.Count + 1}", a));
        }
        return new AttachmentPlan(embedded, omitted);
    }
}
```

- [ ] **Step 4: Run tests**

Run: `dotnet test --nologo -v q`
Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add src/SendToOneNote.Core/Storage/AppSettings.cs src/SendToOneNote.Core/Pages/AttachmentPlanner.cs tests/SendToOneNote.Tests/StorageTests.cs tests/SendToOneNote.Tests/AttachmentPlannerTests.cs
git commit -m "feat: IncludeAttachments/MaxAttachmentBytes settings and AttachmentPlanner (#16)"
git push
```

---

### Task 3: Attachment markup in PageXhtmlBuilder

**Files:**
- Create: `src/SendToOneNote.Core/Pages/AttachmentMarkup.cs`
- Modify: `src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs`
- Create: `tests/SendToOneNote.Tests/AttachmentMarkupTests.cs`
- Modify: `tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs`

**Interfaces:**
- Consumes: `AttachmentPlan`, `PlannedAttachment`, `OmittedAttachment` (Task 2).
- Produces:
  ```csharp
  public static class AttachmentMarkup
  {
      public const string GraphBudgetReason = "too large for OneNote online";
      public static string ObjectElement(PlannedAttachment a);                 // <object data-attachment=… data="name:attN" type=…></object>
      public static string OmittedNote(string fileName, long bytes, string reason);
      public static string AttachedMessageNote(string name);
      public static string FormatSize(long bytes);                            // "31.4 MB" / "412 KB"
      public static Regex ObjectRegex(string partName);                       // whole element, any attribute order, "/>" or "></object>"
      public static string StripAll(string xhtml);                            // remove every attachment object element
  }
  public static class PageXhtmlBuilder
  {
      public static string Build(ParsedEmail email);                          // == Build(email, AttachmentPlan.Empty)
      public static string Build(ParsedEmail email, AttachmentPlan plan);
  }
  ```

- [ ] **Step 1: Write the failing tests**

Create `tests/SendToOneNote.Tests/AttachmentMarkupTests.cs`:

```csharp
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class AttachmentMarkupTests
{
    private static PlannedAttachment Att(string part, string name, string type = "application/pdf") =>
        new(part, new EmailAttachment(name, type, [1, 2, 3]));

    [Fact]
    public void ObjectElementHasExplicitEndTagAndEncodedName()
    {
        var html = AttachmentMarkup.ObjectElement(Att("att1", "a & b <c>.pdf"));
        Assert.Equal("<object data-attachment=\"a &amp; b &lt;c&gt;.pdf\" data=\"name:att1\" type=\"application/pdf\"></object>", html);
    }

    [Theory]
    [InlineData(512, "512 bytes")]
    [InlineData(1_536, "1.5 KB")]
    [InlineData(412_000, "402.3 KB")]
    [InlineData(32_900_000, "31.4 MB")]
    public void FormatsSizes(long bytes, string expected) =>
        Assert.Equal(expected, AttachmentMarkup.FormatSize(bytes));

    [Fact]
    public void OmittedNoteIsGreyParagraph()
    {
        var html = AttachmentMarkup.OmittedNote("scan.tif", 32_900_000, "over the size limit");
        Assert.Equal("<p style=\"color:#999999\">[attachment omitted: scan.tif, 31.4 MB, over the size limit]</p>", html);
    }

    [Fact]
    public void AttachedMessageNoteIsGreyParagraph() =>
        Assert.Equal("<p style=\"color:#999999\">[attached message not embedded: Re: order]</p>",
            AttachmentMarkup.AttachedMessageNote("Re: order"));

    [Theory]
    [InlineData("<object data-attachment=\"a.pdf\" data=\"name:att1\" type=\"application/pdf\"></object>")]
    [InlineData("<object type=\"application/pdf\" data=\"name:att1\" data-attachment=\"a.pdf\" />")]
    [InlineData("<object data=\"name:att1\" data-attachment=\"a.pdf\">\n</object>")]
    public void ObjectRegexMatchesWholeElementInAnyForm(string element)
    {
        var xhtml = $"<body><p>before</p>{element}<p>after</p></body>";
        Assert.Equal("<body><p>before</p><p>after</p></body>", AttachmentMarkup.ObjectRegex("att1").Replace(xhtml, ""));
    }

    [Fact]
    public void ObjectRegexDoesNotMatchOtherParts()
    {
        var xhtml = "<object data-attachment=\"a.pdf\" data=\"name:att10\" type=\"x\"></object>";
        Assert.False(AttachmentMarkup.ObjectRegex("att1").IsMatch(xhtml));
    }

    [Fact]
    public void StripAllRemovesEveryAttachmentObjectAndNothingElse()
    {
        var xhtml = "<div><object data-attachment=\"a.pdf\" data=\"name:att1\" type=\"x\"></object>" +
                    "<object data-attachment=\"b.pdf\" data=\"name:att2\" type=\"y\"></object></div><img src=\"name:img0\"/>";
        Assert.Equal("<div></div><img src=\"name:img0\"/>", AttachmentMarkup.StripAll(xhtml));
    }
}
```

Append to `tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs` (add `using SendToOneNote.Core.Email;` is already present):

```csharp
    private static ParsedEmail EmailWithMessages(params string[] messageNames) =>
        new("s", "a@b.c", "d@e.f", null, null, "<p>hi</p>", null, [], [], messageNames);

    [Fact]
    public void EmptyPlanIsByteIdenticalToSingleArgumentBuild()
    {
        var email = Email(html: "<p>hi</p>", attachments: ["report.pdf"]);
        Assert.Equal(PageXhtmlBuilder.Build(email), PageXhtmlBuilder.Build(email, AttachmentPlan.Empty));
        Assert.DoesNotContain("stn-attachments", PageXhtmlBuilder.Build(email));
    }

    [Fact]
    public void EmbeddedAttachmentsAppearAfterHeaderBeforeRule()
    {
        var email = Email(html: "<p>body</p>", attachments: ["report.pdf"]);
        var plan = new AttachmentPlan([new PlannedAttachment("att1", email.Attachments[0])], []);
        var x = PageXhtmlBuilder.Build(email, plan);
        var table = x.IndexOf("</table>", StringComparison.Ordinal);
        var obj = x.IndexOf("<object data-attachment=\"report.pdf\" data=\"name:att1\" type=\"application/octet-stream\"></object>", StringComparison.Ordinal);
        var hr = x.IndexOf("<hr/>", StringComparison.Ordinal);
        Assert.True(table > 0 && obj > table && hr > obj, x);
        Assert.Contains("<div class=\"stn-attachments\">", x);
    }

    [Fact]
    public void OmittedAndAttachedMessageNotesFollowTheObjects()
    {
        var email = new ParsedEmail("s", "a@b.c", "d@e.f", null, null, "<p>hi</p>", null, [],
            [new EmailAttachment("ok.pdf", "application/pdf", [1]), new EmailAttachment("scan.tif", "image/tiff", [1])],
            ["Re: order"]);
        var plan = new AttachmentPlan(
            [new PlannedAttachment("att1", email.Attachments[0])],
            [new OmittedAttachment("scan.tif", 32_900_000, "over the size limit")]); // scan.tif omitted by the planner
        var x = PageXhtmlBuilder.Build(email, plan);
        var obj = x.IndexOf("name:att1", StringComparison.Ordinal);
        var omitted = x.IndexOf("[attachment omitted: scan.tif, 31.4 MB, over the size limit]", StringComparison.Ordinal);
        var msg = x.IndexOf("[attached message not embedded: Re: order]", StringComparison.Ordinal);
        Assert.True(obj > 0 && omitted > obj && msg > omitted, x);
        Assert.Contains("Attachments</td><td>ok.pdf; scan.tif; Re: order", x); // header row still lists all names
    }

    [Fact]
    public void AttachedMessageOnlyStillGetsANote()
    {
        var x = PageXhtmlBuilder.Build(EmailWithMessages("Fwd: hello"), AttachmentPlan.Empty);
        Assert.Contains("[attached message not embedded: Fwd: hello]", x);
    }

    [Fact]
    public async Task ObjectSurvivesXhtmlNormalizationWithoutSwallowingTheBody()
    {
        // ImageResolver re-serializes the page through AngleSharp; a self-closing <object/> would
        // become an open element containing the whole body. With an explicit end tag the body follows it.
        var email = Email(html: "<p>body text</p>", attachments: ["report.pdf"]);
        var plan = new AttachmentPlan([new PlannedAttachment("att1", email.Attachments[0])], []);
        var r = await new ImageResolver().ResolveWithReportAsync(PageXhtmlBuilder.Build(email, plan), []);
        var m = AttachmentMarkup.ObjectRegex("att1").Match(r.Xhtml);
        Assert.True(m.Success, r.Xhtml);
        Assert.Contains("data-attachment=\"report.pdf\"", m.Value);
        Assert.Contains("<p>body text</p>", r.Xhtml[(m.Index + m.Length)..]);
    }
```

Add `using SendToOneNote.Core.Pages;` at the top if not present (it is).

- [ ] **Step 2: Run to verify failure**

Run: `dotnet build SendToOneNote.slnx --nologo -v q`
Expected: compile errors for `AttachmentMarkup` and the two-argument `Build`.

- [ ] **Step 3: Implement AttachmentMarkup**

Create `src/SendToOneNote.Core/Pages/AttachmentMarkup.cs`:

```csharp
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

/// <summary>
/// The one place that knows what attachment markup looks like in the page XHTML. The builder emits it,
/// PagePlanner replaces it when Graph's budget is exceeded, and the desktop backend strips it (an
/// HTMLBlock cannot host a file; the file goes in a one:InsertedFile outline instead).
/// </summary>
public static class AttachmentMarkup
{
    public const string GraphBudgetReason = "too large for OneNote online";
    private const string NoteStyle = "color:#999999";

    /// <summary>Graph-canonical embedded-file element. Explicit end tag: AngleSharp's HTML parser treats a
    /// self-closing &lt;object/&gt; as an open element and would swallow the rest of the page into it.</summary>
    public static string ObjectElement(PlannedAttachment a) =>
        $"<object data-attachment=\"{WebUtility.HtmlEncode(a.Attachment.FileName)}\" " +
        $"data=\"name:{a.PartName}\" type=\"{WebUtility.HtmlEncode(a.Attachment.ContentType)}\"></object>";

    public static string OmittedNote(string fileName, long bytes, string reason) =>
        $"<p style=\"{NoteStyle}\">[attachment omitted: {WebUtility.HtmlEncode(fileName)}, {FormatSize(bytes)}, {WebUtility.HtmlEncode(reason)}]</p>";

    public static string AttachedMessageNote(string name) =>
        $"<p style=\"{NoteStyle}\">[attached message not embedded: {WebUtility.HtmlEncode(name)}]</p>";

    public static string FormatSize(long bytes) => bytes switch
    {
        >= 1_048_576 => (bytes / 1_048_576d).ToString("0.#", CultureInfo.InvariantCulture) + " MB",
        >= 1_024 => (bytes / 1_024d).ToString("0.#", CultureInfo.InvariantCulture) + " KB",
        _ => bytes.ToString(CultureInfo.InvariantCulture) + " bytes"
    };

    /// <summary>Matches the whole element for one part regardless of attribute order, whether serialized
    /// as &lt;object …/&gt; or &lt;object …&gt;&lt;/object&gt; (AngleSharp may emit either).</summary>
    public static Regex ObjectRegex(string partName) =>
        new($"<object\\b[^>]*\\bdata=\"name:{Regex.Escape(partName)}\"[^>]*(?:/>|>\\s*</object>)", RegexOptions.Singleline);

    private static readonly Regex AnyObject =
        new("<object\\b[^>]*\\bdata=\"name:att\\d+\"[^>]*(?:/>|>\\s*</object>)", RegexOptions.Singleline | RegexOptions.Compiled);

    public static string StripAll(string xhtml) => AnyObject.Replace(xhtml, "");
}
```

Check against the `FormatsSizes` theory: 412,000 / 1024 = 402.34 → "402.3 KB"; 32,900,000 / 1,048,576 = 31.38 → "31.4 MB"; 1,536 / 1024 = 1.5 → "1.5 KB". If a value disagrees, fix the test expectation to the true arithmetic, not the formatter.

- [ ] **Step 4: Implement the builder change**

In `src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs`:

```csharp
    public static string Build(ParsedEmail email) => Build(email, AttachmentPlan.Empty);

    public static string Build(ParsedEmail email, AttachmentPlan plan)
    {
        var sb = new StringBuilder();
        sb.Append("<html><head><title>")
          .Append(WebUtility.HtmlEncode(email.Subject))
          .Append("</title></head><body>");
        AppendHeaderTable(sb, email);
        AppendAttachments(sb, email, plan);
        sb.Append("<hr/><div>");
        if (email.HtmlBody is not null)
            sb.Append(email.HtmlBody); // normalized to XHTML later by ImageResolver
        else
            AppendTextBody(sb, email.TextBody ?? "");
        sb.Append("</div></body></html>");
        return sb.ToString();
    }
```

Move the `<hr/>` out of `AppendHeaderTable` (its last line becomes `sb.Append("</table>");`) and add:

```csharp
    // Between the header table and the rule: file objects, then notes. Omitted entirely when there is
    // nothing to say, so pre-v2 output is byte-identical.
    private static void AppendAttachments(StringBuilder sb, ParsedEmail e, AttachmentPlan plan)
    {
        if (plan.Embedded.Count == 0 && plan.Omitted.Count == 0 && e.AttachedMessageNames.Count == 0) return;
        sb.Append("<div class=\"stn-attachments\">");
        foreach (var a in plan.Embedded) sb.Append(AttachmentMarkup.ObjectElement(a));
        foreach (var o in plan.Omitted) sb.Append(AttachmentMarkup.OmittedNote(o.FileName, o.Bytes, o.Reason));
        foreach (var m in e.AttachedMessageNames) sb.Append(AttachmentMarkup.AttachedMessageNote(m));
        sb.Append("</div>");
    }
```

- [ ] **Step 5: Run tests**

Run: `dotnet test --nologo -v q`
Expected: all pass, including `ObjectSurvivesXhtmlNormalizationWithoutSwallowingTheBody` (no network: the page has no images).

- [ ] **Step 6: Commit**

```bash
git add src/SendToOneNote.Core/Pages/AttachmentMarkup.cs src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs tests/SendToOneNote.Tests/AttachmentMarkupTests.cs tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs
git commit -m "feat: attachment object markup and notes after the header block (#16)"
git push
```

---

### Task 4: PageContent record on the backend seam (mechanical)

**Files:**
- Create: `src/SendToOneNote.Core/Pages/PageContent.cs`
- Modify: `src/SendToOneNote.Core/Backends/IOneNoteBackend.cs`, `GraphBackend.cs`, `DesktopOneNoteBackend.cs`
- Modify: `src/SendToOneNote.Core/Pages/PagePlanner.cs` (signature only)
- Modify: `src/SendToOneNote/SavePipeline.cs` (construct the record; attachments still `[]`)
- Modify: `tests/SendToOneNote.Tests/PagePlannerTests.cs`, `GraphBackendTests.cs`, `DesktopOneNoteBackendTests.cs`, `DesktopIntegrationSmokeTests.cs`, `IntegrationSmokeTests.cs`

**Interfaces:**
- Consumes: `PlannedAttachment` (Task 2).
- Produces:
  ```csharp
  public sealed record PageContent(string Xhtml, IReadOnlyList<ResolvedImage> Images, IReadOnlyList<PlannedAttachment> Attachments);
  Task<CreatedPage> IOneNoteBackend.CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default);
  public static PagePlan PagePlanner.Plan(PageContent content);
  ```
  No behaviour change in this task: both backends and the planner ignore `Attachments`.

- [ ] **Step 1: Add the record**

Create `src/SendToOneNote.Core/Pages/PageContent.cs`:

```csharp
namespace SendToOneNote.Core.Pages;

/// <summary>Everything a backend needs to create one page: resolved XHTML (src="name:imgN",
/// data="name:attN"), the image parts, and the attachments to embed.</summary>
public sealed record PageContent(string Xhtml, IReadOnlyList<ResolvedImage> Images,
    IReadOnlyList<PlannedAttachment> Attachments);
```

- [ ] **Step 2: Change the seam and its implementations**

`IOneNoteBackend.cs`:

```csharp
    /// <param name="content">Page XHTML from PageXhtmlBuilder + ImageResolver, plus image and attachment parts.</param>
    Task<CreatedPage> CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default);
```

`GraphBackend.cs`:

```csharp
    public Task<CreatedPage> CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default) =>
        client.CreatePageAsync(sectionId, PagePlanner.Plan(content), ct);
```

`DesktopOneNoteBackend.cs` — signature `CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default)`; inside, `OneNotePageXmlBuilder.ExtractTitle(content.Xhtml)` and `DataUriInliner.Inline(content.Xhtml, content.Images)`.

`PagePlanner.cs` — `public static PagePlan Plan(PageContent content)`; first lines `var xhtml = content.Xhtml; var images = content.Images;` and the rest unchanged.

- [ ] **Step 3: Update every caller**

`src/SendToOneNote/SavePipeline.cs`:

```csharp
            var content = new PageContent(resolution.Xhtml, resolution.Images, []);
            var page = await backend.CreatePageAsync(pick.SectionId, content);
```

and in the diagnostics block `PagePlanner.Plan(content).DroppedPartNames`.

Tests — every `PagePlanner.Plan(x, imgs)` becomes `PagePlanner.Plan(new PageContent(x, imgs, []))`; every `backend.CreatePageAsync(id, xhtml, images[, ct])` becomes `backend.CreatePageAsync(id, new PageContent(xhtml, images, [])[, ct])`. In `PagePlannerTests` add a helper to keep the diffs small:

```csharp
    private static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images) =>
        PagePlanner.Plan(new PageContent(xhtml, images, []));
```

and replace `PagePlanner.Plan(` with `Plan(` in that file's tests. Add `using SendToOneNote.Core.Pages;` where missing.

- [ ] **Step 4: Build and run tests**

Run: `dotnet build SendToOneNote.slnx --nologo -v q && dotnet test --nologo -v q`
Expected: build clean (no warnings introduced), all tests pass, same counts as before this task.

- [ ] **Step 5: Commit**

```bash
git add src tests
git commit -m "refactor: PageContent record on the IOneNoteBackend seam (#16)"
git push
```

---

### Task 5: Graph path — attachments as parts under the caps

**Files:**
- Modify: `src/SendToOneNote.Core/Pages/PagePlanner.cs`
- Modify: `tests/SendToOneNote.Tests/PagePlannerTests.cs`
- Modify: `tests/SendToOneNote.Tests/OneNoteClientTests.cs`
- Modify: `tests/SendToOneNote.Tests/GraphBackendTests.cs`

**Interfaces:**
- Consumes: `PageContent`, `AttachmentMarkup.ObjectRegex/OmittedNote/GraphBudgetReason`.
- Produces: `PagePlan.Parts` now contains image parts (document order) followed by attachment parts (email order); `DroppedPartNames` includes dropped `attN` names.

- [ ] **Step 1: Write the failing tests**

Append to `tests/SendToOneNote.Tests/PagePlannerTests.cs`:

```csharp
    private static PlannedAttachment Att(int n, int bytes, string name = "f.pdf") =>
        new($"att{n}", new SendToOneNote.Core.Email.EmailAttachment(name, "application/pdf", new byte[bytes]));

    private static string XhtmlWith(int images, params PlannedAttachment[] atts)
    {
        var objs = string.Concat(atts.Select(AttachmentMarkup.ObjectElement));
        var imgs = string.Join("", Enumerable.Range(0, images).Select(i => $"<img src=\"name:img{i}\"/>"));
        return $"<html><head><title>t</title></head><body>{objs}{imgs}</body></html>";
    }

    [Fact]
    public void AttachmentThatFitsBecomesAPartAfterImages()
    {
        var a = Att(1, 1000);
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(2, a), Images(2), [a]));
        Assert.Equal(["img0", "img1", "att1"], plan.Parts.Select(p => p.Name));
        Assert.Equal("application/pdf", plan.Parts[2].ContentType);
        Assert.Contains("data=\"name:att1\"", plan.PresentationXhtml);
        Assert.Empty(plan.DroppedPartNames);
    }

    [Fact]
    public void AttachmentOverRemainingBudgetIsDroppedWithNote()
    {
        // 1.75 MB image + 2 MB attachment > 3.5 MB: the image keeps priority, the file is noted.
        var a = Att(1, 2_000_000, "big.pdf");
        var images = new List<ResolvedImage> { new("img0", "image/png", new byte[1_750_000]) };
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(1, a), images, [a]));
        Assert.Equal(["img0"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["att1"], plan.DroppedPartNames);
        Assert.DoesNotContain("name:att1", plan.PresentationXhtml);
        Assert.Contains("[attachment omitted: big.pdf, 1.9 MB, too large for OneNote online]", plan.PresentationXhtml);
    }

    [Fact]
    public void AttachmentBeyondPartCapIsDropped()
    {
        var a = Att(1, 10);
        var images = Enumerable.Range(0, 30).Select(i => Img(i, 100, 100)).ToList();
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(30, a), images, [a]));
        Assert.Equal(30, plan.Parts.Count);
        Assert.DoesNotContain(plan.Parts, p => p.Name == "att1");
        Assert.Equal(["att1"], plan.DroppedPartNames);
    }

    [Fact]
    public void SmallerAttachmentStillFitsAfterLargerOneIsDropped()
    {
        var big = Att(1, 3_600_000, "big.pdf"); // alone exceeds the 3.5 MB budget
        var small = Att(2, 100, "small.pdf");
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(0, big, small), [], [big, small]));
        Assert.Equal(["att2"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["att1"], plan.DroppedPartNames);
    }
```

Append to `tests/SendToOneNote.Tests/OneNoteClientTests.cs` (inside the class, reuse `CreatedJson` and `FakeTokens` already defined there):

```csharp
    [Fact]
    public async Task AttachmentPartIsSentWithItsOwnContentType()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Created)
        { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") });
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>",
            [new OneNoteRequestPart("att1", "application/pdf", [0x25, 0x50, 0x44, 0x46])], []);
        await new OneNoteClient(new FakeTokens(), stub).CreatePageAsync("s1", plan);
        var body = await Assert.Single(stub.Requests).Content!.ReadAsStringAsync();
        Assert.Contains("name=att1", body.Replace("\"", ""));
        Assert.Contains("Content-Type: application/pdf", body);
    }
```

Extend `GraphBackendTests.CreatePagePlansAndPostsThroughOneNoteClient`: pass an attachment through and assert it reaches the request —

```csharp
        var att = new PlannedAttachment("att1", new EmailAttachment("a.pdf", "application/pdf", [1, 2, 3]));
        var xhtml = "<html><head><title>t</title></head><body>" + AttachmentMarkup.ObjectElement(att) +
                    "<img src=\"name:img0\"/></body></html>";
        var page = await backend.CreatePageAsync("s1", new PageContent(xhtml, [new("img0", "image/png", StubPng())], [att]));
        …
        var body = await req.Content!.ReadAsStringAsync();
        Assert.Contains("name=att1", body.Replace("\"", ""));
```

(add `using SendToOneNote.Core.Email; using SendToOneNote.Core.Pages;`).

- [ ] **Step 2: Run to verify failure**

Run: `dotnet test --nologo -v q --filter "FullyQualifiedName~PagePlannerTests|FullyQualifiedName~GraphBackendTests"`
Expected: the four new planner tests and the Graph backend test fail (no `att1` part); the client test passes already (the client is generic) — that is fine, it pins the contract.

- [ ] **Step 3: Implement**

In `PagePlanner.Plan`, after the image selection loop and the drop loop, before `var parts = …`:

```csharp
        // Attachments come after images (readable content wins) in email order, under the same caps.
        // One that doesn't fit is noted on the page; a later smaller one may still fit.
        var attachmentParts = new List<OneNoteRequestPart>();
        foreach (var a in content.Attachments)
        {
            var bytes = a.Attachment.Data.Length;
            if (selected.Count + attachmentParts.Count >= MaxBinaryPartsPerRequest || bytes > budget)
            {
                xhtml = AttachmentMarkup.ObjectRegex(a.PartName).Replace(xhtml,
                    AttachmentMarkup.OmittedNote(a.Attachment.FileName, bytes, AttachmentMarkup.GraphBudgetReason).Replace("$", "$$"));
                dropped.Add(a.PartName);
                continue;
            }
            budget -= bytes;
            attachmentParts.Add(new OneNoteRequestPart(a.PartName, a.Attachment.ContentType, a.Attachment.Data));
        }
```

and build `parts` as the image parts followed by `attachmentParts`:

```csharp
        var parts = kept.Where(i => selected.Contains(i.PartName))   // document order for stability
            .Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data))
            .Concat(attachmentParts)
            .ToList();
```

- [ ] **Step 4: Run all tests**

Run: `dotnet test --nologo -v q`
Expected: all pass. (`1.9 MB` in the budget test is 2,000,000 / 1,048,576 = 1.907 → "1.9 MB".)

- [ ] **Step 5: Commit**

```bash
git add src/SendToOneNote.Core/Pages/PagePlanner.cs tests/SendToOneNote.Tests/PagePlannerTests.cs tests/SendToOneNote.Tests/OneNoteClientTests.cs tests/SendToOneNote.Tests/GraphBackendTests.cs
git commit -m "feat: Graph path embeds attachments as parts under the request caps (#16)"
git push
```

---

### Task 6: Desktop path — temp files and one:InsertedFile

**Files:**
- Create: `src/SendToOneNote.Core/Desktop/AttachmentTempFolder.cs`
- Modify: `src/SendToOneNote.Core/Desktop/OneNotePageXmlBuilder.cs`
- Modify: `src/SendToOneNote.Core/Backends/DesktopOneNoteBackend.cs`
- Create: `tests/SendToOneNote.Tests/AttachmentTempFolderTests.cs`
- Modify: `tests/SendToOneNote.Tests/OneNotePageXmlBuilderTests.cs`
- Modify: `tests/SendToOneNote.Tests/FakeOneNoteApplication.cs`
- Modify: `tests/SendToOneNote.Tests/DesktopOneNoteBackendTests.cs`

**Interfaces:**
- Consumes: `PageContent`, `PlannedAttachment`, `AttachmentMarkup.StripAll`.
- Produces:
  ```csharp
  public sealed record InsertedFile(string PathSource, string PreferredName);
  public sealed class AttachmentTempFolder : IDisposable
  {
      public static string DefaultRoot { get; }                                  // %TEMP%\SendToOneNote
      public string? Path { get; }                                               // null when no attachments
      public IReadOnlyList<InsertedFile> Files { get; }
      public static AttachmentTempFolder Create(IReadOnlyList<PlannedAttachment> attachments, string? root = null);
      public static string SanitiseFileName(string name);
  }
  public static string OneNotePageXmlBuilder.Build(string pageId, string title, string html, IReadOnlyList<InsertedFile> files);
  ```

- [ ] **Step 1: Write the failing tests**

Create `tests/SendToOneNote.Tests/AttachmentTempFolderTests.cs`:

```csharp
using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class AttachmentTempFolderTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "stn-tests", Guid.NewGuid().ToString("N"));

    private static PlannedAttachment Att(int n, string name, byte[]? data = null) =>
        new($"att{n}", new EmailAttachment(name, "application/octet-stream", data ?? [1, 2, 3]));

    [Fact]
    public void WritesEachFileAndDeletesTheFolderOnDispose()
    {
        string? folder;
        using (var temp = AttachmentTempFolder.Create([Att(1, "a.pdf", [9, 8]), Att(2, "b.txt")], _root))
        {
            folder = temp.Path;
            Assert.NotNull(folder);
            Assert.StartsWith(_root, folder);
            Assert.Equal(2, temp.Files.Count);
            Assert.Equal("a.pdf", temp.Files[0].PreferredName);
            Assert.Equal(new byte[] { 9, 8 }, File.ReadAllBytes(temp.Files[0].PathSource));
            Assert.Equal(Path.Combine(folder, "b.txt"), temp.Files[1].PathSource);
        }
        Assert.False(Directory.Exists(folder));
    }

    [Fact]
    public void NoAttachmentsCreatesNothing()
    {
        using var temp = AttachmentTempFolder.Create([], _root);
        Assert.Null(temp.Path);
        Assert.Empty(temp.Files);
        Assert.False(Directory.Exists(_root));
    }

    [Theory]
    [InlineData("in:va|id?.pdf", "in_va_id_.pdf")]
    [InlineData("   ", "attachment")]
    [InlineData("..", "attachment")]
    [InlineData("plain.docx", "plain.docx")]
    public void SanitisesNames(string input, string expected) =>
        Assert.Equal(expected, AttachmentTempFolder.SanitiseFileName(input));

    [Fact]
    public void DuplicateNamesGetNumericSuffixButKeepPreferredName()
    {
        using var temp = AttachmentTempFolder.Create([Att(1, "same.pdf"), Att(2, "same.pdf")], _root);
        Assert.Equal("same.pdf", Path.GetFileName(temp.Files[0].PathSource));
        Assert.Equal("same (2).pdf", Path.GetFileName(temp.Files[1].PathSource));
        Assert.Equal("same.pdf", temp.Files[1].PreferredName);
    }

    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }
}
```

Append to `tests/SendToOneNote.Tests/OneNotePageXmlBuilderTests.cs`:

```csharp
    [Fact]
    public void InsertedFilesGoInTheirOwnOutlineBeforeTheHtmlBlock()
    {
        var xml = OneNotePageXmlBuilder.Build("{P}", "t", "<p>hi</p>",
            [new InsertedFile(@"C:\tmp\a & b.pdf", "a & b.pdf"), new InsertedFile(@"C:\tmp\c.docx", "c.docx")]);
        var doc = XDocument.Parse(xml);
        XNamespace one = OneNoteConstants.Namespace2013;
        var outlines = doc.Root!.Elements(one + "Outline").ToList();
        Assert.Equal(2, outlines.Count);
        var files = outlines[0].Element(one + "OEChildren")!.Elements(one + "OE")
            .Select(oe => oe.Element(one + "InsertedFile")!).ToList();
        Assert.Equal(2, files.Count);
        Assert.Equal(@"C:\tmp\a & b.pdf", files[0].Attribute("pathSource")!.Value);
        Assert.Equal("a & b.pdf", files[0].Attribute("preferredName")!.Value);
        Assert.NotNull(outlines[1].Element(one + "OEChildren")!.Element(one + "HTMLBlock"));
    }

    [Fact]
    public void NoFilesIsIdenticalToThreeArgumentBuild() =>
        Assert.Equal(OneNotePageXmlBuilder.Build("{P}", "t", "<p>hi</p>"),
            OneNotePageXmlBuilder.Build("{P}", "t", "<p>hi</p>", []));
```

In `tests/SendToOneNote.Tests/FakeOneNoteApplication.cs` make the fake observe the temp files while the call is in flight — add a property and extend `UpdatePageContent`:

```csharp
    /// <summary>pathSource → bytes, read DURING UpdatePageContent (the files must exist then, not after).</summary>
    public Dictionary<string, byte[]> FilesSeenDuringUpdate { get; } = [];

    public void UpdatePageContent(string bstrPageChangesXmlIn, DateTime dateExpectedLastModified, int xsSchema, bool force)
    {
        Touch();
        if (ThrowOnUpdate is not null) throw ThrowOnUpdate;
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(bstrPageChangesXmlIn, "pathSource=\"([^\"]+)\""))
        {
            var path = System.Net.WebUtility.HtmlDecode(m.Groups[1].Value);
            FilesSeenDuringUpdate[path] = File.Exists(path) ? File.ReadAllBytes(path) : [];
        }
        UpdatedPageXml.Add(bstrPageChangesXmlIn);
    }
```

Append to `tests/SendToOneNote.Tests/DesktopOneNoteBackendTests.cs`:

```csharp
    private static PlannedAttachment Att(string name, byte[] data) =>
        new("att1", new SendToOneNote.Core.Email.EmailAttachment(name, "application/pdf", data));

    private static string XhtmlWithAttachment(PlannedAttachment a) =>
        "<html><head><title>t</title></head><body>" + AttachmentMarkup.ObjectElement(a) + "<p>hi</p></body></html>";

    [Fact]
    public async Task AttachmentIsWrittenForTheComCallThenCleanedUpAndObjectMarkupStripped()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var a = Att("invoice.pdf", [0x25, 0x50, 0x44, 0x46]);

        await backend.CreatePageAsync("{S1}", new PageContent(XhtmlWithAttachment(a), [], [a]));

        var (path, bytes) = Assert.Single(fake.FilesSeenDuringUpdate);
        Assert.EndsWith("invoice.pdf", path);
        Assert.Equal(a.Attachment.Data, bytes);           // existed, with the right bytes, during the call
        Assert.False(File.Exists(path));                   // gone afterwards
        Assert.False(Directory.Exists(Path.GetDirectoryName(path)));
        var xml = Assert.Single(fake.UpdatedPageXml);
        Assert.Contains("one:InsertedFile", xml);
        Assert.Contains("preferredName=\"invoice.pdf\"", xml);
        Assert.DoesNotContain("<object", xml);
    }

    [Fact]
    public async Task TempFolderIsRemovedWhenTheComCallThrows()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("boom", unchecked((int)0x8004200B)) };
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var a = Att("x.pdf", [1]);
        var before = Directory.Exists(AttachmentTempFolder.DefaultRoot)
            ? Directory.GetDirectories(AttachmentTempFolder.DefaultRoot).ToHashSet() : [];

        await Assert.ThrowsAsync<DesktopOneNoteException>(() =>
            backend.CreatePageAsync("{S1}", new PageContent(XhtmlWithAttachment(a), [], [a])));

        var after = Directory.Exists(AttachmentTempFolder.DefaultRoot)
            ? Directory.GetDirectories(AttachmentTempFolder.DefaultRoot).ToHashSet() : [];
        Assert.Subset(before, after); // nothing new left behind
    }

    [Fact]
    public async Task NoAttachmentsProducesSingleOutlineXml()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        await backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], []));
        Assert.DoesNotContain("one:InsertedFile", Assert.Single(fake.UpdatedPageXml));
        Assert.Empty(fake.FilesSeenDuringUpdate);
    }
```

- [ ] **Step 2: Run to verify failure**

Run: `dotnet build SendToOneNote.slnx --nologo -v q`
Expected: compile errors for `AttachmentTempFolder`, `InsertedFile`, four-argument `Build`.

- [ ] **Step 3: Implement AttachmentTempFolder**

Create `src/SendToOneNote.Core/Desktop/AttachmentTempFolder.cs`:

```csharp
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Desktop;

/// <summary>A file OneNote should embed: pathSource is read at import; preferredName is what the page shows.</summary>
public sealed record InsertedFile(string PathSource, string PreferredName);

/// <summary>
/// one:InsertedFile needs the bytes on disk. This writes each attachment to a per-save folder under
/// %TEMP%\SendToOneNote and deletes the folder on Dispose — OneNote copies the bytes into the notebook
/// during UpdatePageContent, so the files are not needed afterwards.
/// </summary>
public sealed class AttachmentTempFolder : IDisposable
{
    public static string DefaultRoot { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SendToOneNote");

    public string? Path { get; }
    public IReadOnlyList<InsertedFile> Files { get; }

    private AttachmentTempFolder(string? path, IReadOnlyList<InsertedFile> files) { Path = path; Files = files; }

    public static AttachmentTempFolder Create(IReadOnlyList<PlannedAttachment> attachments, string? root = null)
    {
        if (attachments.Count == 0) return new AttachmentTempFolder(null, []);

        var folder = System.IO.Path.Combine(root ?? DefaultRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var files = new List<InsertedFile>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var a in attachments)
            {
                var name = SanitiseFileName(a.Attachment.FileName);
                var onDisk = name;
                for (var n = 2; !used.Add(onDisk); n++)
                    onDisk = System.IO.Path.GetFileNameWithoutExtension(name) + $" ({n})" + System.IO.Path.GetExtension(name);
                var path = System.IO.Path.Combine(folder, onDisk);
                File.WriteAllBytes(path, a.Attachment.Data);
                files.Add(new InsertedFile(path, a.Attachment.FileName));
            }
        }
        catch
        {
            TryDelete(folder);
            throw;
        }
        return new AttachmentTempFolder(folder, files);
    }

    public static string SanitiseFileName(string name)
    {
        var invalid = System.IO.Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
        return cleaned.Length == 0 || cleaned.All(c => c == '.' || c == '_') ? "attachment" : cleaned;
    }

    public void Dispose() { if (Path is not null) TryDelete(Path); }

    // Best effort: a leftover temp folder must never turn a successful save into a failure.
    private static void TryDelete(string folder)
    {
        try { if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
```

Check the sanitiser theory: `"in:va|id?.pdf"` → `:`, `|`, `?` are invalid → `in_va_id_.pdf` ✓; `"   "` → empty → `attachment` ✓; `".."` → TrimEnd('.') → empty → `attachment` ✓.

- [ ] **Step 4: Implement the XML builder overload**

In `src/SendToOneNote.Core/Desktop/OneNotePageXmlBuilder.cs`:

```csharp
    public static string Build(string pageId, string title, string html) => Build(pageId, title, html, []);

    public static string Build(string pageId, string title, string html, IReadOnlyList<InsertedFile> files) =>
        $"""
        <?xml version="1.0"?>
        <one:Page xmlns:one="{OneNoteConstants.Namespace2013}" ID="{pageId}">
          <one:Title><one:OE><one:T><![CDATA[{CdataEscape(title)}]]></one:T></one:OE></one:Title>
        {FileOutline(files)}  <one:Outline><one:OEChildren>
            <one:HTMLBlock><one:Data><![CDATA[{CdataEscape(html)}]]></one:Data></one:HTMLBlock>
          </one:OEChildren></one:Outline>
        </one:Page>
        """;

    // Files sit in their own outline ahead of the HTML block (an HTMLBlock cannot contain an InsertedFile).
    private static string FileOutline(IReadOnlyList<InsertedFile> files)
    {
        if (files.Count == 0) return "";
        var oes = string.Concat(files.Select(f =>
            $"    <one:OE><one:InsertedFile pathSource=\"{XmlAttr(f.PathSource)}\" preferredName=\"{XmlAttr(f.PreferredName)}\"/></one:OE>\n"));
        return "  <one:Outline><one:OEChildren>\n" + oes + "  </one:OEChildren></one:Outline>\n";
    }

    private static string XmlAttr(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace("\"", "&quot;");
```

Verify `NoFilesIsIdenticalToThreeArgumentBuild` passes exactly — with no files the interpolation inserts an empty string, leaving the original layout. If the raw-string indentation fights you, build the whole document with a `StringBuilder` instead; the tests parse the XML, they do not compare whitespace, except that the three- and four-argument forms must match each other.

- [ ] **Step 5: Implement the backend change**

In `DesktopOneNoteBackend.CreatePageAsync`:

```csharp
            return Guard(() =>
            {
                var title = OneNotePageXmlBuilder.ExtractTitle(content.Xhtml);
                var html = DataUriInliner.Inline(AttachmentMarkup.StripAll(content.Xhtml), content.Images);
                // Bytes must be on disk for one:InsertedFile; the folder is deleted when this block exits,
                // success or failure — OneNote has copied them into the notebook by then.
                using var temp = AttachmentTempFolder.Create(content.Attachments);
                App.CreateNewPage(sectionId, out var pageId, OneNoteConstants.NpsDefault);
                App.UpdatePageContent(OneNotePageXmlBuilder.Build(pageId, title, html, temp.Files),
                    DateTime.MinValue, OneNoteConstants.Xs2013, false);
                App.GetHyperlinkToObject(pageId, "", out var link);
                return new CreatedPage(pageId, link, null);
            });
```

`Guard` only catches `COMException`; an `IOException` from the temp write propagates as-is to the pipeline's generic catch (Failed folder + "Unexpected error — see log"), which is the spec's intent.

- [ ] **Step 6: Run all tests**

Run: `dotnet test --nologo -v q`
Expected: all pass.

- [ ] **Step 7: Commit**

```bash
git add src/SendToOneNote.Core/Desktop src/SendToOneNote.Core/Backends/DesktopOneNoteBackend.cs tests/SendToOneNote.Tests/AttachmentTempFolderTests.cs tests/SendToOneNote.Tests/OneNotePageXmlBuilderTests.cs tests/SendToOneNote.Tests/FakeOneNoteApplication.cs tests/SendToOneNote.Tests/DesktopOneNoteBackendTests.cs
git commit -m "feat: desktop path embeds attachments via one:InsertedFile from a per-save temp folder (#16)"
git push
```

---

### Task 7: Pipeline wiring and gated smoke tests

**Files:**
- Modify: `src/SendToOneNote/SavePipeline.cs`
- Modify: `tests/SendToOneNote.Tests/DesktopIntegrationSmokeTests.cs`
- Modify: `tests/SendToOneNote.Tests/IntegrationSmokeTests.cs`

**Interfaces:**
- Consumes: everything from Tasks 1–6.
- Produces: the feature end to end.

- [ ] **Step 1: Wire the pipeline**

In `src/SendToOneNote/SavePipeline.cs`, replace the three lines from `var xhtml = PageXhtmlBuilder.Build(email);` through `var page = await backend.CreatePageAsync(…)` with:

```csharp
            var attachmentPlan = AttachmentPlanner.Plan(email, settings.IncludeAttachments, settings.MaxAttachmentBytes);
            foreach (var o in attachmentPlan.Omitted)
                log.Info($"Attachment skipped ({o.Reason}): {o.FileName}, {o.Bytes} bytes");
            var xhtml = PageXhtmlBuilder.Build(email, attachmentPlan);
            var resolution = await _images.ResolveWithReportAsync(xhtml, email.InlineImages);
            var content = new PageContent(resolution.Xhtml, resolution.Images, attachmentPlan.Embedded);
            var page = await backend.CreatePageAsync(pick.SectionId, content);
            if (attachmentPlan.Embedded.Count > 0)
                log.Info($"Embedded {attachmentPlan.Embedded.Count} attachment(s) on the {backend.Name} path");
```

(`settings` is already loaded before the picker.) Update the diagnostics comment: "Only the Graph path can drop images or attachments (the 3.5 MB request cap)". Keep `PagePlanner.Plan(content).DroppedPartNames` from Task 4.

- [ ] **Step 2: Build and run the unit tests**

Run: `dotnet build SendToOneNote.slnx --nologo -v q && dotnet test --nologo -v q`
Expected: clean build, all pass.

- [ ] **Step 3: Add the desktop smoke test**

Append to `tests/SendToOneNote.Tests/DesktopIntegrationSmokeTests.cs` (add `using SendToOneNote.Core.Email;`):

```csharp
    [SkippableFact]
    public async Task EmbedsPdfAttachmentAsInsertedFile()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1", "Set STN_INTEGRATION=1 to run against desktop OneNote.");
        using var worker = new StaComWorker();
        Skip.If(!await worker.RunAsync(() => DesktopOneNoteProbe.IsAvailable()), "Desktop OneNote not installed.");
        var backend = new DesktopOneNoteBackend(worker);
        var tree = await backend.GetTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections).FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        // The real production pieces, on the synthetic PDF fixture.
        var email = EmlParser.Parse(Fixtures.Open("pdf-attachment.eml"));
        var plan = AttachmentPlanner.Plan(email, true, 26_214_400);
        var resolution = await new ImageResolver().ResolveWithReportAsync(PageXhtmlBuilder.Build(email, plan), email.InlineImages);
        var content = new PageContent(resolution.Xhtml, resolution.Images, plan.Embedded);

        var page = await backend.CreatePageAsync(scratch!.Id, content);
        Assert.NotEmpty(page.Id);

        var insertedFiles = await worker.RunAsync(() =>
        {
            var app = (IApplication)Activator.CreateInstance(Type.GetTypeFromProgID(OneNoteConstants.ProgId)!)!;
            app.GetPageContent(page.Id, out var xml, OneNoteConstants.PiBinaryData, OneNoteConstants.Xs2013);
            XNamespace one = OneNoteConstants.Namespace2013;
            return XDocument.Parse(xml).Descendants(one + "InsertedFile")
                .Select(f => f.Attribute("preferredName")?.Value).ToList();
        });
        Assert.Equal(["invoice.pdf"], insertedFiles);

        // Temp folder cleaned up (nothing of ours left under %TEMP%\SendToOneNote).
        if (Directory.Exists(AttachmentTempFolder.DefaultRoot))
            Assert.Empty(Directory.GetDirectories(AttachmentTempFolder.DefaultRoot));
    }
```

- [ ] **Step 4: Add the Graph smoke test**

Append to `tests/SendToOneNote.Tests/IntegrationSmokeTests.cs` (add `using SendToOneNote.Core.Backends; using SendToOneNote.Core.Email;`):

```csharp
    [SkippableFact]
    public async Task EmbedsPdfAttachmentViaGraph()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1",
            "Set STN_INTEGRATION=1 to run against the real Graph API.");
        var tokens = new MsalTokenProvider(Path.Combine(Path.GetTempPath(), "stn-int"));
        var backend = new GraphBackend(new OneNoteClient(tokens));
        var tree = await backend.GetTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections).FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        var email = EmlParser.Parse(Fixtures.Open("pdf-attachment.eml"));
        var plan = AttachmentPlanner.Plan(email, true, 26_214_400);
        var resolution = await new ImageResolver().ResolveWithReportAsync(PageXhtmlBuilder.Build(email, plan), email.InlineImages);
        var content = new PageContent(resolution.Xhtml, resolution.Images, plan.Embedded);
        Assert.Single(PagePlanner.Plan(content).Parts); // the PDF is the only part

        var page = await backend.CreatePageAsync(scratch!.Id, content);
        Assert.NotEmpty(page.Id);
    }
```

- [ ] **Step 5: Run the gated tests if this is the owner's machine, otherwise confirm they skip**

Run: `dotnet test --nologo -v q --filter "FullyQualifiedName~SmokeTests"`
Expected without `STN_INTEGRATION`: the new tests are reported as skipped. With `STN_INTEGRATION=1` on the owner's machine: both pass; if the desktop read-back finds no `InsertedFile`, or OneNote cannot open the file after the temp folder is gone, **stop and report** — that is the spec's flagged risk, and the fix (outline position, or keeping temp files) is a spec decision. Do not silently patch.

- [ ] **Step 6: Commit**

```bash
git add src/SendToOneNote/SavePipeline.cs tests/SendToOneNote.Tests/DesktopIntegrationSmokeTests.cs tests/SendToOneNote.Tests/IntegrationSmokeTests.cs
git commit -m "feat: save pipeline embeds attachments; gated smoke tests for both backends (#16)"
git push
```

---

### Task 8: Documentation

**Files:**
- Modify: `agents/knowledge/architecture.md`
- Modify: `agents/rules/privacy.md`
- Modify: `README.md`
- Modify: `docs/e2e-checklist.md`
- Modify: `AGENTS.md`

- [ ] **Step 1: architecture.md**

In "The save pipeline", after `EmlParser (headers, best body, inline CID images, attachment names)` change to `… inline CID images, attachments with bytes, attached-message names)`, and insert after the picker step: `→ AttachmentPlanner (IncludeAttachments switch + MaxAttachmentBytes cap → att1…attN) → PageXhtmlBuilder (title, header table, attachment objects/notes, body; …)`. Change the backend step to `IOneNoteBackend.CreatePageAsync(sectionId, PageContent{Xhtml, Images, Attachments})`.

In "Backends": the seam signature becomes `CreatePageAsync(sectionId, PageContent content)`; add to the `DesktopOneNoteBackend` bullet: "Attachments: object markup is stripped from the HTML, bytes are written to `%TEMP%\SendToOneNote\<guid>\` on the worker thread, the page XML gets a `one:InsertedFile` outline ahead of the HTML block, and the temp folder is deleted in a `finally`." Add to the `GraphBackend` bullet: "attachments become `attN` parts after the images under the same caps; ones that don't fit are replaced by a grey note."

In "Graph OneNote API constraints" add: `- Attachments: <object data-attachment="name" data="name:attN" type="…"></object> per file, counted against the 30-part / 3.5 MB caps after images.`

In "App data (runtime)" add `IncludeAttachments`, `MaxAttachmentBytes` to the settings list and a line: `%TEMP%\SendToOneNote\<guid>\ — per-save attachment temp files for the COM import, deleted after each save.`

In "Testing model" add the two new fixtures and the two new gated smoke tests to their bullets.

- [ ] **Step 2: privacy.md**

Under "Runtime data-handling invariants" add:

`- Attachment bytes go only where the email body goes: into the local desktop notebook (COM) or to Graph under the user's token. On the COM path they are written to a per-save folder under %TEMP%\SendToOneNote and deleted immediately after the import. Logs record attachment names and sizes, never content.`

- [ ] **Step 3: README.md**

In "How it saves" (after the Backend paragraph) add:

```
File attachments are embedded on the page as real file objects, right
under the header block, on both paths. Attached emails are listed by
name only. `"IncludeAttachments": false` in settings.json turns this off;
`"MaxAttachmentBytes"` (default 26214400 = 25 MB) skips larger files with
a note on the page. On the Graph path a file that doesn't fit the 3.5 MB
request budget is also skipped with a note.
```

- [ ] **Step 4: e2e-checklist.md**

Append:

```
- [ ] Email with a PDF attachment → file icon under the header on the desktop path; double-click opens the PDF after the temp folder is gone
- [ ] Same email with `"Backend": "graph"` → file icon on the page via Graph
- [ ] `"MaxAttachmentBytes": 1000` → grey "[attachment omitted: …, over the size limit]" note, page still created
- [ ] `"IncludeAttachments": false` → names listed in the header only, no icons, no notes
- [ ] Email with an attached message → "[attached message not embedded: …]" note
```

- [ ] **Step 5: AGENTS.md**

After the v1.1 line in "Planning & Scope Rules" add:

`- v2 attachments spec: `docs/superpowers/specs/2026-09-15-attachments-design.md`. Plan: `docs/superpowers/plans/2026-09-15-attachments.md`. GitHub issues #34–#41 map 1:1 to its Tasks 1–8; #42–#43 are follow-ups (image attachments as pictures, PDF printouts).`

Update the "v2 candidates" sentence so "attachments option" is no longer listed as out of scope (it is implemented); the remaining candidates stay.

- [ ] **Step 6: Build, test, commit**

Run: `dotnet test --nologo -v q`
Expected: all pass (docs only).

```bash
git add agents AGENTS.md README.md docs/e2e-checklist.md
git commit -m "docs: attachments feature in agent docs, README, and E2E checklist (#16)"
git push
```

---

## Self-review (done at plan time)

- **Spec coverage:** data model (T1), settings + planner (T2), markup with end tag and notes (T3), `PageContent` seam (T4), Graph budget rule + `DroppedPartNames` (T5), desktop temp files + InsertedFile outline + finally-cleanup (T6), pipeline order + logging + diagnostics (T7), fixtures/unit/local/gated tests (T1–T7), docs (T8). Backlog issues for picture-rendering and printouts are created outside the plan by the orchestrator.
- **Type consistency:** `PlannedAttachment(PartName, Attachment)` and `AttachmentPlan.Embedded` are used identically in T2–T7; `InsertedFile(PathSource, PreferredName)` in T6–T7; `AttachmentMarkup.ObjectElement/ObjectRegex/StripAll/OmittedNote/GraphBudgetReason` in T3, T5, T6.
- **Known judgment calls left to the implementer:** raw-string layout in `OneNotePageXmlBuilder` (tests parse XML, don't compare whitespace, except the three- vs four-arg equality); MimeKit's `Attachments` including `MessagePart` (fallback documented in T1).
