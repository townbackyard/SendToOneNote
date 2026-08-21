# SendToOneNote v1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Windows tray app: drag an email (.eml) from new Outlook into a watched folder, pick a OneNote section from a classic-style picker, and the email becomes a OneNote page (subject as title, From/To/Date block, HTML body with embedded images) via the Microsoft Graph OneNote API.

**Architecture:** Three projects — `SendToOneNote.Core` (all logic: parsing, page building, Graph client, auth, watcher, picker view-model; no UI dependencies beyond the Windows TFM), `SendToOneNote` (WPF tray app + windows), `SendToOneNote.Tests` (xUnit against Core using committed synthetic fixtures). The pipeline is: watcher → EmlParser → PageXhtmlBuilder → ImageResolver → PagePlanner → OneNoteClient.

**Tech Stack:** .NET 10 / C# / WPF (`net10.0-windows`), MimeKit (eml parsing), AngleSharp (HTML→XHTML + DOM rewriting), MSAL (Microsoft.Identity.Client + .Broker + .Extensions.Msal), H.NotifyIcon.Wpf (tray), System.Drawing.Common (image shrink), xUnit.

**Spec:** `docs/superpowers/specs/2026-08-21-sendtoonenote-design.md`

## Global Constraints

- .NET 10, TFM `net10.0-windows` for all three projects; Windows-only; WPF for UI.
- License: MIT. No telemetry of any kind. Email content is sent only to Microsoft Graph under the signed-in user's token.
- Delegated Graph scopes, exactly: `User.Read`, `Notes.ReadWrite`. Authority: `https://login.microsoftonline.com/common` (work/school AND personal accounts).
- Graph OneNote caps (from spec): ≤ 4 MB per request — planner uses a 3,500,000-byte safety cap; ≤ 5 binary parts per request in addition to the `Presentation`/`Commands` part.
- App data root: `%APPDATA%\SendToOneNote\` (settings.json, cache.json, `logs\`).
- Drop-folder behavior: `.eml` deleted on success (setting `DeleteOnSuccess`, default true); failures moved to `Failed\` subfolder; non-`.eml` files ignored and logged once.
- `fixtures/local/` is gitignored (real personal emails) — tests reference ONLY `fixtures/synthetic/`.
- v1 scope only. Do NOT implement v2 candidates (attachment embedding, auto-send mode, .txt/.png/.jpg drops, right-click handler).
- Commit style: `feat:`/`test:`/`chore:`/`docs:` prefixes, present tense.
- Windows-only APIs (System.Drawing, WPF, WAM broker) are acceptable everywhere; this app never targets other OSes.

---

### Task 1: Solution scaffold, CI, license

**Files:**
- Create: `SendToOneNote.sln`, `src/SendToOneNote.Core/SendToOneNote.Core.csproj`, `src/SendToOneNote/SendToOneNote.csproj`, `src/SendToOneNote/App.xaml`, `src/SendToOneNote/App.xaml.cs`, `tests/SendToOneNote.Tests/SendToOneNote.Tests.csproj`, `tests/SendToOneNote.Tests/SmokeTests.cs`, `.github/workflows/build.yml`, `LICENSE`, `Directory.Build.props`

**Interfaces:**
- Consumes: nothing.
- Produces: a building solution every later task adds code to; CI that runs `dotnet build` + `dotnet test` on push/PR.

- [ ] **Step 1: Create solution and projects**

```powershell
dotnet new sln -n SendToOneNote
dotnet new classlib -o src/SendToOneNote.Core -n SendToOneNote.Core
dotnet new wpf -o src/SendToOneNote -n SendToOneNote
dotnet new xunit -o tests/SendToOneNote.Tests -n SendToOneNote.Tests
dotnet sln add src/SendToOneNote.Core src/SendToOneNote tests/SendToOneNote.Tests
dotnet add tests/SendToOneNote.Tests reference src/SendToOneNote.Core
dotnet add src/SendToOneNote reference src/SendToOneNote.Core
```

- [ ] **Step 2: Pin TFMs and shared properties**

Create `Directory.Build.props` at repo root:

```xml
<Project>
  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <LangVersion>latest</LangVersion>
  </PropertyGroup>
</Project>
```

Remove any `<TargetFramework>` lines the templates put in the three `.csproj` files so the props file wins (keep `<UseWPF>true</UseWPF>` in `src/SendToOneNote/SendToOneNote.csproj`).

- [ ] **Step 3: Replace the template test with a smoke test**

`tests/SendToOneNote.Tests/SmokeTests.cs` (delete the template `UnitTest1.cs`):

```csharp
namespace SendToOneNote.Tests;

public class SmokeTests
{
    [Fact]
    public void SolutionBuildsAndTestsRun() => Assert.True(true);
}
```

- [ ] **Step 4: Build and test**

Run: `dotnet build && dotnet test`
Expected: build succeeds, 1 test passes.

- [ ] **Step 5: Add MIT license**

`LICENSE` — standard MIT text:

```text
MIT License

Copyright (c) 2026 TownBackyard

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

- [ ] **Step 6: Add CI workflow**

`.github/workflows/build.yml`:

```yaml
name: build
on:
  push:
    branches: [main]
  pull_request:
jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: 10.0.x
      - run: dotnet build --configuration Release
      - run: dotnet test --configuration Release --no-build
```

- [ ] **Step 7: Commit and verify CI**

```powershell
git add -A
git commit -m "chore: scaffold solution, CI, MIT license"
git push
```

Expected: GitHub Actions run on `main` goes green.

---

### Task 2: Synthetic test fixtures

**Files:**
- Create: `fixtures/synthetic/html-remote-images.eml`, `fixtures/synthetic/plain-text-receipt.eml`, `fixtures/synthetic/inline-cid-image.eml`, `fixtures/synthetic/malformed.eml`, `tests/SendToOneNote.Tests/Fixtures.cs`

**Interfaces:**
- Consumes: nothing.
- Produces: `static class Fixtures { static Stream Open(string name); static string Dir { get; } }` used by all parser/builder tests. Fixture file names above are load-bearing — later tasks reference them exactly.

- [ ] **Step 1: Write the fixture helper and a failing existence test**

`tests/SendToOneNote.Tests/Fixtures.cs`:

```csharp
namespace SendToOneNote.Tests;

public static class Fixtures
{
    public static string Dir { get; } = FindDir();

    public static Stream Open(string name) =>
        File.OpenRead(Path.Combine(Dir, name));

    private static string FindDir()
    {
        var d = AppContext.BaseDirectory;
        while (d is not null && !Directory.Exists(Path.Combine(d, "fixtures", "synthetic")))
            d = Path.GetDirectoryName(d);
        return Path.Combine(d ?? throw new DirectoryNotFoundException("fixtures/synthetic not found"),
            "fixtures", "synthetic");
    }
}

public class FixtureTests
{
    [Theory]
    [InlineData("html-remote-images.eml")]
    [InlineData("plain-text-receipt.eml")]
    [InlineData("inline-cid-image.eml")]
    [InlineData("malformed.eml")]
    public void FixtureExists(string name) =>
        Assert.True(File.Exists(Path.Combine(Fixtures.Dir, name)));
}
```

Run: `dotnet test --filter FixtureExists`
Expected: FAIL (files missing).

- [ ] **Step 2: Create `html-remote-images.eml`**

```text
From: Acme Newsletter <news@example.com>
To: Pat Sample <pat@example.net>
Subject: Weekly update: three things to know
Date: Thu, 20 Aug 2026 18:00:35 +0000
MIME-Version: 1.0
Content-Type: multipart/alternative; boundary="b1"

--b1
Content-Type: text/plain; charset="utf-8"

Three things to know this week.
--b1
Content-Type: text/html; charset="utf-8"

<html><body>
<h1>Weekly update</h1>
<p>Three things to know this week.</p>
<img src="https://img.example.com/banner.png" width="600">
<table><tr><td>Item one</td><td><img src="https://img.example.com/icon1.png"></td></tr>
<tr><td>Item two</td><td><img src="https://img.example.com/icon2.png"></td></tr></table>
<p>See you next week.<br>The Acme Team</p>
</body></html>
--b1--
```

Note the deliberately non-XHTML bits (`<img ...>` unclosed, `<br>`) — the pipeline must normalize them.

- [ ] **Step 3: Create `plain-text-receipt.eml`**

```text
From: "noreply@example.edu" <noreply@example.edu>
To: "pat@example.net" <pat@example.net>
Subject: Payment Processed
Date: Sat, 15 Aug 2026 13:39:07 +0000
MIME-Version: 1.0
Content-Type: text/plain; charset="iso-8859-1"
Content-Transfer-Encoding: quoted-printable

Receipt Number: 1234567
Customer: SAMPLE, PAT

Description                                                        Amount
------------------------------------------------------------------------
Installment Plan Payment                                       $1,000.00
                                                    Total      $1,000.00

Thank you for the payment.
```

- [ ] **Step 4: Create `inline-cid-image.eml`** (1×1 transparent PNG as inline CID part)

```text
From: Sender <sender@example.com>
To: pat@example.net
Subject: Photo inside
Date: Fri, 21 Aug 2026 09:00:00 +0000
MIME-Version: 1.0
Content-Type: multipart/related; boundary="b2"

--b2
Content-Type: text/html; charset="utf-8"

<html><body><p>Here is the photo:</p><img src="cid:photo1@example"><p>Regards</p></body></html>
--b2
Content-Type: image/png; name="photo.png"
Content-Transfer-Encoding: base64
Content-ID: <photo1@example>
Content-Disposition: inline; filename="photo.png"

iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAE
hQGAhKmMIQAAAABJRU5ErkJggg==
--b2--
```

- [ ] **Step 5: Create `malformed.eml`**

```text
this is not an email at all
just some bytes
```

- [ ] **Step 6: Verify and commit**

Run: `dotnet test --filter FixtureExists`
Expected: PASS (4/4).

```powershell
git add fixtures/synthetic tests
git commit -m "test: add synthetic .eml fixtures and fixture helper"
```

---

### Task 3: EmlParser (MimeKit)

**Files:**
- Create: `src/SendToOneNote.Core/Email/ParsedEmail.cs`, `src/SendToOneNote.Core/Email/EmlParser.cs`
- Test: `tests/SendToOneNote.Tests/EmlParserTests.cs`

**Interfaces:**
- Consumes: `Fixtures.Open(name)` from Task 2.
- Produces:

```csharp
namespace SendToOneNote.Core.Email;
public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);
public sealed record ParsedEmail(
    string Subject, string From, string To, string? Cc,
    DateTimeOffset? SentDate, string? HtmlBody, string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<string> AttachmentNames);
public static class EmlParser { public static ParsedEmail Parse(Stream emlStream); } // throws EmlParseException on garbage
public sealed class EmlParseException(string message, Exception? inner = null) : Exception(message, inner);
```

- [ ] **Step 1: Add MimeKit**

Run: `dotnet add src/SendToOneNote.Core package MimeKit`

- [ ] **Step 2: Write failing tests**

`tests/SendToOneNote.Tests/EmlParserTests.cs`:

```csharp
using SendToOneNote.Core.Email;

namespace SendToOneNote.Tests;

public class EmlParserTests
{
    [Fact]
    public void ParsesHtmlEmail()
    {
        var e = EmlParser.Parse(Fixtures.Open("html-remote-images.eml"));
        Assert.Equal("Weekly update: three things to know", e.Subject);
        Assert.Contains("news@example.com", e.From);
        Assert.Contains("pat@example.net", e.To);
        Assert.Null(e.Cc);
        Assert.Equal(new DateTimeOffset(2026, 8, 20, 18, 0, 35, TimeSpan.Zero), e.SentDate);
        Assert.NotNull(e.HtmlBody);
        Assert.Contains("banner.png", e.HtmlBody);
        Assert.Empty(e.InlineImages);
        Assert.Empty(e.AttachmentNames);
    }

    [Fact]
    public void ParsesPlainTextEmail()
    {
        var e = EmlParser.Parse(Fixtures.Open("plain-text-receipt.eml"));
        Assert.Null(e.HtmlBody);
        Assert.NotNull(e.TextBody);
        Assert.Contains("Receipt Number: 1234567", e.TextBody);
    }

    [Fact]
    public void ExtractsInlineCidImage()
    {
        var e = EmlParser.Parse(Fixtures.Open("inline-cid-image.eml"));
        var img = Assert.Single(e.InlineImages);
        Assert.Equal("photo1@example", img.ContentId);
        Assert.Equal("image/png", img.ContentType);
        Assert.True(img.Data.Length > 20); // decoded PNG bytes, not base64 text
        Assert.Equal(0x89, img.Data[0]);   // PNG magic
    }

    [Fact]
    public void MissingSubjectBecomesPlaceholder()
    {
        var raw = "From: a@b.c\r\nTo: d@e.f\r\n\r\nbody";
        var e = EmlParser.Parse(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(raw)));
        Assert.Equal("(no subject)", e.Subject);
    }

    [Fact]
    public void GarbageThrowsEmlParseException()
    {
        // MimeKit is lenient; "garbage" means empty content — no body AND no headers of interest
        var e = Record.Exception(() => EmlParser.Parse(Fixtures.Open("malformed.eml")));
        // Accept either behavior contract: exception, or a parse with no usable body treated upstream.
        // Contract chosen: throw when there is no HTML body, no text body, and no From header.
        Assert.IsType<EmlParseException>(e);
    }
}
```

Run: `dotnet test --filter EmlParserTests`
Expected: FAIL (types missing).

- [ ] **Step 3: Implement**

`src/SendToOneNote.Core/Email/ParsedEmail.cs`:

```csharp
namespace SendToOneNote.Core.Email;

public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);

public sealed record ParsedEmail(
    string Subject,
    string From,
    string To,
    string? Cc,
    DateTimeOffset? SentDate,
    string? HtmlBody,
    string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<string> AttachmentNames);

public sealed class EmlParseException(string message, Exception? inner = null)
    : Exception(message, inner);
```

`src/SendToOneNote.Core/Email/EmlParser.cs`:

```csharp
using MimeKit;

namespace SendToOneNote.Core.Email;

public static class EmlParser
{
    public static ParsedEmail Parse(Stream emlStream)
    {
        MimeMessage msg;
        try
        {
            msg = MimeMessage.Load(emlStream);
        }
        catch (Exception ex)
        {
            throw new EmlParseException("Not a readable email file.", ex);
        }

        var html = msg.HtmlBody;
        var text = msg.TextBody;
        var from = msg.From.ToString();

        if (html is null && text is null && string.IsNullOrWhiteSpace(from))
            throw new EmlParseException("File has no email content (no body, no sender).");

        var inline = new List<InlineImage>();
        foreach (var part in msg.BodyParts.OfType<MimePart>())
        {
            if (part.ContentType.MediaType != "image" || part.ContentId is null)
                continue;
            using var ms = new MemoryStream();
            part.Content.DecodeTo(ms);
            inline.Add(new InlineImage(
                part.ContentId.Trim('<', '>'),
                part.FileName ?? "image",
                part.ContentType.MimeType,
                ms.ToArray()));
        }

        var attachments = msg.Attachments.OfType<MimePart>()
            .Where(p => p.ContentId is null) // inline images already captured
            .Select(p => p.FileName ?? "attachment")
            .ToList();

        DateTimeOffset? sent = msg.Date == DateTimeOffset.MinValue ? null : msg.Date;

        return new ParsedEmail(
            Subject: string.IsNullOrWhiteSpace(msg.Subject) ? "(no subject)" : msg.Subject,
            From: from,
            To: msg.To.ToString(),
            Cc: msg.Cc.Count > 0 ? msg.Cc.ToString() : null,
            SentDate: sent,
            HtmlBody: html,
            TextBody: text,
            InlineImages: inline,
            AttachmentNames: attachments);
    }
}
```

- [ ] **Step 4: Run tests**

Run: `dotnet test --filter EmlParserTests`
Expected: PASS (5/5). If `GarbageThrowsEmlParseException` fails because MimeKit parsed the garbage into a From-less body-less message differently, adjust only the guard clause (the contract is: no HtmlBody + no TextBody + blank From ⇒ throw).

- [ ] **Step 5: Sanity-check against real local fixtures (manual, not committed)**

Add a temporary console check or run the tests below only if `fixtures/local/` exists (skip otherwise):

```csharp
// tests/SendToOneNote.Tests/LocalFixtureTests.cs
using SendToOneNote.Core.Email;

namespace SendToOneNote.Tests;

public class LocalFixtureTests
{
    public static IEnumerable<object[]> LocalEmls()
    {
        var dir = Path.Combine(Path.GetDirectoryName(Fixtures.Dir)!, "local");
        if (!Directory.Exists(dir)) yield break;
        foreach (var f in Directory.GetFiles(dir, "*.eml"))
            yield return new object[] { f };
    }

    [Theory]
    [MemberData(nameof(LocalEmls))]
    public void ParsesRealEmailsWithoutThrowing(string path)
    {
        using var s = File.OpenRead(path);
        var e = EmlParser.Parse(s);
        Assert.False(string.IsNullOrWhiteSpace(e.Subject));
        Assert.True(e.HtmlBody is not null || e.TextBody is not null);
    }
}
```

Run: `dotnet test --filter LocalFixtureTests`
Expected: PASS on the dev machine (3 real emails), skipped (no data) on CI.

- [ ] **Step 6: Commit**

```powershell
git add src tests
git commit -m "feat: EmlParser extracts headers, bodies, inline CID images"
```

---

### Task 4: PageXhtmlBuilder (header block + body composition)

**Files:**
- Create: `src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs`
- Test: `tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs`

**Interfaces:**
- Consumes: `ParsedEmail` (Task 3).
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public static class PageXhtmlBuilder
{
    // Returns the full page XHTML: <html><head><title>subject</title></head><body>headerTable + body</body></html>
    // For HTML emails, bodyHtml is inserted raw (normalized later by ImageResolver).
    // For text-only emails, the text is converted here (paragraphs, or <pre> when columnar).
    public static string Build(ParsedEmail email);
    public static bool LooksColumnar(string text); // exposed for tests
}
```

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/PageXhtmlBuilderTests.cs`:

```csharp
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class PageXhtmlBuilderTests
{
    private static ParsedEmail Email(string? html = null, string? text = null,
        string subject = "S & T <test>", IReadOnlyList<string>? attachments = null) =>
        new(subject, "a@b.c", "d@e.f", null,
            new DateTimeOffset(2026, 8, 20, 18, 0, 35, TimeSpan.Zero),
            html, text, [], attachments ?? []);

    [Fact]
    public void TitleIsEscapedSubject()
    {
        var x = PageXhtmlBuilder.Build(Email(html: "<p>hi</p>"));
        Assert.Contains("<title>S &amp; T &lt;test&gt;</title>", x);
    }

    [Fact]
    public void HeaderBlockContainsFromToDate()
    {
        var x = PageXhtmlBuilder.Build(Email(html: "<p>hi</p>"));
        Assert.Contains("a@b.c", x);
        Assert.Contains("d@e.f", x);
        Assert.Contains("2026", x);
    }

    [Fact]
    public void AttachmentNamesListedWhenPresent()
    {
        var x = PageXhtmlBuilder.Build(Email(html: "<p>hi</p>", attachments: ["report.pdf"]));
        Assert.Contains("report.pdf", x);
    }

    [Fact]
    public void ColumnarTextUsesPre()
    {
        var text = "Description                Amount\r\n" +
                   "Payment               $1,000.00\r\n" +
                   "Total                 $1,000.00";
        Assert.True(PageXhtmlBuilder.LooksColumnar(text));
        var x = PageXhtmlBuilder.Build(Email(text: text));
        Assert.Contains("<pre", x);
    }

    [Fact]
    public void ProseTextUsesParagraphsWithEscaping()
    {
        var text = "Hello there.\r\n\r\nSecond paragraph with <angle> & ampersand.";
        Assert.False(PageXhtmlBuilder.LooksColumnar(text));
        var x = PageXhtmlBuilder.Build(Email(text: text));
        Assert.Contains("<p>Hello there.</p>", x);
        Assert.Contains("&lt;angle&gt; &amp; ampersand", x);
        Assert.DoesNotContain("<pre", x);
    }
}
```

Run: `dotnet test --filter PageXhtmlBuilderTests`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Pages/PageXhtmlBuilder.cs`:

```csharp
using System.Net;
using System.Text;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public static class PageXhtmlBuilder
{
    public static string Build(ParsedEmail email)
    {
        var sb = new StringBuilder();
        sb.Append("<html><head><title>")
          .Append(WebUtility.HtmlEncode(email.Subject))
          .Append("</title></head><body>");
        AppendHeaderTable(sb, email);
        sb.Append("<div>");
        if (email.HtmlBody is not null)
            sb.Append(email.HtmlBody); // normalized to XHTML later by ImageResolver
        else
            AppendTextBody(sb, email.TextBody ?? "");
        sb.Append("</div></body></html>");
        return sb.ToString();
    }

    private static void AppendHeaderTable(StringBuilder sb, ParsedEmail e)
    {
        sb.Append("<table style=\"font-size:10pt;color:#5b5b5b\">");
        Row(sb, "From", e.From);
        Row(sb, "To", e.To);
        if (e.Cc is not null) Row(sb, "Cc", e.Cc);
        if (e.SentDate is { } d) Row(sb, "Sent", d.ToLocalTime().ToString("f"));
        if (e.AttachmentNames.Count > 0)
            Row(sb, "Attachments", string.Join("; ", e.AttachmentNames));
        sb.Append("</table><hr/>");

        static void Row(StringBuilder sb, string label, string value) =>
            sb.Append("<tr><td style=\"font-weight:bold\">").Append(label)
              .Append("</td><td>").Append(WebUtility.HtmlEncode(value))
              .Append("</td></tr>");
    }

    private static void AppendTextBody(StringBuilder sb, string text)
    {
        if (LooksColumnar(text))
        {
            sb.Append("<pre style=\"font-family:Consolas;font-size:10pt\">")
              .Append(WebUtility.HtmlEncode(text)).Append("</pre>");
            return;
        }
        var blocks = text.Replace("\r\n", "\n")
            .Split("\n\n", StringSplitOptions.RemoveEmptyEntries);
        foreach (var block in blocks)
        {
            var lines = block.Split('\n').Select(WebUtility.HtmlEncode);
            sb.Append("<p>").Append(string.Join("<br/>", lines)).Append("</p>");
        }
    }

    public static bool LooksColumnar(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n')
            .Where(l => l.Trim().Length > 0).ToList();
        if (lines.Count == 0) return false;
        var columnar = lines.Count(l => l.Contains("   ")); // 3+ consecutive spaces
        return columnar >= Math.Max(2, lines.Count / 5);
    }
}
```

- [ ] **Step 3: Run tests**

Run: `dotnet test --filter PageXhtmlBuilderTests`
Expected: PASS (5/5).

- [ ] **Step 4: Commit**

```powershell
git add src tests
git commit -m "feat: PageXhtmlBuilder composes title, header block, and body"
```

---

### Task 5: ImageResolver (CID + remote download, XHTML normalization)

**Files:**
- Create: `src/SendToOneNote.Core/Pages/ImageResolver.cs`, `tests/SendToOneNote.Tests/StubHttpHandler.cs`
- Test: `tests/SendToOneNote.Tests/ImageResolverTests.cs`

**Interfaces:**
- Consumes: `InlineImage` (Task 3); page XHTML string from `PageXhtmlBuilder.Build` (Task 4).
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data);
public sealed class ImageResolver(HttpMessageHandler? handler = null)
{
    // Parses pageXhtml with AngleSharp (also fixing non-XHTML email markup),
    // resolves <img> tags: cid: → matching InlineImage part; http(s): → downloaded part.
    // Successful resolutions rewrite src to "name:img0", "name:img1", ...
    // Failures (missing cid, download error/timeout) leave the original src untouched.
    // Returns well-formed XHTML and the resolved parts in src-name order.
    public Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default);
}
```

- Per-image download timeout: 10 seconds. Parallelism: max 4 concurrent downloads.

- [ ] **Step 1: Add AngleSharp**

Run: `dotnet add src/SendToOneNote.Core package AngleSharp`

- [ ] **Step 2: Write the stub HTTP handler for tests**

`tests/SendToOneNote.Tests/StubHttpHandler.cs`:

```csharp
using System.Net;

namespace SendToOneNote.Tests;

public sealed class StubHttpHandler : HttpMessageHandler
{
    public List<HttpRequestMessage> Requests { get; } = [];
    private readonly Func<HttpRequestMessage, HttpResponseMessage> _responder;

    public StubHttpHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
        => _responder = responder;

    public static HttpResponseMessage Png(byte[] bytes) =>
        new(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(bytes)
            {
                Headers = { ContentType = new("image/png") }
            }
        };

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken ct)
    {
        Requests.Add(request);
        return Task.FromResult(_responder(request));
    }
}
```

- [ ] **Step 3: Write failing tests**

`tests/SendToOneNote.Tests/ImageResolverTests.cs`:

```csharp
using System.Net;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class ImageResolverTests
{
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public async Task RemoteImageDownloadedAndRewritten()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/a.png\"></body></html>", []);
        var img = Assert.Single(images);
        Assert.Equal("img0", img.PartName);
        Assert.Equal("image/png", img.ContentType);
        Assert.Contains("src=\"name:img0\"", xhtml);
    }

    [Fact]
    public async Task CidImageResolvedFromInlineParts()
    {
        var r = new ImageResolver(new StubHttpHandler(_ => throw new InvalidOperationException("no http expected")));
        var inline = new[] { new InlineImage("photo1@example", "p.png", "image/png", PngBytes) };
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"cid:photo1@example\"></body></html>", inline);
        Assert.Single(images);
        Assert.Contains("name:img0", xhtml);
        Assert.DoesNotContain("cid:", xhtml);
    }

    [Fact]
    public async Task FailedDownloadKeepsOriginalUrl()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/gone.png\"></body></html>", []);
        Assert.Empty(images);
        Assert.Contains("https://x.example/gone.png", xhtml);
    }

    [Fact]
    public async Task OutputIsWellFormedXhtml()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = new ImageResolver(stub);
        var (xhtml, _) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><p>a<br>b</p><img src=\"https://x.example/a.png\"></body></html>", []);
        // Must load as XML (self-closed br/img)
        var doc = System.Xml.Linq.XDocument.Parse(xhtml);
        Assert.NotNull(doc.Root);
    }
}
```

Run: `dotnet test --filter ImageResolverTests`
Expected: FAIL.

- [ ] **Step 4: Implement**

`src/SendToOneNote.Core/Pages/ImageResolver.cs`:

```csharp
using AngleSharp;
using AngleSharp.Html.Parser;
using AngleSharp.Xhtml;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data);

public sealed class ImageResolver
{
    private readonly HttpClient _http;

    public ImageResolver(HttpMessageHandler? handler = null)
    {
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var parser = new HtmlParser();
        var doc = await parser.ParseDocumentAsync(pageXhtml, ct);
        var imgs = doc.QuerySelectorAll("img").ToList();

        var resolved = new List<ResolvedImage>();
        var gate = new SemaphoreSlim(4);
        var work = imgs.Select(async img =>
        {
            var src = img.GetAttribute("src") ?? "";
            byte[]? data = null;
            string? contentType = null;

            if (src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
            {
                var cid = src[4..].Trim('<', '>');
                var match = inlineImages.FirstOrDefault(i =>
                    string.Equals(i.ContentId, cid, StringComparison.OrdinalIgnoreCase));
                if (match is not null) { data = match.Data; contentType = match.ContentType; }
            }
            else if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                     src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                await gate.WaitAsync(ct);
                try
                {
                    var resp = await _http.GetAsync(src, ct);
                    if (resp.IsSuccessStatusCode)
                    {
                        data = await resp.Content.ReadAsByteArrayAsync(ct);
                        contentType = resp.Content.Headers.ContentType?.MediaType ?? "image/png";
                    }
                }
                catch (Exception e) when (e is HttpRequestException or TaskCanceledException or UriFormatException)
                {
                    // leave original src
                }
                finally { gate.Release(); }
            }

            return (img, data, contentType);
        }).ToList();

        var results = await Task.WhenAll(work);
        foreach (var (img, data, contentType) in results)
        {
            if (data is null || contentType is null) continue;
            var name = $"img{resolved.Count}";
            resolved.Add(new ResolvedImage(name, contentType, data));
            img.SetAttribute("src", $"name:{name}");
        }

        var xhtml = doc.ToHtml(XhtmlMarkupFormatter.Instance);
        return (xhtml, resolved);
    }
}
```

- [ ] **Step 5: Run tests**

Run: `dotnet test --filter ImageResolverTests`
Expected: PASS (4/4). If `XhtmlMarkupFormatter` is in a different namespace in the current AngleSharp version, fix the `using` (search AngleSharp docs for "XhtmlMarkupFormatter") — do not hand-roll serialization.

- [ ] **Step 6: Commit**

```powershell
git add src tests
git commit -m "feat: ImageResolver embeds cid/remote images and normalizes XHTML"
```

---

### Task 6: PagePlanner (Graph request batching + image shrinking)

**Files:**
- Create: `src/SendToOneNote.Core/Pages/PagePlanner.cs`, `src/SendToOneNote.Core/Pages/ImageShrinker.cs`
- Test: `tests/SendToOneNote.Tests/PagePlannerTests.cs`

**Interfaces:**
- Consumes: XHTML + `ResolvedImage` list (Task 5).
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record AppendPlan(string CommandsJson, IReadOnlyList<OneNoteRequestPart> Parts);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts,
    IReadOnlyList<AppendPlan> Appends);
public static class PagePlanner
{
    public const int MaxRequestBytes = 3_500_000;
    public const int MaxBinaryPartsPerRequest = 5;
    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images);
}
public static class ImageShrinker
{
    // Returns input unchanged if already under maxBytes; otherwise re-encodes as JPEG,
    // halving dimensions until under maxBytes (floor 200px width).
    public static (byte[] Data, string ContentType) ShrinkIfNeeded(byte[] data, string contentType, int maxBytes);
}
```

- Batching contract: the first ≤5 images (subject to total size cap) ship with the create request, `src="name:imgN"` intact. Every overflow image's `<img src="name:imgN"/>` in the create XHTML is REPLACED by `<div data-id="slot-imgN"></div>`; each `AppendPlan` carries a Graph PATCH `Commands` JSON array of `{"target":"#slot-imgN","action":"replace","content":"<img src=\"name:imgN\"/>"}` plus ≤5 binary parts, each append batch under the size cap.

- [ ] **Step 1: Add System.Drawing.Common**

Run: `dotnet add src/SendToOneNote.Core package System.Drawing.Common`

- [ ] **Step 2: Write failing tests**

`tests/SendToOneNote.Tests/PagePlannerTests.cs`:

```csharp
using System.Text.Json;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class PagePlannerTests
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private static string XhtmlWith(int n)
    {
        var imgs = string.Join("", Enumerable.Range(0, n).Select(i => $"<img src=\"name:img{i}\"/>"));
        return $"<html><head><title>t</title></head><body>{imgs}</body></html>";
    }

    private static IReadOnlyList<ResolvedImage> Images(int n) =>
        Enumerable.Range(0, n).Select(i => new ResolvedImage($"img{i}", "image/png", Png)).ToList();

    [Fact]
    public void FewImagesSingleRequest()
    {
        var plan = PagePlanner.Plan(XhtmlWith(3), Images(3));
        Assert.Equal(3, plan.Parts.Count);
        Assert.Empty(plan.Appends);
        Assert.Contains("name:img2", plan.PresentationXhtml);
    }

    [Fact]
    public void OverflowImagesBecomeSlotsAndAppends()
    {
        var plan = PagePlanner.Plan(XhtmlWith(8), Images(8));
        Assert.Equal(5, plan.Parts.Count);
        Assert.DoesNotContain("name:img5", plan.PresentationXhtml);
        Assert.Contains("data-id=\"slot-img5\"", plan.PresentationXhtml);
        var append = Assert.Single(plan.Appends);
        Assert.Equal(3, append.Parts.Count);
        var cmds = JsonDocument.Parse(append.CommandsJson).RootElement;
        Assert.Equal(3, cmds.GetArrayLength());
        Assert.Equal("#slot-img5", cmds[0].GetProperty("target").GetString());
        Assert.Equal("replace", cmds[0].GetProperty("action").GetString());
        Assert.Contains("name:img5", cmds[0].GetProperty("content").GetString());
    }

    [Fact]
    public void OversizedImageGetsShrunk()
    {
        // 6 MB of fake bytes is not a decodable image; use a real bitmap instead
        using var bmp = new System.Drawing.Bitmap(3000, 3000);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp); // BMP = huge
        var (data, ct) = ImageShrinker.ShrinkIfNeeded(ms.ToArray(), "image/bmp", 500_000);
        Assert.True(data.Length <= 500_000);
        Assert.Equal("image/jpeg", ct);
    }

    [Fact]
    public void SmallImageUntouched()
    {
        var (data, ct) = ImageShrinker.ShrinkIfNeeded(Png, "image/png", 500_000);
        Assert.Same(Png, data);
        Assert.Equal("image/png", ct);
    }

    [Fact]
    public void UndecodableOversizedImageIsDropped()
    {
        var big = new byte[4_000_000]; // not a decodable image; shrinker passes it through
        var plan = PagePlanner.Plan(XhtmlWith(1), [new ResolvedImage("img0", "image/png", big)]);
        Assert.Empty(plan.Parts);
        Assert.Empty(plan.Appends);
        Assert.DoesNotContain("name:img0", plan.PresentationXhtml);
        Assert.Contains("image omitted", plan.PresentationXhtml);
    }
}
```

Run: `dotnet test --filter "PagePlannerTests"`
Expected: FAIL.

- [ ] **Step 3: Implement ImageShrinker**

`src/SendToOneNote.Core/Pages/ImageShrinker.cs`:

```csharp
using System.Drawing;
using System.Drawing.Imaging;

namespace SendToOneNote.Core.Pages;

public static class ImageShrinker
{
    public static (byte[] Data, string ContentType) ShrinkIfNeeded(
        byte[] data, string contentType, int maxBytes)
    {
        if (data.Length <= maxBytes) return (data, contentType);
        try
        {
            using var src = new MemoryStream(data);
            using var img = Image.FromStream(src);
            var (w, h) = (img.Width, img.Height);
            var current = data;
            while (current.Length > maxBytes && w >= 200)
            {
                using var bmp = new Bitmap(img, w, h);
                using var outMs = new MemoryStream();
                var jpeg = ImageCodecInfo.GetImageEncoders()
                    .First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                using var p = new EncoderParameters(1);
                p.Param[0] = new EncoderParameter(Encoder.Quality, 80L);
                bmp.Save(outMs, jpeg, p);
                current = outMs.ToArray();
                (w, h) = (w / 2, h / 2);
            }
            return (current, "image/jpeg");
        }
        catch (Exception)
        {
            return (data, contentType); // undecodable: pass through, planner may drop it
        }
    }
}
```

- [ ] **Step 4: Implement PagePlanner**

`src/SendToOneNote.Core/Pages/PagePlanner.cs`:

```csharp
using System.Text;
using System.Text.Json;

namespace SendToOneNote.Core.Pages;

public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record AppendPlan(string CommandsJson, IReadOnlyList<OneNoteRequestPart> Parts);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts,
    IReadOnlyList<AppendPlan> Appends);

public static class PagePlanner
{
    public const int MaxRequestBytes = 3_500_000;
    public const int MaxBinaryPartsPerRequest = 5;

    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        // Shrink anything that alone would blow the cap.
        var shrunk = images.Select(i =>
        {
            var (data, ct) = ImageShrinker.ShrinkIfNeeded(i.Data, i.ContentType, MaxRequestBytes / 2);
            return new ResolvedImage(i.PartName, ct, data);
        }).ToList();

        // Drop anything STILL over the cap (undecodable blobs the shrinker passed through):
        // a part that can never fit would 413 the request it rides on.
        var kept = new List<ResolvedImage>();
        foreach (var img in shrunk)
        {
            if (img.Data.Length > MaxRequestBytes - 4096)
                xhtml = xhtml.Replace($"<img src=\"name:{img.PartName}\"/>",
                    "<p style=\"color:#999999\">[image omitted: too large]</p>");
            else kept.Add(img);
        }

        // Greedy first batch for the create request.
        var firstBatch = new List<ResolvedImage>();
        long budget = MaxRequestBytes - Encoding.UTF8.GetByteCount(xhtml) - 4096;
        foreach (var img in kept)
        {
            if (firstBatch.Count >= MaxBinaryPartsPerRequest || img.Data.Length > budget) break;
            firstBatch.Add(img);
            budget -= img.Data.Length;
        }

        var overflow = kept.Skip(firstBatch.Count).ToList();
        var presentation = xhtml;
        foreach (var img in overflow)
            presentation = presentation.Replace(
                $"<img src=\"name:{img.PartName}\"/>",
                $"<div data-id=\"slot-{img.PartName}\"></div>");

        var appends = new List<AppendPlan>();
        var batch = new List<ResolvedImage>();
        long batchBytes = 0;
        foreach (var img in overflow)
        {
            if (batch.Count >= MaxBinaryPartsPerRequest ||
                batchBytes + img.Data.Length > MaxRequestBytes - 4096)
            {
                if (batch.Count > 0) appends.Add(ToAppend(batch));
                batch = []; batchBytes = 0;
            }
            batch.Add(img);
            batchBytes += img.Data.Length;
        }
        if (batch.Count > 0) appends.Add(ToAppend(batch));

        return new PagePlan(presentation,
            firstBatch.Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList(),
            appends);
    }

    private static AppendPlan ToAppend(List<ResolvedImage> batch)
    {
        var commands = batch.Select(i => new
        {
            target = $"#slot-{i.PartName}",
            action = "replace",
            content = $"<img src=\"name:{i.PartName}\"/>"
        });
        return new AppendPlan(JsonSerializer.Serialize(commands),
            batch.Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList());
    }
}
```

- [ ] **Step 5: Run tests**

Run: `dotnet test --filter "PagePlannerTests"`
Expected: PASS (4/4). Note: the `Replace` in Step 4 depends on ImageResolver emitting exactly `<img src="name:imgN"/>` (AngleSharp XHTML self-closing form). If the serialized form differs (e.g., includes other attributes), switch the replace to a regex over `<img[^>]*src="name:{img.PartName}"[^>]*/>` — and add a planner test with an `alt` attribute to lock it in.

- [ ] **Step 6: Commit**

```powershell
git add src tests
git commit -m "feat: PagePlanner batches Graph requests under size/part caps"
```

---

### Task 7: Settings + section cache persistence

**Files:**
- Create: `src/SendToOneNote.Core/Storage/AppSettings.cs`, `src/SendToOneNote.Core/Storage/JsonFileStore.cs`, `src/SendToOneNote.Core/OneNote/NotebookModels.cs`
- Test: `tests/SendToOneNote.Tests/StorageTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces:

```csharp
namespace SendToOneNote.Core.OneNote;
public sealed record SectionNode(string Id, string Name);
public sealed record GroupNode(string Id, string Name,
    IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);
public sealed record NotebookNode(string Id, string Name,
    IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);
public sealed record NotebookTree(IReadOnlyList<NotebookNode> Notebooks, DateTimeOffset FetchedUtc);

namespace SendToOneNote.Core.Storage;
public sealed class AppSettings
{
    public string? DropFolder { get; set; }
    public string? ClientIdOverride { get; set; }
    public bool DeleteOnSuccess { get; set; } = true;
    public List<string> RecentSectionIds { get; set; } = [];
}
public sealed class JsonFileStore(string? rootDir = null) // default %APPDATA%\SendToOneNote
{
    public string RootDir { get; }
    public AppSettings LoadSettings();          // missing/corrupt file → defaults
    public void SaveSettings(AppSettings s);    // atomic: write tmp then File.Move overwrite
    public NotebookTree? LoadTreeCache();       // cache.json; missing/corrupt → null
    public void SaveTreeCache(NotebookTree t);
}
```

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/StorageTests.cs`:

```csharp
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Storage;

namespace SendToOneNote.Tests;

public class StorageTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-tests-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void SettingsRoundTrip()
    {
        var store = new JsonFileStore(_dir);
        var s = store.LoadSettings();          // defaults on first run
        Assert.True(s.DeleteOnSuccess);
        s.DropFolder = @"C:\Drop";
        s.RecentSectionIds.Add("sec1");
        store.SaveSettings(s);
        var s2 = new JsonFileStore(_dir).LoadSettings();
        Assert.Equal(@"C:\Drop", s2.DropFolder);
        Assert.Equal(["sec1"], s2.RecentSectionIds);
    }

    [Fact]
    public void CorruptSettingsFallBackToDefaults()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(Path.Combine(_dir, "settings.json"), "{not json!!");
        var s = new JsonFileStore(_dir).LoadSettings();
        Assert.Null(s.DropFolder);
        Assert.True(s.DeleteOnSuccess);
    }

    [Fact]
    public void TreeCacheRoundTrip()
    {
        var store = new JsonFileStore(_dir);
        Assert.Null(store.LoadTreeCache());
        var tree = new NotebookTree(
            [new NotebookNode("n1", "General",
                [new SectionNode("s1", "Inbox")],
                [new GroupNode("g1", "Taxes", [new SectionNode("s2", "Taxes 2026")], [])])],
            DateTimeOffset.UtcNow);
        store.SaveTreeCache(tree);
        var loaded = new JsonFileStore(_dir).LoadTreeCache();
        Assert.NotNull(loaded);
        Assert.Equal("Taxes 2026", loaded!.Notebooks[0].Groups[0].Sections[0].Name);
    }
}
```

Run: `dotnet test --filter StorageTests`
Expected: FAIL.

- [ ] **Step 2: Implement models and store**

`src/SendToOneNote.Core/OneNote/NotebookModels.cs`: exactly the records from the Interfaces block above.

`src/SendToOneNote.Core/Storage/AppSettings.cs`: exactly the class from the Interfaces block above.

`src/SendToOneNote.Core/Storage/JsonFileStore.cs`:

```csharp
using System.Text.Json;
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Core.Storage;

public sealed class JsonFileStore
{
    private static readonly JsonSerializerOptions Opts = new() { WriteIndented = true };
    public string RootDir { get; }

    public JsonFileStore(string? rootDir = null)
    {
        RootDir = rootDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendToOneNote");
        Directory.CreateDirectory(RootDir);
    }

    public AppSettings LoadSettings() => Load<AppSettings>("settings.json") ?? new AppSettings();
    public void SaveSettings(AppSettings s) => Save("settings.json", s);
    public NotebookTree? LoadTreeCache() => Load<NotebookTree>("cache.json");
    public void SaveTreeCache(NotebookTree t) => Save("cache.json", t);

    private T? Load<T>(string file) where T : class
    {
        var path = Path.Combine(RootDir, file);
        if (!File.Exists(path)) return null;
        try { return JsonSerializer.Deserialize<T>(File.ReadAllText(path), Opts); }
        catch (JsonException) { return null; }
    }

    private void Save<T>(string file, T value)
    {
        var path = Path.Combine(RootDir, file);
        var tmp = path + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(value, Opts));
        File.Move(tmp, path, overwrite: true);
    }
}
```

- [ ] **Step 3: Run tests**

Run: `dotnet test --filter StorageTests`
Expected: PASS (3/3).

- [ ] **Step 4: Commit**

```powershell
git add src tests
git commit -m "feat: JSON settings and notebook-tree cache with atomic writes"
```

---

### Task 8: MsalTokenProvider + Entra registration doc

**Files:**
- Create: `src/SendToOneNote.Core/Auth/ITokenProvider.cs`, `src/SendToOneNote.Core/Auth/MsalTokenProvider.cs`, `docs/entra-app-registration.md`
- Test: manual verification (MSAL cannot be meaningfully unit-tested; everything downstream depends only on `ITokenProvider`)

**Interfaces:**
- Consumes: `JsonFileStore.RootDir` (Task 7) for the token cache location; `AppSettings.ClientIdOverride`.
- Produces:

```csharp
namespace SendToOneNote.Core.Auth;
public interface ITokenProvider
{
    Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default);
    string? SignedInUser { get; } // UPN/email after first acquisition, else null
}
public sealed class MsalTokenProvider : ITokenProvider
{
    public const string DefaultClientId = "00000000-0000-0000-0000-000000000000"; // replaced after Entra registration
    public MsalTokenProvider(string cacheDir, string? clientIdOverride = null, IntPtr parentWindow = default);
}
public sealed class AuthRequiredException(string message) : Exception(message); // thrown when silent fails and interactiveAllowed=false
```

- [ ] **Step 1: Add MSAL packages**

```powershell
dotnet add src/SendToOneNote.Core package Microsoft.Identity.Client
dotnet add src/SendToOneNote.Core package Microsoft.Identity.Client.Broker
dotnet add src/SendToOneNote.Core package Microsoft.Identity.Client.Extensions.Msal
```

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Auth/ITokenProvider.cs`:

```csharp
namespace SendToOneNote.Core.Auth;

public interface ITokenProvider
{
    Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default);
    string? SignedInUser { get; }
}

public sealed class AuthRequiredException(string message) : Exception(message);
```

`src/SendToOneNote.Core/Auth/MsalTokenProvider.cs`:

```csharp
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;

namespace SendToOneNote.Core.Auth;

public sealed class MsalTokenProvider : ITokenProvider
{
    public const string DefaultClientId = "00000000-0000-0000-0000-000000000000";
    private static readonly string[] Scopes = ["User.Read", "Notes.ReadWrite"];

    private readonly IPublicClientApplication _pca;
    private readonly SemaphoreSlim _init = new(1, 1);
    private bool _cacheAttached;
    private readonly string _cacheDir;

    public string? SignedInUser { get; private set; }

    public MsalTokenProvider(string cacheDir, string? clientIdOverride = null,
        IntPtr parentWindow = default)
    {
        _cacheDir = cacheDir;
        _pca = PublicClientApplicationBuilder
            .Create(string.IsNullOrWhiteSpace(clientIdOverride) ? DefaultClientId : clientIdOverride)
            .WithAuthority("https://login.microsoftonline.com/common")
            .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows))
            .WithParentActivityOrWindow(() => parentWindow)
            .Build();
    }

    public async Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default)
    {
        await EnsureCacheAsync();
        var accounts = await _pca.GetAccountsAsync();
        var account = accounts.FirstOrDefault();
        try
        {
            var result = await _pca.AcquireTokenSilent(Scopes, account).ExecuteAsync(ct);
            SignedInUser = result.Account.Username;
            return result.AccessToken;
        }
        catch (MsalUiRequiredException)
        {
            if (!interactiveAllowed)
                throw new AuthRequiredException("Sign-in required. Open SendToOneNote and sign in.");
            var result = await _pca.AcquireTokenInteractive(Scopes).ExecuteAsync(ct);
            SignedInUser = result.Account.Username;
            return result.AccessToken;
        }
    }

    private async Task EnsureCacheAsync()
    {
        if (_cacheAttached) return;
        await _init.WaitAsync();
        try
        {
            if (_cacheAttached) return;
            var props = new StorageCreationPropertiesBuilder("msal_cache.bin", _cacheDir).Build();
            var helper = await MsalCacheHelper.CreateAsync(props);
            helper.RegisterCache(_pca.UserTokenCache);
            _cacheAttached = true;
        }
        finally { _init.Release(); }
    }
}
```

- [ ] **Step 3: Build**

Run: `dotnet build`
Expected: success. (If `WithParentActivityOrWindow(Func<IntPtr>)` has a different overload shape in the current MSAL version, use the documented WPF pattern from Microsoft Learn "Use MSAL.NET with WAM" — adjust only that line.)

- [ ] **Step 4: Write the Entra registration doc**

`docs/entra-app-registration.md` — the click-by-click the repo owner (and BYO users) follow:

```markdown
# Registering the Entra app (one-time, owner or BYO)

1. https://entra.microsoft.com → Identity → Applications → App registrations → New registration.
2. Name: `SendToOneNote`. Supported account types: **Accounts in any organizational
   directory and personal Microsoft accounts**. Redirect URI: leave blank for now. Register.
3. On the app page, copy the **Application (client) ID** — this replaces
   `MsalTokenProvider.DefaultClientId` (owner) or goes into `settings.json` →
   `ClientIdOverride` (BYO users).
4. Authentication → Add a platform → **Mobile and desktop applications** → check
   `https://login.microsoftonline.com/common/oauth2/nativeclient` → also add redirect URI
   `ms-appx-web://microsoft.aad.brokerplugin/{client-id}` (required for WAM broker) → Save.
   Set **Allow public client flows** = Yes.
5. API permissions → Add a permission → Microsoft Graph → Delegated →
   `Notes.ReadWrite` (User.Read is present by default). Do NOT grant admin consent for
   the whole org unless you intend to.
6. Branding & properties → set Publisher domain to your verified domain.
7. Publisher verification (recommended before public release): with a Partner Center
   account whose MPN ID is verified and whose domain matches the publisher domain,
   enter the MPN ID under Branding & properties → Publisher verification → Verify.
```

- [ ] **Step 5: Manual verification (owner)**

After the owner completes the registration and pastes the real client ID into `MsalTokenProvider.DefaultClientId`: run the Task 9 integration smoke test (below) — first run pops the Windows account picker; second run is silent.

- [ ] **Step 6: Commit**

```powershell
git add src docs
git commit -m "feat: MSAL token provider with WAM broker; Entra registration doc"
```

---

### Task 9: OneNoteClient (notebook tree, create page, appends)

**Files:**
- Create: `src/SendToOneNote.Core/OneNote/OneNoteClient.cs`
- Test: `tests/SendToOneNote.Tests/OneNoteClientTests.cs`

**Interfaces:**
- Consumes: `ITokenProvider` (Task 8), `PagePlan`/`OneNoteRequestPart`/`AppendPlan` (Task 6), notebook models (Task 7), `StubHttpHandler` (Task 5).
- Produces:

```csharp
namespace SendToOneNote.Core.OneNote;
public sealed record CreatedPage(string Id, string? ClientUrl, string? WebUrl);
public sealed class OneNoteApiException(int statusCode, string message) : Exception(message)
{ public int StatusCode { get; } = statusCode; }
public sealed class OneNoteClient(SendToOneNote.Core.Auth.ITokenProvider tokens, HttpMessageHandler? handler = null)
{
    public Task<NotebookTree> GetNotebookTreeAsync(CancellationToken ct = default);
    public Task<CreatedPage> CreatePageAsync(string sectionId, SendToOneNote.Core.Pages.PagePlan plan, CancellationToken ct = default);
}
```

- Graph endpoints used (base `https://graph.microsoft.com/v1.0`):
  - `GET /me/onenote/notebooks?$expand=sections($select=id,displayName),sectionGroups($expand=sections($select=id,displayName))&$select=id,displayName`
  - `GET /me/onenote/sectionGroups/{id}/sectionGroups?$expand=sections($select=id,displayName)&$select=id,displayName` (recursion for nested groups)
  - `POST /me/onenote/sections/{id}/pages` — multipart/form-data, part `Presentation` (Content-Type `application/xhtml+xml`) + binary parts by name
  - `PATCH /me/onenote/pages/{id}/content` — multipart/form-data, part `Commands` (Content-Type `application/json`) + binary parts by name
- Follow `@odata.nextLink` on all GETs. Non-2xx → `OneNoteApiException(status, bodyText)`.

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/OneNoteClientTests.cs`:

```csharp
using System.Net;
using System.Text;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

file sealed class FakeTokens : ITokenProvider
{
    public string? SignedInUser => "test@example.com";
    public Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default)
        => Task.FromResult("FAKE_TOKEN");
}

public class OneNoteClientTests
{
    private const string NotebooksJson = """
    {"value":[{"id":"n1","displayName":"General",
      "sections":[{"id":"s1","displayName":"Inbox"}],
      "sectionGroups":[{"id":"g1","displayName":"Taxes",
        "sections":[{"id":"s2","displayName":"Taxes 2026"}]}]}]}
    """;
    private const string EmptyGroupsJson = """{"value":[]}""";
    private const string CreatedJson = """
    {"id":"p1","links":{"oneNoteClientUrl":{"href":"onenote:https://x/p1"},
      "oneNoteWebUrl":{"href":"https://x/p1"}}}
    """;

    [Fact]
    public async Task BuildsNotebookTreeWithGroups()
    {
        var stub = new StubHttpHandler(req => new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                req.RequestUri!.AbsolutePath.Contains("sectionGroups") ? EmptyGroupsJson : NotebooksJson,
                Encoding.UTF8, "application/json")
        });
        var tree = await new OneNoteClient(new FakeTokens(), stub).GetNotebookTreeAsync();
        var nb = Assert.Single(tree.Notebooks);
        Assert.Equal("General", nb.Name);
        Assert.Equal("Inbox", Assert.Single(nb.Sections).Name);
        Assert.Equal("Taxes 2026", Assert.Single(Assert.Single(nb.Groups).Sections).Name);
    }

    [Fact]
    public async Task CreatePagePostsMultipartWithPresentationAndParts()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Created)
        { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") });
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>",
            [new OneNoteRequestPart("img0", "image/png", [1, 2, 3])], []);
        var page = await new OneNoteClient(new FakeTokens(), stub).CreatePageAsync("s1", plan);

        Assert.Equal("p1", page.Id);
        Assert.Equal("onenote:https://x/p1", page.ClientUrl);
        var req = Assert.Single(stub.Requests);
        Assert.Equal(HttpMethod.Post, req.Method);
        Assert.Contains("/me/onenote/sections/s1/pages", req.RequestUri!.ToString());
        Assert.Equal("Bearer", req.Headers.Authorization!.Scheme);
        var body = await req.Content!.ReadAsStringAsync();
        Assert.Contains("name=Presentation", body.Replace("\"", ""));
        Assert.Contains("name=img0", body.Replace("\"", ""));
    }

    [Fact]
    public async Task AppendsSentAsPatchPerBatch()
    {
        var responses = new Queue<HttpResponseMessage>([
            new(HttpStatusCode.Created) { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") },
            new(HttpStatusCode.NoContent)]);
        var stub = new StubHttpHandler(_ => responses.Dequeue());
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"replace","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);
        await new OneNoteClient(new FakeTokens(), stub).CreatePageAsync("s1", plan);

        Assert.Equal(2, stub.Requests.Count);
        Assert.Equal(HttpMethod.Patch, stub.Requests[1].Method);
        Assert.Contains("/me/onenote/pages/p1/content", stub.Requests[1].RequestUri!.ToString());
    }

    [Fact]
    public async Task ErrorSurfacesStatusAndBody()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Forbidden)
        { Content = new StringContent("nope") });
        var ex = await Assert.ThrowsAsync<OneNoteApiException>(() =>
            new OneNoteClient(new FakeTokens(), stub).GetNotebookTreeAsync());
        Assert.Equal(403, ex.StatusCode);
        Assert.Contains("nope", ex.Message);
    }
}
```

Run: `dotnet test --filter OneNoteClientTests`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/OneNote/OneNoteClient.cs`:

```csharp
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.OneNote;

public sealed record CreatedPage(string Id, string? ClientUrl, string? WebUrl);

public sealed class OneNoteApiException(int statusCode, string message) : Exception(message)
{
    public int StatusCode { get; } = statusCode;
}

public sealed class OneNoteClient
{
    private const string Base = "https://graph.microsoft.com/v1.0";
    private readonly ITokenProvider _tokens;
    private readonly HttpClient _http;

    public OneNoteClient(ITokenProvider tokens, HttpMessageHandler? handler = null)
    {
        _tokens = tokens;
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(100);
    }

    public async Task<NotebookTree> GetNotebookTreeAsync(CancellationToken ct = default)
    {
        var url = $"{Base}/me/onenote/notebooks?$expand=sections($select=id,displayName)," +
                  "sectionGroups($expand=sections($select=id,displayName))&$select=id,displayName";
        var notebooks = new List<NotebookNode>();
        foreach (var el in await GetPagedValuesAsync(url, ct))
        {
            var groups = new List<GroupNode>();
            if (el.TryGetProperty("sectionGroups", out var sgs))
                foreach (var sg in sgs.EnumerateArray())
                    groups.Add(await BuildGroupAsync(sg, ct));
            notebooks.Add(new NotebookNode(
                el.GetProperty("id").GetString()!,
                el.GetProperty("displayName").GetString() ?? "(unnamed)",
                ReadSections(el), groups));
        }
        return new NotebookTree(notebooks, DateTimeOffset.UtcNow);
    }

    private async Task<GroupNode> BuildGroupAsync(JsonElement sg, CancellationToken ct)
    {
        var id = sg.GetProperty("id").GetString()!;
        var nested = new List<GroupNode>();
        var nestedUrl = $"{Base}/me/onenote/sectionGroups/{id}/sectionGroups" +
                        "?$expand=sections($select=id,displayName)&$select=id,displayName";
        foreach (var child in await GetPagedValuesAsync(nestedUrl, ct))
            nested.Add(await BuildGroupAsync(child, ct));
        return new GroupNode(id, sg.GetProperty("displayName").GetString() ?? "(unnamed)",
            ReadSections(sg), nested);
    }

    private static IReadOnlyList<SectionNode> ReadSections(JsonElement el)
    {
        if (!el.TryGetProperty("sections", out var secs)) return [];
        return secs.EnumerateArray().Select(s => new SectionNode(
            s.GetProperty("id").GetString()!,
            s.GetProperty("displayName").GetString() ?? "(unnamed)")).ToList();
    }

    private async Task<List<JsonElement>> GetPagedValuesAsync(string url, CancellationToken ct)
    {
        var all = new List<JsonElement>();
        string? next = url;
        while (next is not null)
        {
            var doc = JsonDocument.Parse(await SendAsync(
                () => new HttpRequestMessage(HttpMethod.Get, next), ct));
            all.AddRange(doc.RootElement.GetProperty("value").EnumerateArray()
                .Select(e => e.Clone()));
            next = doc.RootElement.TryGetProperty("@odata.nextLink", out var link)
                ? link.GetString() : null;
        }
        return all;
    }

    public async Task<CreatedPage> CreatePageAsync(string sectionId, PagePlan plan,
        CancellationToken ct = default)
    {
        var body = await SendAsync(() =>
        {
            var content = new MultipartFormDataContent();
            var pres = new StringContent(plan.PresentationXhtml, Encoding.UTF8, "application/xhtml+xml");
            content.Add(pres, "Presentation");
            foreach (var p in plan.Parts)
                content.Add(MakeBinary(p), p.Name);
            return new HttpRequestMessage(HttpMethod.Post,
                $"{Base}/me/onenote/sections/{sectionId}/pages") { Content = content };
        }, ct);

        var doc = JsonDocument.Parse(body).RootElement;
        var page = new CreatedPage(
            doc.GetProperty("id").GetString()!,
            Href(doc, "oneNoteClientUrl"), Href(doc, "oneNoteWebUrl"));

        foreach (var append in plan.Appends)
        {
            await SendAsync(() =>
            {
                var content = new MultipartFormDataContent();
                content.Add(new StringContent(append.CommandsJson, Encoding.UTF8, "application/json"),
                    "Commands");
                foreach (var p in append.Parts)
                    content.Add(MakeBinary(p), p.Name);
                return new HttpRequestMessage(HttpMethod.Patch,
                    $"{Base}/me/onenote/pages/{page.Id}/content") { Content = content };
            }, ct);
        }
        return page;

        static string? Href(JsonElement doc, string name) =>
            doc.TryGetProperty("links", out var links) &&
            links.TryGetProperty(name, out var l) &&
            l.TryGetProperty("href", out var h) ? h.GetString() : null;

        static ByteArrayContent MakeBinary(OneNoteRequestPart p)
        {
            var c = new ByteArrayContent(p.Data);
            c.Headers.ContentType = new MediaTypeHeaderValue(p.ContentType);
            return c;
        }
    }

    private async Task<string> SendAsync(Func<HttpRequestMessage> makeRequest, CancellationToken ct)
    {
        var req = makeRequest();
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer",
            await _tokens.GetAccessTokenAsync(interactiveAllowed: false, ct));
        var resp = await _http.SendAsync(req, ct);
        var body = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
            throw new OneNoteApiException((int)resp.StatusCode, body);
        return body;
    }
}
```

- [ ] **Step 3: Run tests**

Run: `dotnet test --filter OneNoteClientTests`
Expected: PASS (4/4).

- [ ] **Step 4: Add the gated integration smoke test**

`tests/SendToOneNote.Tests/IntegrationSmokeTests.cs` — runs ONLY when env var `STN_INTEGRATION=1` (owner's machine, real sign-in, scratch section):

```csharp
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class IntegrationSmokeTests
{
    [SkippableFact]
    public async Task CreatesRealPageInScratchSection()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1",
            "Set STN_INTEGRATION=1 to run against the real Graph API.");
        var tokens = new MsalTokenProvider(Path.Combine(Path.GetTempPath(), "stn-int"));
        var token = await tokens.GetAccessTokenAsync(interactiveAllowed: true);
        Assert.NotEmpty(token);

        var client = new OneNoteClient(tokens);
        var tree = await client.GetNotebookTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections)
            .FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        var page = await client.CreatePageAsync(scratch!.Id, new PagePlan(
            "<html><head><title>Integration smoke</title></head><body><p>hello</p></body></html>",
            [], []));
        Assert.NotEmpty(page.Id);
    }
}
```

Run: `dotnet add tests/SendToOneNote.Tests package Xunit.SkippableFact`, then `dotnet test --filter IntegrationSmokeTests`
Expected: SKIPPED locally and on CI (no env var). The owner runs it with `STN_INTEGRATION=1` after Task 8's registration is done.

- [ ] **Step 5: Commit**

```powershell
git add src tests
git commit -m "feat: OneNoteClient lists notebooks and creates pages with appends"
```

---

### Task 10: DropFolderWatcher

**Files:**
- Create: `src/SendToOneNote.Core/Watch/DropFolderWatcher.cs`, `src/SendToOneNote.Core/Watch/FileReadiness.cs`
- Test: `tests/SendToOneNote.Tests/DropFolderWatcherTests.cs`

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces:

```csharp
namespace SendToOneNote.Core.Watch;
public static class FileReadiness
{
    // Polls every 250 ms until the file opens with exclusive read; false on timeout.
    public static Task<bool> WaitUntilUnlockedAsync(string path, TimeSpan timeout, CancellationToken ct = default);
}
public sealed class DropFolderWatcher : IDisposable
{
    public DropFolderWatcher(string folder);            // creates the folder if missing
    public event Action<string>? EmlReady;              // fired once per settled .eml (full path)
    public event Action<string>? NonEmlIgnored;         // fired once per non-eml file
    public void Start();
    public void Dispose();
}
```

- Dedupe rule: a path currently being handled is not re-raised (Created + Renamed + Changed storms from Explorer/Outlook must yield ONE `EmlReady`).

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/DropFolderWatcherTests.cs`:

```csharp
using SendToOneNote.Core.Watch;

namespace SendToOneNote.Tests;

public class DropFolderWatcherTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-watch-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public async Task RaisesEmlReadyOnceForDroppedFile()
    {
        using var w = new DropFolderWatcher(_dir);
        var hits = new List<string>();
        var signal = new TaskCompletionSource();
        w.EmlReady += p => { lock (hits) hits.Add(p); signal.TrySetResult(); };
        w.Start();

        var path = Path.Combine(_dir, "mail.eml");
        await File.WriteAllTextAsync(path, "From: a@b.c\r\n\r\nhi");
        await Task.WhenAny(signal.Task, Task.Delay(5000));
        await Task.Delay(500); // absorb any duplicate events

        Assert.Equal([path], hits);
    }

    [Fact]
    public async Task IgnoresNonEmlWithNotice()
    {
        using var w = new DropFolderWatcher(_dir);
        string? ignored = null;
        var signal = new TaskCompletionSource();
        w.NonEmlIgnored += p => { ignored = p; signal.TrySetResult(); };
        w.EmlReady += _ => throw new InvalidOperationException("must not fire");
        w.Start();

        await File.WriteAllTextAsync(Path.Combine(_dir, "note.txt"), "hi");
        await Task.WhenAny(signal.Task, Task.Delay(5000));
        Assert.EndsWith("note.txt", ignored);
    }

    [Fact]
    public async Task WaitsForLockedFileToBeReleased()
    {
        Directory.CreateDirectory(_dir);
        var path = Path.Combine(_dir, "locked.eml");
        await File.WriteAllTextAsync(path, "x");
        Task<bool> wait;
        using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            wait = FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(10));
            await Task.Delay(600);
            Assert.False(wait.IsCompleted);
        }
        Assert.True(await wait);
    }

    [Fact]
    public async Task TimesOutOnPermanentLock()
    {
        Directory.CreateDirectory(_dir);
        var path = Path.Combine(_dir, "stuck.eml");
        await File.WriteAllTextAsync(path, "x");
        using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None);
        Assert.False(await FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(1)));
    }
}
```

Run: `dotnet test --filter DropFolderWatcherTests`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Watch/FileReadiness.cs`:

```csharp
namespace SendToOneNote.Core.Watch;

public static class FileReadiness
{
    public static async Task<bool> WaitUntilUnlockedAsync(string path, TimeSpan timeout,
        CancellationToken ct = default)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return true;
            }
            catch (IOException) { await Task.Delay(250, ct); }
            catch (FileNotFoundException) { return false; }
        }
        return false;
    }
}
```

`src/SendToOneNote.Core/Watch/DropFolderWatcher.cs`:

```csharp
using System.Collections.Concurrent;

namespace SendToOneNote.Core.Watch;

public sealed class DropFolderWatcher : IDisposable
{
    private readonly FileSystemWatcher _fsw;
    private readonly ConcurrentDictionary<string, byte> _inFlight = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, byte> _ignoredOnce = new(StringComparer.OrdinalIgnoreCase);

    public event Action<string>? EmlReady;
    public event Action<string>? NonEmlIgnored;

    public DropFolderWatcher(string folder)
    {
        Directory.CreateDirectory(folder);
        _fsw = new FileSystemWatcher(folder) { IncludeSubdirectories = false };
        _fsw.Created += (_, e) => Handle(e.FullPath);
        _fsw.Renamed += (_, e) => Handle(e.FullPath);
    }

    public void Start() => _fsw.EnableRaisingEvents = true;

    private void Handle(string path)
    {
        if (Directory.Exists(path)) return;
        if (!path.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
        {
            if (_ignoredOnce.TryAdd(path, 0)) NonEmlIgnored?.Invoke(path);
            return;
        }
        if (!_inFlight.TryAdd(path, 0)) return;
        _ = Task.Run(async () =>
        {
            try
            {
                if (await FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(30)))
                    EmlReady?.Invoke(path);
            }
            finally { _inFlight.TryRemove(path, out _); }
        });
    }

    public void Dispose() => _fsw.Dispose();
}
```

- [ ] **Step 3: Run tests**

Run: `dotnet test --filter DropFolderWatcherTests`
Expected: PASS (4/4). These tests involve real file-system timing; if flaky on CI, raise the 5000 ms waits to 10000 ms — do not add `Thread.Sleep` in the implementation.

- [ ] **Step 4: Commit**

```powershell
git add src tests
git commit -m "feat: drop-folder watcher with readiness polling and dedupe"
```

---

### Task 11: SectionPickerViewModel (filter + recents logic)

**Files:**
- Create: `src/SendToOneNote.Core/Picker/SectionPickerViewModel.cs`
- Test: `tests/SendToOneNote.Tests/SectionPickerViewModelTests.cs`

**Interfaces:**
- Consumes: `NotebookTree`, `SectionNode`, `GroupNode`, `NotebookNode` (Task 7).
- Produces:

```csharp
namespace SendToOneNote.Core.Picker;
public sealed record PickerItem(string SectionId, string SectionName, string Path); // Path e.g. "TownBackyard » Taxes"
public sealed class SectionPickerViewModel
{
    public SectionPickerViewModel(NotebookTree tree, IReadOnlyList<string> recentSectionIds);
    public IReadOnlyList<PickerItem> AllSections { get; }            // flattened, tree order
    public IReadOnlyList<PickerItem> Filter(string query);           // empty query → recents then all
    public static List<string> PushRecent(List<string> recents, string sectionId, int cap = 10);
}
```

- Filter matching: case-insensitive substring against `SectionName` and `Path`. Empty/whitespace query returns recents (in recency order) followed by all remaining sections.

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/SectionPickerViewModelTests.cs`:

```csharp
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Picker;

namespace SendToOneNote.Tests;

public class SectionPickerViewModelTests
{
    private static NotebookTree Tree() => new(
        [
            new NotebookNode("n1", "General",
                [new SectionNode("s1", "Inbox"), new SectionNode("s2", "Quick Notes")], []),
            new NotebookNode("n2", "Projects",
                [new SectionNode("s3", "AirBnb")],
                [new GroupNode("g1", "Taxes", [new SectionNode("s4", "Taxes 2026")], [])])
        ], DateTimeOffset.UtcNow);

    [Fact]
    public void FlattensWithPaths()
    {
        var vm = new SectionPickerViewModel(Tree(), []);
        Assert.Equal(4, vm.AllSections.Count);
        var taxes = vm.AllSections.Single(s => s.SectionId == "s4");
        Assert.Equal("Projects » Taxes", taxes.Path);
    }

    [Fact]
    public void EmptyQueryPutsRecentsFirst()
    {
        var vm = new SectionPickerViewModel(Tree(), ["s4", "s1"]);
        var items = vm.Filter("");
        Assert.Equal("s4", items[0].SectionId);
        Assert.Equal("s1", items[1].SectionId);
        Assert.Equal(4, items.Count); // no duplicates
    }

    [Fact]
    public void QueryMatchesNameAndPathCaseInsensitive()
    {
        var vm = new SectionPickerViewModel(Tree(), []);
        Assert.Equal(["s4"], vm.Filter("taxes 2").Select(i => i.SectionId));
        Assert.Equal(["s4"], vm.Filter("projects » tax").Select(i => i.SectionId).Distinct());
        Assert.Contains("s3", vm.Filter("PROJECT").Select(i => i.SectionId)); // path match
    }

    [Fact]
    public void PushRecentDedupesAndCaps()
    {
        var r = SectionPickerViewModel.PushRecent(["a", "b"], "b");
        Assert.Equal(["b", "a"], r);
        var many = Enumerable.Range(0, 12).Select(i => $"s{i}").ToList();
        var capped = SectionPickerViewModel.PushRecent(many, "new");
        Assert.Equal(10, capped.Count);
        Assert.Equal("new", capped[0]);
    }
}
```

Run: `dotnet test --filter SectionPickerViewModelTests`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Picker/SectionPickerViewModel.cs`:

```csharp
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Core.Picker;

public sealed record PickerItem(string SectionId, string SectionName, string Path);

public sealed class SectionPickerViewModel
{
    private readonly IReadOnlyList<string> _recents;
    public IReadOnlyList<PickerItem> AllSections { get; }

    public SectionPickerViewModel(NotebookTree tree, IReadOnlyList<string> recentSectionIds)
    {
        _recents = recentSectionIds;
        var items = new List<PickerItem>();
        foreach (var nb in tree.Notebooks)
        {
            items.AddRange(nb.Sections.Select(s => new PickerItem(s.Id, s.Name, nb.Name)));
            foreach (var g in nb.Groups) Walk(g, nb.Name, items);
        }
        AllSections = items;

        static void Walk(GroupNode g, string path, List<PickerItem> items)
        {
            var p = $"{path} » {g.Name}";
            items.AddRange(g.Sections.Select(s => new PickerItem(s.Id, s.Name, p)));
            foreach (var child in g.Groups) Walk(child, p, items);
        }
    }

    public IReadOnlyList<PickerItem> Filter(string query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            var recents = _recents
                .Select(id => AllSections.FirstOrDefault(s => s.SectionId == id))
                .Where(s => s is not null).Cast<PickerItem>().ToList();
            return [.. recents, .. AllSections.Where(s => !_recents.Contains(s.SectionId))];
        }
        var q = query.Trim();
        return AllSections.Where(s =>
            s.SectionName.Contains(q, StringComparison.OrdinalIgnoreCase) ||
            $"{s.Path} » {s.SectionName}".Contains(q, StringComparison.OrdinalIgnoreCase) ||
            s.Path.Contains(q, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    public static List<string> PushRecent(List<string> recents, string sectionId, int cap = 10)
    {
        var r = new List<string> { sectionId };
        r.AddRange(recents.Where(x => x != sectionId));
        return r.Take(cap).ToList();
    }
}
```

- [ ] **Step 3: Run tests**

Run: `dotnet test --filter SectionPickerViewModelTests`
Expected: PASS (4/4).

- [ ] **Step 4: Commit**

```powershell
git add src tests
git commit -m "feat: section picker view-model with search and recents"
```

---

### Task 12: WPF app — picker window, tray, first-run, pipeline

**Files:**
- Create: `src/SendToOneNote/PickerWindow.xaml`, `src/SendToOneNote/PickerWindow.xaml.cs`, `src/SendToOneNote/FirstRunWindow.xaml`, `src/SendToOneNote/FirstRunWindow.xaml.cs`, `src/SendToOneNote/SavePipeline.cs`, `src/SendToOneNote/TrayContext.cs`, `src/SendToOneNote/Logging/FileLog.cs`
- Modify: `src/SendToOneNote/App.xaml`, `src/SendToOneNote/App.xaml.cs` (remove `MainWindow`/`StartupUri`, wire `TrayContext`)

**Interfaces:**
- Consumes: everything — `DropFolderWatcher`, `EmlParser`, `PageXhtmlBuilder`, `ImageResolver`, `PagePlanner`, `OneNoteClient`, `MsalTokenProvider`, `JsonFileStore`, `SectionPickerViewModel`.
- Produces: the runnable app. `SavePipeline.HandleEmlAsync(string path)` is the orchestration entry point.

This task is UI-heavy; automated coverage stays in Core (already done). Each step ends with a manual verification.

- [ ] **Step 1: Add the tray package and app plumbing**

Run: `dotnet add src/SendToOneNote package H.NotifyIcon.Wpf`

`src/SendToOneNote/App.xaml` (no StartupUri):

```xml
<Application x:Class="SendToOneNote.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             ShutdownMode="OnExplicitShutdown">
    <Application.Resources/>
</Application>
```

`src/SendToOneNote/App.xaml.cs`:

```csharp
using System.Windows;

namespace SendToOneNote;

public partial class App : Application
{
    private TrayContext? _tray;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _tray = new TrayContext();
        _tray.Run();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _tray?.Dispose();
        base.OnExit(e);
    }
}
```

`src/SendToOneNote/Logging/FileLog.cs`:

```csharp
namespace SendToOneNote.Logging;

public sealed class FileLog(string dir)
{
    private readonly object _gate = new();

    public void Info(string msg) => Write("INFO", msg);
    public void Error(string msg, Exception? ex = null) =>
        Write("ERROR", ex is null ? msg : $"{msg} :: {ex}");

    private void Write(string level, string msg)
    {
        lock (_gate)
        {
            Directory.CreateDirectory(dir);
            File.AppendAllText(
                Path.Combine(dir, $"stn-{DateTime.Now:yyyyMMdd}.log"),
                $"{DateTime.Now:HH:mm:ss} [{level}] {msg}{Environment.NewLine}");
        }
    }
}
```

Manual check: `dotnet build` succeeds.

- [ ] **Step 2: Picker window**

`src/SendToOneNote/PickerWindow.xaml`:

```xml
<Window x:Class="SendToOneNote.PickerWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Location in OneNote" Width="420" Height="520"
        WindowStartupLocation="CenterScreen" Topmost="True"
        ShowInTaskbar="True">
    <DockPanel Margin="10">
        <TextBlock DockPanel.Dock="Top" Margin="0,0,0,6"
                   Text="Pick a section in which to put the e-mail:"/>
        <TextBox x:Name="SearchBox" DockPanel.Dock="Top" Margin="0,0,0,6"
                 TextChanged="SearchBox_TextChanged" KeyDown="SearchBox_KeyDown"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button Content="OK" Width="80" Margin="0,0,8,0" IsDefault="True" Click="Ok_Click"/>
            <Button Content="Cancel" Width="80" IsCancel="True"/>
        </StackPanel>
        <ListBox x:Name="Results" MouseDoubleClick="Ok_Click">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="{Binding SectionName}" FontWeight="SemiBold"/>
                        <TextBlock Text="{Binding Path, StringFormat='  ({0})'}" Foreground="Gray"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>
    </DockPanel>
</Window>
```

`src/SendToOneNote/PickerWindow.xaml.cs`:

```csharp
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SendToOneNote.Core.Picker;

namespace SendToOneNote;

public partial class PickerWindow : Window
{
    private readonly SectionPickerViewModel _vm;
    public PickerItem? Selected { get; private set; }

    public PickerWindow(SectionPickerViewModel vm, string emailSubject)
    {
        InitializeComponent();
        _vm = vm;
        Title = $"Send to OneNote — {emailSubject}";
        Results.ItemsSource = _vm.Filter("");
        if (Results.Items.Count > 0) Results.SelectedIndex = 0;
        Loaded += (_, _) => SearchBox.Focus();
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        Results.ItemsSource = _vm.Filter(SearchBox.Text);
        if (Results.Items.Count > 0) Results.SelectedIndex = 0;
    }

    private void SearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Down && Results.Items.Count > 0)
        {
            Results.SelectedIndex = Math.Min(Results.SelectedIndex + 1, Results.Items.Count - 1);
            e.Handled = true;
        }
        else if (e.Key == Key.Up && Results.SelectedIndex > 0)
        {
            Results.SelectedIndex--;
            e.Handled = true;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        Selected = Results.SelectedItem as PickerItem;
        if (Selected is null) return;
        DialogResult = true;
    }
}
```

Manual check: temporarily construct the window with a fake tree from `App.OnStartup`, run `dotnet run --project src/SendToOneNote`, verify: opens focused on search, typing filters, arrows move selection, Enter accepts, Esc cancels. Remove the temporary code after checking.

- [ ] **Step 3: First-run window**

`src/SendToOneNote/FirstRunWindow.xaml`:

```xml
<Window x:Class="SendToOneNote.FirstRunWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SendToOneNote setup" Width="460" Height="300"
        WindowStartupLocation="CenterScreen">
    <StackPanel Margin="16">
        <TextBlock TextWrapping="Wrap" Margin="0,0,0,12"
                   Text="Sign in with the Microsoft account whose OneDrive holds your OneNote notebooks, then choose the folder you'll drag emails into."/>
        <Button x:Name="SignInBtn" Content="Sign in to Microsoft" Click="SignIn_Click"
                Width="200" HorizontalAlignment="Left"/>
        <TextBlock x:Name="SignedInAs" Margin="0,6,0,12" Foreground="Gray"/>
        <DockPanel Margin="0,0,0,12">
            <Button DockPanel.Dock="Right" Content="Browse…" Click="Browse_Click" Width="80"/>
            <TextBox x:Name="FolderBox" Margin="0,0,8,0"/>
        </DockPanel>
        <CheckBox x:Name="StartupBox" Content="Start SendToOneNote when Windows starts"
                  IsChecked="True"/>
        <TextBlock TextWrapping="Wrap" Margin="0,10,0,0" Foreground="Gray"
                   Text="Tip: drag the drop folder into Explorer's Quick Access so it's always one drag away."/>
        <Button Content="Finish" Width="100" HorizontalAlignment="Right" Margin="0,14,0,0"
                Click="Finish_Click"/>
    </StackPanel>
</Window>
```

`src/SendToOneNote/FirstRunWindow.xaml.cs`:

```csharp
using System.Windows;
using Microsoft.Win32;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Storage;

namespace SendToOneNote;

public partial class FirstRunWindow : Window
{
    private readonly ITokenProvider _tokens;
    private readonly AppSettings _settings;
    public bool Completed { get; private set; }

    public FirstRunWindow(ITokenProvider tokens, AppSettings settings)
    {
        InitializeComponent();
        _tokens = tokens;
        _settings = settings;
        FolderBox.Text = settings.DropFolder ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "SendToOneNote Drop");
    }

    private async void SignIn_Click(object sender, RoutedEventArgs e)
    {
        SignInBtn.IsEnabled = false;
        try
        {
            await _tokens.GetAccessTokenAsync(interactiveAllowed: true);
            SignedInAs.Text = $"Signed in as {_tokens.SignedInUser}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Sign-in failed");
        }
        finally { SignInBtn.IsEnabled = true; }
    }

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog();
        if (dlg.ShowDialog(this) == true) FolderBox.Text = dlg.FolderName;
    }

    private void Finish_Click(object sender, RoutedEventArgs e)
    {
        if (_tokens.SignedInUser is null)
        {
            MessageBox.Show(this, "Please sign in first.", "SendToOneNote");
            return;
        }
        _settings.DropFolder = FolderBox.Text;
        Directory.CreateDirectory(_settings.DropFolder);
        if (StartupBox.IsChecked == true) CreateStartupShortcut();
        Completed = true;
        Close();
    }

    private static void CreateStartupShortcut()
    {
        var startup = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        var lnk = Path.Combine(startup, "SendToOneNote.lnk");
        var exe = Environment.ProcessPath!;
        dynamic shell = Activator.CreateInstance(
            Type.GetTypeFromProgID("WScript.Shell")!)!;
        var sc = shell.CreateShortcut(lnk);
        sc.TargetPath = exe;
        sc.WorkingDirectory = Path.GetDirectoryName(exe);
        sc.Save();
    }
}
```

Manual check: runs, Browse picks a folder, Finish refuses without sign-in.

- [ ] **Step 4: SavePipeline**

`src/SendToOneNote/SavePipeline.cs`:

```csharp
using System.IO;
using System.Windows;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;
using SendToOneNote.Core.Picker;
using SendToOneNote.Core.Storage;
using SendToOneNote.Logging;

namespace SendToOneNote;

public sealed class SavePipeline(
    JsonFileStore store, OneNoteClient client, FileLog log)
{
    public event Action<string, string?>? Saved;   // (message, onenoteClientUrl)
    public event Action<string>? Failed;           // message

    public async Task HandleEmlAsync(string path)
    {
        try
        {
            ParsedEmail email;
            await using (var s = File.OpenRead(path))
                email = EmlParser.Parse(s);

            var tree = store.LoadTreeCache() ?? await RefreshTreeAsync();
            _ = RefreshTreeAsync(); // background refresh for next time

            var settings = store.LoadSettings();
            var vm = new SectionPickerViewModel(tree, settings.RecentSectionIds);

            PickerItem? pick = null;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var win = new PickerWindow(vm, email.Subject);
                if (win.ShowDialog() == true) pick = win.Selected;
            });
            if (pick is null) { log.Info($"Cancelled: {path}"); return; }

            var xhtml = PageXhtmlBuilder.Build(email);
            var (normalized, images) = await new ImageResolver().ResolveAsync(xhtml, email.InlineImages);
            var plan = PagePlanner.Plan(normalized, images);
            var page = await client.CreatePageAsync(pick.SectionId, plan);

            settings.RecentSectionIds =
                SectionPickerViewModel.PushRecent(settings.RecentSectionIds, pick.SectionId);
            store.SaveSettings(settings);

            if (settings.DeleteOnSuccess) File.Delete(path);
            log.Info($"Saved '{email.Subject}' to {pick.SectionName}");
            Saved?.Invoke($"Saved to {pick.SectionName}", page.ClientUrl);
        }
        catch (Exception ex)
        {
            log.Error($"Failed for {path}", ex);
            MoveToFailed(path);
            Failed?.Invoke(ex switch
            {
                EmlParseException => "That file isn't a readable email.",
                AuthRequiredException => "Sign-in required — open SendToOneNote from the tray.",
                OneNoteApiException o => $"OneNote API error {o.StatusCode}.",
                _ => "Unexpected error — see log."
            });
        }
    }

    private async Task<NotebookTree> RefreshTreeAsync()
    {
        var tree = await client.GetNotebookTreeAsync();
        store.SaveTreeCache(tree);
        return tree;
    }

    private void MoveToFailed(string path)
    {
        try
        {
            var failed = Path.Combine(Path.GetDirectoryName(path)!, "Failed");
            Directory.CreateDirectory(failed);
            File.Move(path, Path.Combine(failed, Path.GetFileName(path)), overwrite: true);
        }
        catch (IOException) { /* leave in place */ }
    }
}
```

Note: `AuthRequiredException` surfaces when the cached token expired and interactive sign-in is needed — the pipeline itself always calls Graph with `interactiveAllowed: false` (via `OneNoteClient`); re-sign-in happens from the tray menu.

- [ ] **Step 5: TrayContext wiring**

`src/SendToOneNote/TrayContext.cs`:

```csharp
using System.Diagnostics;
using System.Windows;
using H.NotifyIcon;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Storage;
using SendToOneNote.Core.Watch;
using SendToOneNote.Logging;

namespace SendToOneNote;

public sealed class TrayContext : IDisposable
{
    private readonly JsonFileStore _store = new();
    private readonly FileLog _log;
    private TaskbarIcon? _icon;
    private DropFolderWatcher? _watcher;
    private MsalTokenProvider? _tokens;
    private SavePipeline? _pipeline;

    public TrayContext() => _log = new FileLog(Path.Combine(_store.RootDir, "logs"));

    public void Run()
    {
        var settings = _store.LoadSettings();
        _tokens = new MsalTokenProvider(_store.RootDir, settings.ClientIdOverride);

        if (settings.DropFolder is null)
        {
            var first = new FirstRunWindow(_tokens, settings);
            first.ShowDialog();
            if (!first.Completed) { Application.Current.Shutdown(); return; }
            _store.SaveSettings(settings);
        }

        var client = new OneNoteClient(_tokens);
        _pipeline = new SavePipeline(_store, client, _log);
        _pipeline.Saved += (msg, url) => Notify("SendToOneNote", msg, url);
        _pipeline.Failed += msg => Notify("SendToOneNote — failed", msg, null);

        _watcher = new DropFolderWatcher(settings.DropFolder!);
        _watcher.EmlReady += p => _ = _pipeline.HandleEmlAsync(p);
        _watcher.NonEmlIgnored += p => _log.Info($"Ignored non-eml: {p}");
        _watcher.Start();

        _icon = new TaskbarIcon { ToolTipText = "SendToOneNote" };
        var menu = new System.Windows.Controls.ContextMenu();
        AddItem(menu, "Open drop folder", () =>
            Process.Start("explorer.exe", settings.DropFolder!));
        AddItem(menu, "Sign in again", async () =>
        {
            try { await _tokens.GetAccessTokenAsync(interactiveAllowed: true); }
            catch (Exception ex) { _log.Error("Interactive sign-in failed", ex); }
        });
        AddItem(menu, "Exit", () => Application.Current.Shutdown());
        _icon.ContextMenu = menu;
        _log.Info("Started");
    }

    private static void AddItem(System.Windows.Controls.ContextMenu menu, string header, Action act)
    {
        var mi = new System.Windows.Controls.MenuItem { Header = header };
        mi.Click += (_, _) => act();
        menu.Items.Add(mi);
    }

    private void Notify(string title, string message, string? url)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            _icon?.ShowNotification(title, message);
            if (url is not null)
                _icon!.TrayBalloonTipClicked += (_, _) =>
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        });
    }

    public void Dispose()
    {
        _icon?.Dispose();
        _watcher?.Dispose();
        _tokens = null;
    }
}
```

(If `ShowNotification`/`TrayBalloonTipClicked` member names differ in the installed H.NotifyIcon.Wpf version, use that version's documented notification API — the contract is: show a toast/balloon with title+message, clicking it opens `url` with the shell.)

- [ ] **Step 6: Build and full manual E2E (owner's machine, after Task 8 registration)**

Run: `dotnet run --project src/SendToOneNote`
Verify in order: first-run appears → sign-in works → finish creates the folder → drag a real email from new Outlook into the folder → picker pops with real notebooks → pick section → toast appears → page exists in OneNote with title, header block, images → .eml gone from folder. Then: drop a .txt (ignored, logged), drop with network off (lands in `Failed\` with toast).

- [ ] **Step 7: Commit**

```powershell
git add src
git commit -m "feat: WPF tray app with picker, first-run, and save pipeline"
```

---

### Task 13: README, release workflow, E2E checklist

**Files:**
- Create: `README.md`, `.github/workflows/release.yml`, `docs/e2e-checklist.md`

**Interfaces:**
- Consumes: the finished app; `docs/entra-app-registration.md` (Task 8).
- Produces: the public face of the repo and the release pipeline.

- [ ] **Step 1: Write README.md**

Must contain these sections (write real prose, not stubs — base it on the spec):

```markdown
# SendToOneNote

Bring classic Outlook's "Send to OneNote" back — for new Outlook, including Gmail
and IMAP accounts where add-ins can't run.

Drag an email from new Outlook into a folder → pick a OneNote section → the email
becomes an editable OneNote page: subject as the title, From/To/Date block, full
HTML body with images embedded.

## Why a folder?
New Outlook doesn't allow add-ins on non-Microsoft accounts and blocks direct
drag-and-drop onto other apps. Dragging to a folder is the one path that works
for every account type. Pin the drop folder to Quick Access and it's one drag +
one click.

## Install
1. Download the latest release zip, unzip, run SendToOneNote.exe.
   (SmartScreen may warn because the exe is unsigned: More info → Run anyway.)
2. Sign in with the Microsoft account whose OneDrive holds your notebooks
   (work/school or personal).
3. Choose your drop folder. Done — drag emails in.

## Company (work/school) accounts
Your organization may require an admin to approve the app once
(Notes.ReadWrite). If sign-in is blocked: ask your admin to consent, or register
your own free app ID (docs/entra-app-registration.md) and put it in
%APPDATA%\SendToOneNote\settings.json as "ClientIdOverride".

## Privacy
Your email content goes from the local .eml file directly to Microsoft Graph
under your own sign-in. Nothing is sent anywhere else. No telemetry.

## Fidelity notes
OneNote's API sanitizes HTML: scripts, forms, and complex CSS are removed, so
heavily designed newsletters simplify. Text, links, tables, and images survive.
Remote images are downloaded and embedded so pages outlive expiring links.

## Troubleshooting
- Drag-out from new Outlook sometimes needs the message clicked first
  (known Outlook regression); "… → Save as" always works.
- Failed saves land in the drop folder's Failed subfolder; logs in
  %APPDATA%\SendToOneNote\logs.

## Building from source
dotnet build / dotnet test / dotnet run --project src/SendToOneNote. MIT license.
```

- [ ] **Step 2: Release workflow**

`.github/workflows/release.yml`:

```yaml
name: release
on:
  push:
    tags: ['v*']
jobs:
  release:
    runs-on: windows-latest
    permissions:
      contents: write
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: 10.0.x
      - run: >
          dotnet publish src/SendToOneNote -c Release -r win-x64
          --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
          -o publish
      - run: Compress-Archive -Path publish/* -DestinationPath SendToOneNote-${{ github.ref_name }}-win-x64.zip
      - uses: softprops/action-gh-release@v2
        with:
          files: SendToOneNote-*.zip
          generate_release_notes: true
```

- [ ] **Step 3: E2E checklist**

`docs/e2e-checklist.md`:

```markdown
# Pre-release manual checklist
- [ ] Drag email from an M365 account in new Outlook → page correct (title, header, body)
- [ ] Drag email from a Gmail account in new Outlook → page correct
- [ ] Image-heavy HTML email → images embedded, page opens in OneNote client from toast
- [ ] Plain-text email → readable paragraphs / preserved columns
- [ ] Inline (cid:) image email → image embedded
- [ ] Esc in picker → .eml stays, nothing created
- [ ] Airplane mode → file lands in Failed with toast; drag back after reconnect succeeds
- [ ] Non-.eml file dropped → ignored, one log line
- [ ] Quit + relaunch → silent auth (no prompt), recents preserved
- [ ] Fresh Windows user / second machine → first-run flow works end to end
```

- [ ] **Step 4: Commit, push, tag v0.1.0 (owner confirms tag)**

```powershell
git add README.md .github docs
git commit -m "docs: README, release workflow, E2E checklist"
git push
```

Then, when the owner says ship: `git tag v0.1.0 && git push origin v0.1.0` — verify the release appears with a zip asset.

---

## Task dependency order

1 → 2 → 3 → 4 → 5 → 6 (Core pipeline, strictly ordered) · 7 anytime after 1 · 8 after 7 · 9 after 6+8 · 10 after 1 · 11 after 7 · 12 after all of 3–11 · 13 last. Owner-manual work (Entra registration, Task 8 Step 5) can run in parallel with Tasks 9–12; the real client ID is only needed for integration smoke and E2E.

## GitHub issues mapping

After plan approval, create one GitHub issue per task (Tasks 1–13), titled `Task N: <name>`, body = the task's Files/Interfaces/steps summary + link to this plan file, labeled `v1`. Create two extra issues labeled `manual`: "Entra app registration + publisher verification" (assignee: owner) and "Pre-release E2E checklist run" (assignee: owner). v2 candidates from the spec get one backlog issue each, labeled `v2`, so they're visible but out of scope.
