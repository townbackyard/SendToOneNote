# SendToOneNote v2 — Embed Email Attachments (issue #16)

Date: 2026-09-15
Status: Accepted 2026-09-15 (owner brainstorm); implemented by `docs/superpowers/plans/2026-09-15-attachments.md`
Builds on: `2026-08-28-desktop-onenote-backend-design.md` (backend seam, page XML, Graph caps)

## Problem

Since 8f920c9 the parser lists every attachment by name in the page's header block, but the bytes never
leave the .eml. A receipt with a PDF invoice becomes a page that *says* "Attachments: invoice.pdf" and
nothing more. Classic Outlook's Send-to-OneNote button embedded the files on the page; this feature
restores that on both backends.

## Goals

- Every ordinary file attachment becomes a real file object on the OneNote page, on the desktop (COM)
  path and the Graph path alike, positioned directly under the header block.
- A page never fails because of an attachment. Files that cannot be embedded (over the size cap, or
  over Graph's request budget) are skipped with a visible note; the page is still created.
- Default on, with a settings.json kill switch. No UI.

## Non-goals (tracked separately)

- Rendering image attachments as pictures on the page rather than file icons — new backlog issue.
- PDF printouts (`data-render-src` on Graph; the COM equivalent) — new backlog issue.
- Embedding attached email messages (`message/rfc822`) — listed by name only.
- A per-save checkbox in the picker; any settings UI.

## Decisions (owner, 2026-09-15)

| Question | Decision |
|---|---|
| Always vs opt-in | Setting `IncludeAttachments`, default **true**; no picker checkbox |
| Size cap | Per-attachment `MaxAttachmentBytes`, default 25 MB (26,214,400); exceeded → skip + note on page |
| Types | Embed every `MimePart` attachment regardless of type; attached messages list-only, never embedded |
| Graph PDF printouts | No — attach only, same page shape on both backends |
| Placement | After the header table, before the body |
| Data flow | Bundle page inputs in a `PageContent` record; the backend seam takes that record |

## Data model (Core/Email)

```csharp
public sealed record EmailAttachment(string FileName, string ContentType, byte[] Data);

public sealed record ParsedEmail(
    string Subject, string From, string To, string? Cc, DateTimeOffset? SentDate,
    string? HtmlBody, string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<EmailAttachment> Attachments,        // files with bytes; excludes cid-referenced inline images
    IReadOnlyList<string> AttachedMessageNames)        // message/rfc822 parts: subject, else file name, else "attached message"
{
    /// <summary>Header-row list: file names first, then attached-message names.</summary>
    public IReadOnlyList<string> AttachmentNames => [.. Attachments.Select(a => a.FileName), .. AttachedMessageNames];
}
```

`EmlParser` fills `Attachments` from `msg.Attachments.OfType<MimePart>()` using the same cid-exclusion
rule that exists today (a part referenced by `cid:` from the HTML body is inline, not an attachment),
decoding each part's content. `ContentType` is `part.ContentType.MimeType`, falling back to
`application/octet-stream` when empty; `FileName` falls back to `attachment`. `AttachedMessageNames`
comes from `msg.Attachments.OfType<MessagePart>()`. Existing behaviour for everything else is unchanged.

## Settings (Core/Storage/AppSettings)

- `IncludeAttachments` — bool, default `true`.
- `MaxAttachmentBytes` — long, default `26_214_400` (25 MB). Applies per attachment on both backends.

Existing settings files without these keys get the defaults; nothing else changes.

## Attachment planning (Core/Pages/AttachmentPlanner)

Pure, static, unit-tested:

```csharp
public sealed record PlannedAttachment(string PartName, EmailAttachment Attachment);   // PartName = "att1", "att2", … in email order
public sealed record OmittedAttachment(string FileName, long Bytes, string Reason);    // Reason: "over the size limit"
public sealed record AttachmentPlan(IReadOnlyList<PlannedAttachment> Embedded, IReadOnlyList<OmittedAttachment> Omitted)
{
    public static AttachmentPlan Empty { get; } = new([], []);
}

public static AttachmentPlan Plan(ParsedEmail email, bool includeAttachments, long maxAttachmentBytes);
```

Rules:

- `includeAttachments == false` → `AttachmentPlan.Empty`. The page looks exactly as it does today
  (header row still lists names; no object tags, no notes).
- Otherwise every file whose `Data.Length > maxAttachmentBytes` is omitted with reason
  `over the size limit`; the rest are embedded, numbered `att1…attN` in email order.
- Attached messages are never candidates; the builder notes them separately (see below).

## Page content and the backend seam (Core/Pages, Core/Backends)

```csharp
public sealed record PageContent(string Xhtml, IReadOnlyList<ResolvedImage> Images,
    IReadOnlyList<PlannedAttachment> Attachments);

public interface IOneNoteBackend
{
    string Name { get; }
    Task<NotebookTree> GetTreeAsync(CancellationToken ct = default);
    Task<CreatedPage> CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default);
}
```

`PageContent` replaces the `(string pageXhtml, IReadOnlyList<ResolvedImage> images)` pair everywhere:
both backends, `PagePlanner.Plan`, the test fakes, and both gated smoke tests. `ImageResolver` is
unchanged; `SavePipeline` assembles the record from the resolver output and the attachment plan.

## Page markup (Core/Pages/PageXhtmlBuilder)

`PageXhtmlBuilder.Build(ParsedEmail email, AttachmentPlan plan)` inserts an attachment block between the
header table and the existing `<hr/>`:

```html
<div class="stn-attachments">
  <object data-attachment="invoice.pdf" data="name:att1" type="application/pdf"></object>
  <p style="color:#999999">[attachment omitted: scan.tif, 31.4 MB, over the size limit]</p>
  <p style="color:#999999">[attached message not embedded: Re: your order]</p>
</div>
```

- One `<object>` per `plan.Embedded`, in order. **Explicit end tag, never self-closing**: AngleSharp's
  HTML parser (used by `ImageResolver`) treats `<object/>` as an open element and would swallow the rest
  of the page into it.
- One grey `<p>` per `plan.Omitted`, then one per `email.AttachedMessageNames`. Style matches the
  existing `[image omitted: too large]` marker. Sizes formatted as `{0:0.#} MB` (or `KB` under 1 MB).
- All names are HTML-encoded. With `AttachmentPlan.Empty` and no attached messages the wrapper `<div>`
  is omitted entirely, so today's output is byte-identical.
- The header-row "Attachments:" line keeps listing `AttachmentNames` (files + messages) as it does now.

## Graph path (Core/Pages/PagePlanner, Core/OneNote/OneNoteClient)

`PagePlanner.Plan(PageContent content)`:

1. Images are shrunk, ranked, and placed exactly as today (they are readable content and keep priority).
2. Attachments are then walked in email order. Each is kept only while
   `parts.Count < MaxBinaryPartsPerRequest` (30) **and** `Data.Length <= remaining budget` (the same
   `MaxRequestBytes` 3,500,000 budget the images drew from, minus the XHTML bytes and the 4 KB slack).
3. A dropped attachment's `<object … data="name:attN" …></object>` element is replaced with
   `<p style="color:#999999">[attachment omitted: {name}, {size}, too large for OneNote online]</p>`
   and its part name is appended to `DroppedPartNames`. The regex must match the whole element
   regardless of attribute order, like `ImgTagRegex`.
4. Kept attachments become `OneNoteRequestPart(PartName, ContentType, Data)` after the image parts.

`OneNoteClient` needs no change beyond the record: it already emits each part with its own content type
and part name, which is exactly what `data="name:attN"` references.

## Desktop path (Core/Desktop, Core/Backends/DesktopOneNoteBackend)

`DesktopOneNoteBackend.CreatePageAsync`, still entirely on the `StaComWorker`, inside `Guard`:

1. Strip every attachment `<object …></object>` from the XHTML (an HTML block cannot host a file),
   then inline images as today.
2. Write each `content.Attachments[i].Attachment.Data` to
   `%TEMP%\SendToOneNote\<Guid>\<sanitised FileName>` (invalid path chars → `_`; empty → `attachment`;
   duplicate names within one email get a numeric suffix). The folder is created per save.
3. `CreateNewPage`, then `UpdatePageContent` with the XML below, then `GetHyperlinkToObject`.
4. In a `finally`, delete the per-save temp folder. OneNote copies the bytes into the notebook during
   `UpdatePageContent`, so the file is not needed afterwards. A temp-write failure surfaces as the
   normal Failed-folder path (it is an ordinary exception inside the pipeline's `catch`).

`OneNotePageXmlBuilder.Build(string pageId, string title, string html, IReadOnlyList<InsertedFile> files)`
where `InsertedFile(string PathSource, string PreferredName)`; the file outline precedes the HTML block:

```xml
<?xml version="1.0"?>
<one:Page xmlns:one="…/2013/onenote" ID="{pageId}">
  <one:Title><one:OE><one:T><![CDATA[{title}]]></one:T></one:OE></one:Title>
  <one:Outline><one:OEChildren>
    <one:OE><one:InsertedFile pathSource="C:\…\invoice.pdf" preferredName="invoice.pdf"/></one:OE>
  </one:OEChildren></one:Outline>
  <one:Outline><one:OEChildren>
    <one:HTMLBlock><one:Data><![CDATA[{html}]]></one:Data></one:HTMLBlock>
  </one:OEChildren></one:Outline>
</one:Page>
```

Attribute values are XML-escaped (`&`, `<`, `"`). With an empty `files` list the output is byte-for-byte
today's XML (single outline). The existing three-argument `Build` overload is kept and delegates.

**Verify on the owner's machine (gated smoke test):** (a) the icon lands above the HTML block, and
(b) the page still opens the file after the temp file is deleted. If (a) fails, move the file outline
after the HTML block — content is correct either way, only position changes. If (b) fails, that is a
spec change (keep temp files until process exit) and must be surfaced, not quietly patched.

## Pipeline (WPF SavePipeline)

Order: parse → pick section → `AttachmentPlanner.Plan(email, settings.IncludeAttachments,
settings.MaxAttachmentBytes)` → `PageXhtmlBuilder.Build(email, plan)` → `ImageResolver` → assemble
`PageContent` → `backend.CreatePageAsync`. Each omitted attachment is logged with name, size and reason
(names only — never content). The `ImageDiagnostics` call passes the same `PageContent` to
`PagePlanner.Plan`, so `DroppedPartNames` now includes attachments dropped for budget on Graph.

## Testing

Synthetic fixtures (invented `example.com` data only; committed):

- `fixtures/synthetic/pdf-attachment.eml` — HTML body, one `application/pdf` part with
  `Content-Disposition: attachment` **and** a `Content-ID` (mirrors new Outlook's export), body a
  fabricated minimal PDF of a few hundred bytes.
- `fixtures/synthetic/attached-message.eml` — plain body plus one `message/rfc822` attachment whose
  inner subject is invented.

Unit tests (CI-safe):

- `EmlParser`: PDF fixture → one `Attachments` entry with non-empty bytes and `application/pdf`; message
  fixture → `Attachments` empty, `AttachedMessageNames` has the inner subject, `AttachmentNames` lists it;
  `inline-cid-image.eml` still yields no attachments.
- `AttachmentPlanner`: off → `Empty`; cap boundary (equal passes, one byte over omits); numbering.
- `PageXhtmlBuilder`: object element with explicit end tag after the header table and before `<hr/>`;
  omission note with size; attached-message note; empty plan is byte-identical to today's output.
- `PagePlanner`: attachment kept when it fits; dropped (with note + `DroppedPartNames`) when over budget
  or when images already fill 30 parts; images placed before attachments; part order stable.
- `OneNotePageXmlBuilder`: file outline before the HTML block; attribute escaping; no files → identical
  to the three-argument overload.
- `DesktopOneNoteBackend` with `FakeOneNoteApplication`: the fake's `UpdatePageContent` reads
  `pathSource` from the XML and asserts the file exists with the right bytes; after the call the temp
  folder is gone; the HTML block contains no `<object`; temp folder is also gone when the fake throws.
- `OneNoteClient`: a plan with an `att1` part sends a multipart section named `att1` with the
  attachment's content type.
- `LocalFixtureTests` (skips without `fixtures/local/`): the real PDF email parses to exactly one
  attachment with non-zero bytes — counts only, no content assertions.

Gated (`STN_INTEGRATION=1`, owner's machine):

- `IntegrationSmokeTests`: page with the synthetic PDF via Graph succeeds.
- `DesktopIntegrationSmokeTests`: page with the synthetic PDF via COM; `GetPageContent` read-back
  contains `one:InsertedFile`; temp folder absent afterwards.

## Docs to update in the same change

- `agents/knowledge/architecture.md`: pipeline paragraph (planner step, `PageContent`), backend seam
  signature, settings list, temp-file location under App data.
- `agents/rules/privacy.md`: one line — attachment bytes are written to a per-save temp folder for the
  COM import and deleted immediately after; on Graph they travel in the same request as the page.
- `README.md`: feature bullet, the two settings, the "skipped with a note" behaviour.
- `docs/e2e-checklist.md`: PDF email → file icon under the header on both backends; oversize file → note.
- `AGENTS.md`: v2 scope line pointing at this spec and plan.

## Privacy

Attachment bytes go only where the email body already goes: the local desktop OneNote notebook (COM
path) or Microsoft Graph under the user's own token (Graph path). The temp file is on the user's own
machine, created and deleted within one save. Logs record file names and sizes, never content.

## Decisions log

- Default on with a kill switch (owner) — matches the legacy button; opt-in would hide the feature.
- 25 MB per-attachment cap, skip + note (owner) — only outliers are affected; the page never fails.
- Embed every file type; attached messages list-only (owner) — OneNote stores any bytes; nested .eml
  has no useful handler for new-Outlook users.
- No printouts (owner) — identical pages on both backends; printouts and picture-rendering of image
  attachments are separate backlog issues.
- Files under the header (owner) — where the eye lands on receipts and invoices.
- `PageContent` record on the seam (owner) — cleaner than a fourth parameter; one-time mechanical churn.
