using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class PageXhtmlBuilderTests
{
    private static ParsedEmail Email(string? html = null, string? text = null,
        string subject = "S & T <test>", IReadOnlyList<string>? attachments = null) =>
        new(subject, "a@b.c", "d@e.f", null,
            new DateTimeOffset(2026, 8, 20, 18, 0, 35, TimeSpan.Zero),
            html, text, [],
            (attachments ?? []).Select(n => new EmailAttachment(n, "application/octet-stream", [1, 2, 3])).ToList(),
            []);

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

    [Fact]
    public void HtmlBodyPassesThroughUnescaped()
    {
        var x = PageXhtmlBuilder.Build(Email(html: "<p>hi <b>there</b> &amp; friends<br></p>"));
        Assert.Contains("<p>hi <b>there</b> &amp; friends<br></p>", x);
        Assert.DoesNotContain("&lt;p&gt;", x);
    }

    private static ParsedEmail EmailWithMessages(params string[] messageNames) =>
        new("s", "a@b.c", "d@e.f", null, null, "<p>hi</p>", null, [], [], messageNames);

    [Fact]
    public void EmptyPlanKeepsThePreV2TableHrDivLayout()
    {
        // Pins the pre-v2 shape directly (rather than comparing the two Build overloads, which
        // cannot fail since the one-arg overload just delegates to the two-arg one).
        var email = Email(html: "<p>hi</p>", attachments: ["report.pdf"]);
        var x = PageXhtmlBuilder.Build(email, AttachmentPlan.Empty);
        Assert.Contains("</table><hr/><div>", x);
        Assert.DoesNotContain("stn-attachments", x);
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

    [Fact]
    public async Task ObjectRegexSurvivesXhtmlNormalizationWhenFileNameContainsAngleBracket()
    {
        // WebUtility.HtmlEncode escapes '>' as "&gt;" going in, but AngleSharp's XhtmlMarkupFormatter
        // (which ImageResolver runs the page through) escapes only &, <, " in attribute values, so a
        // decoded '>' is re-emitted raw: data-attachment="a>b.pdf". The attribute scan in ObjectRegex
        // and StripAll must be quote-aware so that raw '>' doesn't end the match early.
        var email = Email(html: "<p>body text</p>", attachments: ["a>b.pdf"]);
        var plan = new AttachmentPlan([new PlannedAttachment("att1", email.Attachments[0])], []);
        var r = await new ImageResolver().ResolveWithReportAsync(PageXhtmlBuilder.Build(email, plan), []);
        Assert.True(AttachmentMarkup.ObjectRegex("att1").IsMatch(r.Xhtml), r.Xhtml);
        Assert.DoesNotContain("data=\"name:att1\"", AttachmentMarkup.StripAll(r.Xhtml));
    }
}
