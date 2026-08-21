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
