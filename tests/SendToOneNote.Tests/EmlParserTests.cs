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

    [Fact]
    public void ExtractsAttachmentNames()
    {
        var raw = """
            From: sender@example.com
            To: recipient@example.com
            Subject: Test with attachment
            MIME-Version: 1.0
            Content-Type: multipart/mixed; boundary="boundary123"

            --boundary123
            Content-Type: text/plain

            This is the body of the email.
            --boundary123
            Content-Type: application/pdf
            Content-Disposition: attachment; filename="report.pdf"
            Content-Transfer-Encoding: base64

            JVBERi0xLjQKJeLjz9MNCiAxIDAgb2JqCjw8L1R5cGUvQ2F0YWxvZy9QYWdlcyAyIDAgUj4+CmVu
            ZG9iCjIgMCBvYmoKPDwvVHlwZS9QYWdlcwovS2lkc1szIDAgUl0vQ291bnQgMT4+CmVuZG9iCjMgMCBv
            YmoKPDwvVHlwZS9QYWdlL1BhcmVudCAyIDAgUi9SZXNvdXJjZXM8PC9Gb250PDwvRjE8PC9UeXBlL0Zv
            bnQvU3VidHlwZS9UeXBlMS9CYXNlRm9udC9IZWx2ZXRpY2E+Pj4+Pj4vUHJvY1NldFsvUERGL1RleHRd
            Pj4vTWVkaWFCb3hbMCAwIDYxMiA3OTJdL0NvbnRlbnRzIDQgMCBSPj4KZW5kb2IKNCAwIG9iago8PC9M
            ZW5ndGggNDQvRmlsdGVyL0ZsYXRlRGVjb2RlPj4Kc3RyZWFtCnicLcktEMAgEEDRM0N1TEJqEqxYAAAs
            P+ADQhiQfvyBVF2UWjuYaKOxJJSEUIqFBTnVCwmSVAEKZW5kc3RyZWFtCmVuZG9iCnhyZWYKMCAxMDAwMDAwMDAw
            IDAwMDAwIG4gCjAwMDAwMDAwMDkgMDAwMDAgbiAKMDAwMDAwMDA3NCAwMDAwMCBuIAowMDAwMDAwMTcwIDAwMDAwIG4g
            CjAwMDAwMDAyOTkgMDAwMDAgbiAKdHJhaWxlcgo8PC9TaXplIDUvUm9vdCAxIDAgUj4+CnN0YXJ0eHJlZgo0NTIK
            JSVFT0YK
            --boundary123--
            """;
        var e = EmlParser.Parse(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(raw)));
        Assert.Equal(["report.pdf"], e.AttachmentNames);
        Assert.Empty(e.InlineImages);
    }
}
