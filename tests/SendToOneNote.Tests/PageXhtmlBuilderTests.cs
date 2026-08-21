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
