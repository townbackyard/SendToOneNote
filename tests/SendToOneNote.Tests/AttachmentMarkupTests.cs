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
        Assert.DoesNotMatch(AttachmentMarkup.ObjectRegex("att1"), xhtml);
    }

    [Fact]
    public void StripAllRemovesEveryAttachmentObjectAndNothingElse()
    {
        var xhtml = "<div><object data-attachment=\"a.pdf\" data=\"name:att1\" type=\"x\"></object>" +
                    "<object data-attachment=\"b.pdf\" data=\"name:att2\" type=\"y\"></object></div><img src=\"name:img0\"/>";
        Assert.Equal("<div></div><img src=\"name:img0\"/>", AttachmentMarkup.StripAll(xhtml));
    }
}
