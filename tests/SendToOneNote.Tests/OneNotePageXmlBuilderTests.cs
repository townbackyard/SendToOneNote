using System.Xml.Linq;
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class OneNotePageXmlBuilderTests
{
    [Fact]
    public void BuildsPageWithTitleAndHtmlBlock()
    {
        var xml = OneNotePageXmlBuilder.Build("{P1}{1}{B0}", "Hello <b>", "<html><body><p>hi</p></body></html>");
        var doc = XDocument.Parse(xml);
        XNamespace one = OneNoteConstants.Namespace2013;
        var page = doc.Root!;
        Assert.Equal(one + "Page", page.Name);
        Assert.Equal("{P1}{1}{B0}", page.Attribute("ID")!.Value);
        Assert.Equal("Hello <b>", page.Element(one + "Title")!.Element(one + "OE")!.Element(one + "T")!.Value);
        var block = page.Element(one + "Outline")!.Element(one + "OEChildren")!.Element(one + "HTMLBlock")!;
        Assert.Equal("<html><body><p>hi</p></body></html>", block.Element(one + "Data")!.Value);
    }

    [Fact]
    public void CdataTerminatorInsidePayloadIsEscaped()
    {
        var xml = OneNotePageXmlBuilder.Build("{P}", "t", "<p>x]]>y</p>");
        var doc = XDocument.Parse(xml); // must remain well-formed
        XNamespace one = OneNoteConstants.Namespace2013;
        Assert.Equal("<p>x]]>y</p>", doc.Descendants(one + "Data").Single().Value);
    }

    [Fact]
    public void ExtractTitleDecodesEntitiesAndFallsBack()
    {
        Assert.Equal("S & T", OneNotePageXmlBuilder.ExtractTitle("<html><head><title>S &amp; T</title></head><body/></html>"));
        Assert.Equal("(no subject)", OneNotePageXmlBuilder.ExtractTitle("<html><body/></html>"));
    }

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
}
