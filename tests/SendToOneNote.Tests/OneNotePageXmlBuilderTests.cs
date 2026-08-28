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
}
