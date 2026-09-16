using System.Net;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Desktop;

public static class OneNotePageXmlBuilder
{
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

    public static string CdataEscape(string s) => s.Replace("]]>", "]]]]><![CDATA[>");

    public static string ExtractTitle(string pageXhtml)
    {
        var m = Regex.Match(pageXhtml, "<title>(.*?)</title>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        var raw = m.Success ? WebUtility.HtmlDecode(m.Groups[1].Value).Trim() : "";
        return raw.Length == 0 ? "(no subject)" : raw;
    }
}
