using System.Net;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Desktop;

public static class OneNotePageXmlBuilder
{
    public static string Build(string pageId, string title, string html) =>
        $"""
        <?xml version="1.0"?>
        <one:Page xmlns:one="{OneNoteConstants.Namespace2013}" ID="{pageId}">
          <one:Title><one:OE><one:T><![CDATA[{CdataEscape(title)}]]></one:T></one:OE></one:Title>
          <one:Outline><one:OEChildren>
            <one:HTMLBlock><one:Data><![CDATA[{CdataEscape(html)}]]></one:Data></one:HTMLBlock>
          </one:OEChildren></one:Outline>
        </one:Page>
        """;

    public static string CdataEscape(string s) => s.Replace("]]>", "]]]]><![CDATA[>");

    public static string ExtractTitle(string pageXhtml)
    {
        var m = Regex.Match(pageXhtml, "<title>(.*?)</title>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        var raw = m.Success ? WebUtility.HtmlDecode(m.Groups[1].Value).Trim() : "";
        return raw.Length == 0 ? "(no subject)" : raw;
    }
}
