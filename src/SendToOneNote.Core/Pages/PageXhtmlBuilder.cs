using System.Net;
using System.Text;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public static class PageXhtmlBuilder
{
    public static string Build(ParsedEmail email)
    {
        var sb = new StringBuilder();
        sb.Append("<html><head><title>")
          .Append(WebUtility.HtmlEncode(email.Subject))
          .Append("</title></head><body>");
        AppendHeaderTable(sb, email);
        sb.Append("<div>");
        if (email.HtmlBody is not null)
            sb.Append(email.HtmlBody); // normalized to XHTML later by ImageResolver
        else
            AppendTextBody(sb, email.TextBody ?? "");
        sb.Append("</div></body></html>");
        return sb.ToString();
    }

    private static void AppendHeaderTable(StringBuilder sb, ParsedEmail e)
    {
        sb.Append("<table style=\"font-size:10pt;color:#5b5b5b\">");
        Row(sb, "From", e.From);
        Row(sb, "To", e.To);
        if (e.Cc is not null) Row(sb, "Cc", e.Cc);
        if (e.SentDate is { } d) Row(sb, "Sent", d.ToLocalTime().ToString("f"));
        if (e.AttachmentNames.Count > 0)
            Row(sb, "Attachments", string.Join("; ", e.AttachmentNames));
        sb.Append("</table><hr/>");

        static void Row(StringBuilder sb, string label, string value) =>
            sb.Append("<tr><td style=\"font-weight:bold\">").Append(label)
              .Append("</td><td>").Append(WebUtility.HtmlEncode(value))
              .Append("</td></tr>");
    }

    private static void AppendTextBody(StringBuilder sb, string text)
    {
        if (LooksColumnar(text))
        {
            sb.Append("<pre style=\"font-family:Consolas;font-size:10pt\">")
              .Append(WebUtility.HtmlEncode(text)).Append("</pre>");
            return;
        }
        var blocks = text.Replace("\r\n", "\n")
            .Split("\n\n", StringSplitOptions.RemoveEmptyEntries);
        foreach (var block in blocks)
        {
            var lines = block.Split('\n').Select(WebUtility.HtmlEncode);
            sb.Append("<p>").Append(string.Join("<br/>", lines)).Append("</p>");
        }
    }

    public static bool LooksColumnar(string text)
    {
        var lines = text.Replace("\r\n", "\n").Split('\n')
            .Where(l => l.Trim().Length > 0).ToList();
        if (lines.Count == 0) return false;
        var columnar = lines.Count(l => l.Contains("   ")); // 3+ consecutive spaces
        return columnar >= Math.Max(2, lines.Count / 5);
    }
}
