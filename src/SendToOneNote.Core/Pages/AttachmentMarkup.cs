using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

/// <summary>
/// The one place that knows what attachment markup looks like in the page XHTML. The builder emits it,
/// PagePlanner replaces it when Graph's budget is exceeded, and the desktop backend strips it (an
/// HTMLBlock cannot host a file; the file goes in a one:InsertedFile outline instead).
/// </summary>
public static class AttachmentMarkup
{
    public const string GraphBudgetReason = "too large for OneNote online";
    private const string NoteStyle = "color:#999999";

    /// <summary>Graph-canonical embedded-file element. Explicit end tag: AngleSharp's HTML parser treats a
    /// self-closing &lt;object/&gt; as an open element and would swallow the rest of the page into it.</summary>
    public static string ObjectElement(PlannedAttachment a) =>
        $"<object data-attachment=\"{WebUtility.HtmlEncode(a.Attachment.FileName)}\" " +
        $"data=\"name:{a.PartName}\" type=\"{WebUtility.HtmlEncode(a.Attachment.ContentType)}\"></object>";

    public static string OmittedNote(string fileName, long bytes, string reason) =>
        $"<p style=\"{NoteStyle}\">[attachment omitted: {WebUtility.HtmlEncode(fileName)}, {FormatSize(bytes)}, {WebUtility.HtmlEncode(reason)}]</p>";

    public static string AttachedMessageNote(string name) =>
        $"<p style=\"{NoteStyle}\">[attached message not embedded: {WebUtility.HtmlEncode(name)}]</p>";

    public static string FormatSize(long bytes) => bytes switch
    {
        >= 1_048_576 => (bytes / 1_048_576d).ToString("0.#", CultureInfo.InvariantCulture) + " MB",
        >= 1_024 => (bytes / 1_024d).ToString("0.#", CultureInfo.InvariantCulture) + " KB",
        _ => bytes.ToString(CultureInfo.InvariantCulture) + " bytes"
    };

    /// <summary>Matches the whole element for one part regardless of attribute order, whether serialized
    /// as &lt;object …/&gt; or &lt;object …&gt;&lt;/object&gt; (AngleSharp may emit either). The attribute
    /// scan is quote-aware ((?:[^>"]|"[^"]*")*?) rather than a plain [^>]*: AngleSharp's XhtmlMarkupFormatter
    /// escapes only &amp;, &lt;, " in attribute values, so a file name containing '>' is re-emitted raw
    /// (data-attachment="a>b.pdf") and a naive [^>]* stops at that '>' before ever reaching data="name:…".</summary>
    public static Regex ObjectRegex(string partName) =>
        new($"<object\\b(?:[^>\"]|\"[^\"]*\")*?\\bdata=\"name:{Regex.Escape(partName)}\"(?:[^>\"]|\"[^\"]*\")*?(?:/>|>\\s*</object>)", RegexOptions.Singleline);

    private static readonly Regex AnyObject =
        new("<object\\b(?:[^>\"]|\"[^\"]*\")*?\\bdata=\"name:att\\d+\"(?:[^>\"]|\"[^\"]*\")*?(?:/>|>\\s*</object>)", RegexOptions.Singleline | RegexOptions.Compiled);

    public static string StripAll(string xhtml) => AnyObject.Replace(xhtml, "");
}
