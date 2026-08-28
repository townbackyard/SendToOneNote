using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

public static class DataUriInliner
{
    /// <summary>Rewrites src="name:{part}" to a base64 data URI for every resolved image.</summary>
    public static string Inline(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        foreach (var img in images)
        {
            var pattern = $"src=\"name:{Regex.Escape(img.PartName)}\"";
            var replacement = $"src=\"data:{img.ContentType};base64,{Convert.ToBase64String(img.Data)}\"";
            xhtml = Regex.Replace(xhtml, pattern, replacement.Replace("$", "$$"));
        }
        return xhtml;
    }
}
