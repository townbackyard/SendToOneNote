using System.Text;

namespace SendToOneNote.Core.Pages;

public static class ImageDiagnosticsWriter
{
    public static string Write(string dropFolder, string emlStem, IReadOnlyList<ImageDecision> decisions,
        IReadOnlyList<ResolvedImage> images, IReadOnlyList<string> droppedMinorPartNames, DateTime now)
    {
        var stem = string.Concat(emlStem.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c)).Trim();
        var folder = Path.Combine(dropFolder, "Diagnostics", $"{stem}-{now:yyyyMMdd-HHmmss}");
        Directory.CreateDirectory(folder);

        var minor = new HashSet<string>(droppedMinorPartNames);
        var sb = new StringBuilder("index,src,part,bytes,width,height,alt,source,decision,reason\n");
        foreach (var d in decisions)
        {
            var decision = d.PartName is not null && minor.Contains(d.PartName) ? "dropped-minor" : d.Decision;
            sb.Append(string.Join(",", new[]
            {
                d.Index.ToString(), Q(d.Src), Q(d.PartName ?? ""), d.Bytes.ToString(), d.Width.ToString(), d.Height.ToString(),
                Q(d.Alt), d.Source, decision, Q(d.Reason)
            })).Append('\n');
        }
        File.WriteAllText(Path.Combine(folder, "images.csv"), sb.ToString(), Encoding.UTF8);

        foreach (var img in images)
            File.WriteAllBytes(Path.Combine(folder, $"{img.PartName}.{Ext(img.ContentType)}"), img.Data);

        return folder;
    }

    private static string Q(string s) => $"\"{s.Replace("\"", "\"\"")}\"";

    private static string Ext(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/jpeg" => "jpg", "image/png" => "png", "image/gif" => "gif", "image/webp" => "webp", _ => "bin"
    };
}
