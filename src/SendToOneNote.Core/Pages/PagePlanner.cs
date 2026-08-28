using System.Text;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record AppendPlan(string CommandsJson, IReadOnlyList<OneNoteRequestPart> Parts);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts,
    IReadOnlyList<AppendPlan> Appends)
{
    public IReadOnlyList<string> DroppedPartNames { get; init; } = [];
}

public static class PagePlanner
{
    public const int MaxRequestBytes = 3_500_000;
    public const int MaxBinaryPartsPerRequest = 30;
    private const int EarlyImageCount = 3;   // logo/banner slots get a ranking boost
    private const int EarlyImageBoost = 4;

    // Must match the whole <img …/> element regardless of other attributes or their order —
    // AngleSharp's XHTML serializer may emit alt/width/style alongside src, and attribute
    // order isn't guaranteed. A miss here leaves a "name:imgN" reference in the XHTML with
    // no matching binary part in the request.
    private static Regex ImgTagRegex(string partName) =>
        new($"<img\\b[^>]*src=\"name:{Regex.Escape(partName)}\"[^>]*/>");

    /// <summary>
    /// Graph path only: everything fits in ONE request or is dropped. Images are ranked by
    /// rendered area (early ones boosted); the desktop backend never calls this.
    /// </summary>
    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        var dropped = new List<string>();

        var shrunk = images.Select(i =>
        {
            var (data, ct) = ImageShrinker.ShrinkIfNeeded(i.Data, i.ContentType, MaxRequestBytes / 2);
            return i with { Data = data, ContentType = ct };
        }).ToList();

        var kept = new List<ResolvedImage>();
        foreach (var img in shrunk)
        {
            if (img.Data.Length > MaxRequestBytes - 4096)
            {
                xhtml = ImgTagRegex(img.PartName).Replace(xhtml, "<p style=\"color:#999999\">[image omitted: too large]</p>");
                dropped.Add(img.PartName);
            }
            else kept.Add(img);
        }

        var ranked = kept
            .Select((img, docIndex) => (img, score: Score(img) * (docIndex < EarlyImageCount ? EarlyImageBoost : 1)))
            .OrderByDescending(x => x.score)
            .Select(x => x.img)
            .ToList();

        var selected = new HashSet<string>();
        long budget = MaxRequestBytes - Encoding.UTF8.GetByteCount(xhtml) - 4096;
        foreach (var img in ranked)
        {
            if (selected.Count >= MaxBinaryPartsPerRequest || img.Data.Length > budget) continue;
            selected.Add(img.PartName);
            budget -= img.Data.Length;
        }

        foreach (var img in kept.Where(i => !selected.Contains(i.PartName)))
        {
            xhtml = ImgTagRegex(img.PartName).Replace(xhtml, "");
            dropped.Add(img.PartName);
        }

        var parts = kept.Where(i => selected.Contains(i.PartName))   // document order for stability
            .Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList();
        return new PagePlan(xhtml, parts, []) { DroppedPartNames = dropped };
    }

    private static long Score(ResolvedImage img)
    {
        if (img.Width > 0 && img.Height > 0) return (long)img.Width * img.Height;
        var (w, h) = ImageShrinker.TryReadDimensions(img.Data);
        return w > 0 && h > 0 ? (long)w * h : img.Data.Length;
    }
}
