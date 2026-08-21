using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record AppendPlan(string CommandsJson, IReadOnlyList<OneNoteRequestPart> Parts);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts,
    IReadOnlyList<AppendPlan> Appends);

public static class PagePlanner
{
    public const int MaxRequestBytes = 3_500_000;
    public const int MaxBinaryPartsPerRequest = 5;

    // Matches the whole <img .../> element for a given part name regardless of other
    // attributes or their order (AngleSharp's XHTML serializer may emit alt/width/style/etc.
    // alongside src, and attribute order is not guaranteed).
    private static Regex ImgTagRegex(string partName) =>
        new($"<img\\b[^>]*src=\"name:{Regex.Escape(partName)}\"[^>]*/>");

    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        // Shrink anything that alone would blow the cap.
        var shrunk = images.Select(i =>
        {
            var (data, ct) = ImageShrinker.ShrinkIfNeeded(i.Data, i.ContentType, MaxRequestBytes / 2);
            return new ResolvedImage(i.PartName, ct, data);
        }).ToList();

        // Drop anything STILL over the cap (undecodable blobs the shrinker passed through):
        // a part that can never fit would 413 the request it rides on.
        var kept = new List<ResolvedImage>();
        foreach (var img in shrunk)
        {
            if (img.Data.Length > MaxRequestBytes - 4096)
                xhtml = ImgTagRegex(img.PartName).Replace(xhtml,
                    "<p style=\"color:#999999\">[image omitted: too large]</p>");
            else kept.Add(img);
        }

        // Greedy first batch for the create request.
        var firstBatch = new List<ResolvedImage>();
        long budget = MaxRequestBytes - Encoding.UTF8.GetByteCount(xhtml) - 4096;
        foreach (var img in kept)
        {
            if (firstBatch.Count >= MaxBinaryPartsPerRequest || img.Data.Length > budget) break;
            firstBatch.Add(img);
            budget -= img.Data.Length;
        }

        var overflow = kept.Skip(firstBatch.Count).ToList();
        var presentation = xhtml;
        foreach (var img in overflow)
            presentation = ImgTagRegex(img.PartName).Replace(presentation,
                $"<div data-id=\"slot-{img.PartName}\"></div>");

        var appends = new List<AppendPlan>();
        var batch = new List<ResolvedImage>();
        long batchBytes = 0;
        foreach (var img in overflow)
        {
            if (batch.Count >= MaxBinaryPartsPerRequest ||
                batchBytes + img.Data.Length > MaxRequestBytes - 4096)
            {
                if (batch.Count > 0) appends.Add(ToAppend(batch));
                batch = []; batchBytes = 0;
            }
            batch.Add(img);
            batchBytes += img.Data.Length;
        }
        if (batch.Count > 0) appends.Add(ToAppend(batch));

        return new PagePlan(presentation,
            firstBatch.Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList(),
            appends);
    }

    private static AppendPlan ToAppend(List<ResolvedImage> batch)
    {
        var commands = batch.Select(i => new
        {
            target = $"#slot-{i.PartName}",
            action = "replace",
            content = $"<img src=\"name:{i.PartName}\"/>"
        });
        return new AppendPlan(JsonSerializer.Serialize(commands),
            batch.Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList());
    }
}
