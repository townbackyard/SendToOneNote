using System.Text.Json;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class PagePlannerTests
{
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private static string XhtmlWith(int n)
    {
        var imgs = string.Join("", Enumerable.Range(0, n).Select(i => $"<img src=\"name:img{i}\"/>"));
        return $"<html><head><title>t</title></head><body>{imgs}</body></html>";
    }

    private static IReadOnlyList<ResolvedImage> Images(int n) =>
        Enumerable.Range(0, n).Select(i => new ResolvedImage($"img{i}", "image/png", Png)).ToList();

    [Fact]
    public void FewImagesSingleRequest()
    {
        var plan = PagePlanner.Plan(XhtmlWith(3), Images(3));
        Assert.Equal(3, plan.Parts.Count);
        Assert.Empty(plan.Appends);
        Assert.Contains("name:img2", plan.PresentationXhtml);
    }

    [Fact]
    public void OverflowImagesBecomeSlotsAndAppends()
    {
        var plan = PagePlanner.Plan(XhtmlWith(8), Images(8));
        Assert.Equal(5, plan.Parts.Count);
        Assert.DoesNotContain("name:img5", plan.PresentationXhtml);
        Assert.Contains("data-id=\"slot-img5\"", plan.PresentationXhtml);
        var append = Assert.Single(plan.Appends);
        Assert.Equal(3, append.Parts.Count);
        var cmds = JsonDocument.Parse(append.CommandsJson).RootElement;
        Assert.Equal(3, cmds.GetArrayLength());
        Assert.Equal("#slot-img5", cmds[0].GetProperty("target").GetString());
        Assert.Equal("replace", cmds[0].GetProperty("action").GetString());
        Assert.Contains("name:img5", cmds[0].GetProperty("content").GetString());
    }

    [Fact]
    public void OversizedImageGetsShrunk()
    {
        // 6 MB of fake bytes is not a decodable image; use a real bitmap instead
        using var bmp = new System.Drawing.Bitmap(3000, 3000);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp); // BMP = huge
        var (data, ct) = ImageShrinker.ShrinkIfNeeded(ms.ToArray(), "image/bmp", 500_000);
        Assert.True(data.Length <= 500_000);
        Assert.Equal("image/jpeg", ct);
    }

    [Fact]
    public void SmallImageUntouched()
    {
        var (data, ct) = ImageShrinker.ShrinkIfNeeded(Png, "image/png", 500_000);
        Assert.Same(Png, data);
        Assert.Equal("image/png", ct);
    }

    [Fact]
    public void UndecodableOversizedImageIsDropped()
    {
        var big = new byte[4_000_000]; // not a decodable image; shrinker passes it through
        var plan = PagePlanner.Plan(XhtmlWith(1), [new ResolvedImage("img0", "image/png", big)]);
        Assert.Empty(plan.Parts);
        Assert.Empty(plan.Appends);
        Assert.DoesNotContain("name:img0", plan.PresentationXhtml);
        Assert.Contains("image omitted", plan.PresentationXhtml);
    }

    [Fact]
    public void OverflowReplacementSurvivesImgAttributes()
    {
        var imgs = string.Join("", Enumerable.Range(0, 8).Select(i =>
            $"<img alt=\"pic {i}\" width=\"600\" src=\"name:img{i}\" style=\"border:0\" />"));
        var xhtml = $"<html><head><title>t</title></head><body>{imgs}</body></html>";
        var plan = PagePlanner.Plan(xhtml, Images(8));
        Assert.Contains("data-id=\"slot-img5\"", plan.PresentationXhtml);
        Assert.DoesNotContain("name:img5", plan.PresentationXhtml);
        Assert.Contains("src=\"name:img4\"", plan.PresentationXhtml); // first-batch imgs keep their tag + attributes
    }
}
