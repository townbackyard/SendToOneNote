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

    private static ResolvedImage Img(int i, int w, int h) => new($"img{i}", "image/png", Png, w, h);

    [Fact]
    public void ImagesBeyondPartCapAreDroppedNotAppended()
    {
        var images = Enumerable.Range(0, 33).Select(i => Img(i, 100, 100)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(33), images);
        Assert.Equal(30, plan.Parts.Count);
        Assert.Empty(plan.Appends);
        Assert.Equal(3, plan.DroppedPartNames.Count);
        foreach (var name in plan.DroppedPartNames)
            Assert.DoesNotContain($"name:{name}", plan.PresentationXhtml);
        Assert.DoesNotContain("data-id=\"slot-", plan.PresentationXhtml);
    }

    [Fact]
    public void LargerImagesWinWhenOverCap()
    {
        // 31 images: index 5 is huge, everything else tiny — the huge one must survive.
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i == 5 ? 800 : 10, i == 5 ? 600 : 10)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.Contains(plan.Parts, p => p.Name == "img5");
        Assert.Single(plan.DroppedPartNames);
    }

    [Fact]
    public void EarlyImagesGetABoost()
    {
        // 31 equal-size images: the first three (logo/banner position) must never be the ones dropped.
        var images = Enumerable.Range(0, 31).Select(i => Img(i, 50, 50)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.DoesNotContain("img0", plan.DroppedPartNames);
        Assert.DoesNotContain("img1", plan.DroppedPartNames);
        Assert.DoesNotContain("img2", plan.DroppedPartNames);
    }

    [Fact]
    public void DroppedImageTagRemovedRegardlessOfAttributes()
    {
        var imgs = string.Join("", Enumerable.Range(0, 31).Select(i =>
            $"<img alt=\"pic {i}\" width=\"600\" src=\"name:img{i}\" style=\"border:0\" />"));
        var xhtml = $"<html><head><title>t</title></head><body>{imgs}</body></html>";
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i == 30 ? 1 : 50, i == 30 ? 1 : 50)).ToList();
        var plan = PagePlanner.Plan(xhtml, images);
        Assert.Equal(["img30"], plan.DroppedPartNames);
        Assert.DoesNotContain("pic 30", plan.PresentationXhtml);
        Assert.Contains("pic 29", plan.PresentationXhtml);
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
}
