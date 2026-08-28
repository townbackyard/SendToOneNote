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
        // 31 images: img30 (LAST in document order) is huge, everything else tiny.
        // Ranking: img30 (480000) > img0..img2 (100x4 boost=400) > img3..img29 (100).
        // 30 selected by rank leaves img29 (last of the equal-score tail, by stable
        // sort) as the one dropped. Without ranking, img30 itself would be dropped
        // (it's last in document order and the cap is hit before reaching it).
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i == 30 ? 800 : 10, i == 30 ? 600 : 10)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.Contains(plan.Parts, p => p.Name == "img30");
        Assert.Equal(["img29"], plan.DroppedPartNames);
    }

    [Fact]
    public void EarlyImagesGetABoost()
    {
        // 31 images: img0..img2 are small (score 100, boosted x4 = 400); img3..img30
        // are bigger (score 225, unboosted). With the boost, img0..2 outrank the 225s
        // and img30 (last, unboosted) is the one dropped. Without the boost, img0
        // (lowest score, first among ties) would be dropped instead.
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i < 3 ? 10 : 15, i < 3 ? 10 : 15)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.Equal(["img30"], plan.DroppedPartNames);
    }

    [Fact]
    public void SmallerImageStillFitsWhenLargerOneBlowsTheBudget()
    {
        // img0 and img1 are undecodable 1.75 MB blobs (Score falls back to
        // Data.Length; the shrinker's threshold is MaxRequestBytes/2 = 1.75 MB, so
        // ShrinkIfNeeded passes them through unchanged). img2 is a tiny real PNG.
        // All three are within the first-3 boost, so ranking is img0/img1 (score
        // ~7,000,000) ahead of img2 (score 10,000) — img0 is selected first, then
        // img1 no longer fits the remaining byte budget and must be SKIPPED (not
        // break out of the loop), so img2 — ranked last but small enough — still
        // gets a chance to fit. With `break` instead of `continue`, img2 would also
        // be dropped once img1 fails to fit.
        var images = new List<ResolvedImage>
        {
            new("img0", "image/png", new byte[1_750_000]),
            new("img1", "image/png", new byte[1_750_000]),
            Img(2, 50, 50),
        };
        var plan = PagePlanner.Plan(XhtmlWith(3), images);
        Assert.Equal(["img0", "img2"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["img1"], plan.DroppedPartNames);
        Assert.DoesNotContain("name:img1", plan.PresentationXhtml);
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
