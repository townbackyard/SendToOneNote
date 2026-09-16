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

    private static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images) =>
        PagePlanner.Plan(new PageContent(xhtml, images, []));

    [Fact]
    public void FewImagesSingleRequest()
    {
        var plan = Plan(XhtmlWith(3), Images(3));
        Assert.Equal(3, plan.Parts.Count);
        Assert.Empty(plan.Appends);
        Assert.Contains("name:img2", plan.PresentationXhtml);
    }

    private static ResolvedImage Img(int i, int w, int h) => new($"img{i}", "image/png", Png, w, h);

    [Fact]
    public void ImagesBeyondPartCapAreDroppedNotAppended()
    {
        var images = Enumerable.Range(0, 33).Select(i => Img(i, 100, 100)).ToList();
        var plan = Plan(XhtmlWith(33), images);
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
        var plan = Plan(XhtmlWith(31), images);
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
        var plan = Plan(XhtmlWith(31), images);
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
        var plan = Plan(XhtmlWith(3), images);
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
        var plan = Plan(xhtml, images);
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
        var plan = Plan(XhtmlWith(1), [new ResolvedImage("img0", "image/png", big)]);
        Assert.Empty(plan.Parts);
        Assert.Empty(plan.Appends);
        Assert.DoesNotContain("name:img0", plan.PresentationXhtml);
        Assert.Contains("image omitted", plan.PresentationXhtml);
    }

    private static PlannedAttachment Att(int n, int bytes, string name = "f.pdf") =>
        new($"att{n}", new SendToOneNote.Core.Email.EmailAttachment(name, "application/pdf", new byte[bytes]));

    private static string XhtmlWith(int images, params PlannedAttachment[] atts)
    {
        var objs = string.Concat(atts.Select(AttachmentMarkup.ObjectElement));
        var imgs = string.Join("", Enumerable.Range(0, images).Select(i => $"<img src=\"name:img{i}\"/>"));
        return $"<html><head><title>t</title></head><body>{objs}{imgs}</body></html>";
    }

    [Fact]
    public void AttachmentThatFitsBecomesAPartAfterImages()
    {
        var a = Att(1, 1000);
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(2, a), Images(2), [a]));
        Assert.Equal(["img0", "img1", "att1"], plan.Parts.Select(p => p.Name));
        Assert.Equal("application/pdf", plan.Parts[2].ContentType);
        Assert.Contains("data=\"name:att1\"", plan.PresentationXhtml);
        Assert.Empty(plan.DroppedPartNames);
    }

    [Fact]
    public void AttachmentOverRemainingBudgetIsDroppedWithNote()
    {
        // 1.75 MB image + 2 MB attachment > 3.5 MB: the image keeps priority, the file is noted.
        var a = Att(1, 2_000_000, "big.pdf");
        var images = new List<ResolvedImage> { new("img0", "image/png", new byte[1_750_000]) };
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(1, a), images, [a]));
        Assert.Equal(["img0"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["att1"], plan.DroppedPartNames);
        Assert.DoesNotContain("name:att1", plan.PresentationXhtml);
        Assert.Contains("[attachment omitted: big.pdf, 1.9 MB, too large for OneNote online]", plan.PresentationXhtml);
    }

    [Fact]
    public void AttachmentBeyondPartCapIsDropped()
    {
        var a = Att(1, 10);
        var images = Enumerable.Range(0, 30).Select(i => Img(i, 100, 100)).ToList();
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(30, a), images, [a]));
        Assert.Equal(30, plan.Parts.Count);
        Assert.DoesNotContain(plan.Parts, p => p.Name == "att1");
        Assert.Equal(["att1"], plan.DroppedPartNames);
    }

    [Fact]
    public void SmallerAttachmentStillFitsAfterLargerOneIsDropped()
    {
        var big = Att(1, 3_600_000, "big.pdf"); // alone exceeds the 3.5 MB budget
        var small = Att(2, 100, "small.pdf");
        var plan = PagePlanner.Plan(new PageContent(XhtmlWith(0, big, small), [], [big, small]));
        Assert.Equal(["att2"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["att1"], plan.DroppedPartNames);
    }

    [Fact]
    public async Task AttachmentNameWithAngleBracketIsDroppedCleanlyAfterXhtmlNormalization()
    {
        // Reproduces AngleSharp's raw '>' in the attribute value after XHTML normalization
        // (see PageXhtmlBuilderTests.ObjectRegexSurvivesXhtmlNormalizationWhenFileNameContainsAngleBracket).
        // The drop path here uses the same ObjectRegex and must still find and replace the whole element.
        var a = Att(1, 2_000_000, "a>b.pdf");
        var images = new List<ResolvedImage> { new("img0", "image/png", new byte[1_750_000]) };
        var resolved = await new ImageResolver().ResolveWithReportAsync(XhtmlWith(1, a), []);
        var plan = PagePlanner.Plan(new PageContent(resolved.Xhtml, images, [a]));
        Assert.Equal(["img0"], plan.Parts.Select(p => p.Name));
        Assert.Equal(["att1"], plan.DroppedPartNames);
        Assert.DoesNotContain("name:att1", plan.PresentationXhtml);
        Assert.Contains("[attachment omitted:", plan.PresentationXhtml);
        Assert.Contains("too large for OneNote online", plan.PresentationXhtml);
    }
}
