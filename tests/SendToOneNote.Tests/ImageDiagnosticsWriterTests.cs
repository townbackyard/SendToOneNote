using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class ImageDiagnosticsWriterTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-diag-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void WritesCsvAndImageFiles()
    {
        var decisions = new List<ImageDecision>
        {
            new(0, "https://x.example/a.png", "img0", 3, 40, 30, "hero, \"big\"", "remote", "embedded", "ok"),
            new(1, "https://x.example/open.gif", null, 0, 0, 0, "", "remote", "dropped-junk", "tracking-pixel URL pattern"),
            new(2, "https://x.example/b.jpg", "img1", 2, 5, 5, "", "remote", "embedded", "ok"),
        };
        var images = new List<ResolvedImage> { new("img0", "image/png", [1, 2, 3], 40, 30), new("img1", "image/jpeg", [4, 5], 5, 5) };

        var folder = ImageDiagnosticsWriter.Write(_dir, "The robotics race", decisions, images, ["img1"],
            new DateTime(2026, 8, 28, 9, 5, 7));

        Assert.Equal(Path.Combine(_dir, "Diagnostics", "The robotics race-20260828-090507"), folder);
        var lines = File.ReadAllLines(Path.Combine(folder, "images.csv"));
        Assert.Equal("index,src,part,bytes,width,height,alt,source,decision,reason", lines[0]);
        Assert.Contains("\"hero, \"\"big\"\"\"", lines[1]);            // CSV quoting
        Assert.Contains(",dropped-junk,", lines[2]);
        Assert.Contains(",dropped-minor,", lines[3]);                  // planner drop overrides "embedded"
        Assert.True(File.Exists(Path.Combine(folder, "img0.png")));
        Assert.True(File.Exists(Path.Combine(folder, "img1.jpg")));
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(folder, "img0.png")));
    }

    [Fact]
    public void SanitizesInvalidFolderCharactersInStem()
    {
        var folder = ImageDiagnosticsWriter.Write(_dir, "Re: fwd/ \"quoted\"?", [], [], [], new DateTime(2026, 1, 1));
        Assert.True(Directory.Exists(folder));
        Assert.DoesNotContain("/", Path.GetFileName(folder));
    }

    [Fact]
    public void FailingImageWriteLeavesNoCsvBehind()
    {
        var decisions = new List<ImageDecision> { new(0, "http://x.example/img", "img0<bad", 1, 1, 1, "", "remote", "embedded", "ok") };
        var images = new List<ResolvedImage> { new("img0<bad", "image/png", [1], 1, 1) };
        var now = new DateTime(2026, 1, 1);

        Assert.Throws<IOException>(() => ImageDiagnosticsWriter.Write(_dir, "test", decisions, images, [], now));
        var folder = Path.Combine(_dir, "Diagnostics", $"test-{now:yyyyMMdd-HHmmss}");
        Assert.True(Directory.Exists(folder));
        Assert.False(File.Exists(Path.Combine(folder, "images.csv")));
    }
}
