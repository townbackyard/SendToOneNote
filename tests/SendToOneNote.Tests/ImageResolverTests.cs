using System.Net;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class ImageResolverTests
{
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public async Task RemoteOneByOneImageIsDroppedAsJunk()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/a.png\"></body></html>", []);
        Assert.Empty(images);
        Assert.DoesNotContain("<img", xhtml);
    }

    [Fact]
    public async Task CidImageResolvedFromInlineParts()
    {
        using var bmp = new System.Drawing.Bitmap(40, 30);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        var r = new ImageResolver(new StubHttpHandler(_ => throw new InvalidOperationException("no http expected")));
        var inline = new[] { new InlineImage("photo1@example", "p.png", "image/png", ms.ToArray()) };
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"cid:photo1@example\"></body></html>", inline);
        Assert.Single(images);
        Assert.Contains("name:img0", xhtml);
        Assert.DoesNotContain("cid:", xhtml);
    }

    [Fact]
    public async Task FailedDownloadKeepsOriginalUrl()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/gone.png\"></body></html>", []);
        Assert.Empty(images);
        Assert.Contains("https://x.example/gone.png", xhtml);
    }

    [Fact]
    public async Task ThrowingDownloadKeepsOriginalUrl()
    {
        var stub = new StubHttpHandler(_ => throw new IOException("connection reset"));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/broken.png\"></body></html>", []);
        Assert.Empty(images);
        Assert.Contains("https://x.example/broken.png", xhtml);
    }

    [Fact]
    public async Task OutputIsWellFormedXhtml()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = new ImageResolver(stub);
        var (xhtml, _) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><p>a<br>b</p><img src=\"https://x.example/a.png\"></body></html>", []);
        // Must load as XML (self-closed br/img)
        var doc = System.Xml.Linq.XDocument.Parse(xhtml);
        Assert.NotNull(doc.Root);
    }

    [Fact]
    public async Task TrackingPixelUrlIsDroppedAndRemovedFromHtml()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/open.gif?u=1\"/><p>text</p></body></html>", []);
        Assert.Empty(r.Images);
        Assert.DoesNotContain("<img", r.Xhtml);
        var d = Assert.Single(r.Decisions);
        Assert.Equal("dropped-junk", d.Decision);
        Assert.Contains("tracking", d.Reason);
    }

    [Fact]
    public async Task OneByOneDecodedImageIsDropped()
    {
        // PngBytes is a 1x1 PNG served from an innocent-looking URL.
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/logo.png\"/></body></html>", []);
        Assert.Empty(r.Images);
        Assert.Equal("dropped-junk", Assert.Single(r.Decisions).Decision);
    }

    [Fact]
    public async Task TinyWidthHeightAttributesAreDroppedWithoutDownloading()
    {
        var stub = new StubHttpHandler(_ => throw new InvalidOperationException("must not download"));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/s.png\" width=\"1\" height=\"1\"/></body></html>", []);
        Assert.Empty(r.Images);
        Assert.Equal("dropped-junk", Assert.Single(r.Decisions).Decision);
        Assert.Empty(stub.Requests);
    }

    [Fact]
    public async Task RealImageIsEmbeddedWithDimensionsAndDecision()
    {
        using var bmp = new System.Drawing.Bitmap(40, 30);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(ms.ToArray()));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img alt=\"hero\" src=\"https://x.example/hero.png\"/></body></html>", []);
        var img = Assert.Single(r.Images);
        Assert.Equal((40, 30), (img.Width, img.Height));
        var d = Assert.Single(r.Decisions);
        Assert.Equal("embedded", d.Decision);
        Assert.Equal("img0", d.PartName);
        Assert.Equal("hero", d.Alt);
        Assert.Equal("remote", d.Source);
    }

    [Fact]
    public async Task LegacyResolveAsyncStillReturnsTuple()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(System.Net.HttpStatusCode.NotFound));
        var (xhtml, images) = await new ImageResolver(stub).ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/gone.png\"/></body></html>", []);
        Assert.Empty(images);
        Assert.Contains("gone.png", xhtml);
    }
}
