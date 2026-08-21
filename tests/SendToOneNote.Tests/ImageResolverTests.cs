using System.Net;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class ImageResolverTests
{
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public async Task RemoteImageDownloadedAndRewritten()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = new ImageResolver(stub);
        var (xhtml, images) = await r.ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/a.png\"></body></html>", []);
        var img = Assert.Single(images);
        Assert.Equal("img0", img.PartName);
        Assert.Equal("image/png", img.ContentType);
        Assert.Contains("src=\"name:img0\"", xhtml);
    }

    [Fact]
    public async Task CidImageResolvedFromInlineParts()
    {
        var r = new ImageResolver(new StubHttpHandler(_ => throw new InvalidOperationException("no http expected")));
        var inline = new[] { new InlineImage("photo1@example", "p.png", "image/png", PngBytes) };
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
}
