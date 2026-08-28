using System.Net;
using System.Text;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Tests;

file sealed class FakeTokens : ITokenProvider
{
    public string? SignedInUser => "test@example.com";
    public Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default) => Task.FromResult("T");
}

public class GraphBackendTests
{
    [Fact]
    public async Task CreatePagePlansAndPostsThroughOneNoteClient()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Created)
        {
            Content = new StringContent("""{"id":"p1","links":{"oneNoteClientUrl":{"href":"onenote:x"}}}""", Encoding.UTF8, "application/json")
        });
        var backend = new GraphBackend(new OneNoteClient(new FakeTokens(), stub));
        var page = await backend.CreatePageAsync("s1",
            "<html><head><title>t</title></head><body><img src=\"name:img0\"/></body></html>",
            [new("img0", "image/png", StubPng())]);
        Assert.Equal("p1", page.Id);
        Assert.Equal("graph", backend.Name);
        var req = Assert.Single(stub.Requests);
        Assert.Contains("/me/onenote/sections/s1/pages", req.RequestUri!.ToString());
    }

    private static byte[] StubPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
}
