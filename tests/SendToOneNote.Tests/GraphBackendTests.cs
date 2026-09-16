using System.Net;
using System.Text;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

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
        var att = new PlannedAttachment("att1", new EmailAttachment("a.pdf", "application/pdf", [1, 2, 3]));
        var xhtml = "<html><head><title>t</title></head><body>" + AttachmentMarkup.ObjectElement(att) +
                    "<img src=\"name:img0\"/></body></html>";
        var page = await backend.CreatePageAsync("s1", new PageContent(xhtml, [new("img0", "image/png", StubPng())], [att]));
        Assert.Equal("p1", page.Id);
        Assert.Equal("graph", backend.Name);
        var req = Assert.Single(stub.Requests);
        Assert.Contains("/me/onenote/sections/s1/pages", req.RequestUri!.ToString());
        var body = await req.Content!.ReadAsStringAsync();
        Assert.Contains("name=att1", body.Replace("\"", ""));
    }

    private static byte[] StubPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
}
