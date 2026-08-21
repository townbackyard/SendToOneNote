using System.Net;
using System.Text;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

file sealed class FakeTokens : ITokenProvider
{
    public string? SignedInUser => "test@example.com";
    public Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default)
        => Task.FromResult("FAKE_TOKEN");
}

public class OneNoteClientTests
{
    private const string NotebooksJson = """
    {"value":[{"id":"n1","displayName":"General",
      "sections":[{"id":"s1","displayName":"Inbox"}],
      "sectionGroups":[{"id":"g1","displayName":"Taxes",
        "sections":[{"id":"s2","displayName":"Taxes 2026"}]}]}]}
    """;
    private const string EmptyGroupsJson = """{"value":[]}""";
    private const string CreatedJson = """
    {"id":"p1","links":{"oneNoteClientUrl":{"href":"onenote:https://x/p1"},
      "oneNoteWebUrl":{"href":"https://x/p1"}}}
    """;

    [Fact]
    public async Task BuildsNotebookTreeWithGroups()
    {
        var stub = new StubHttpHandler(req => new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(
                req.RequestUri!.AbsolutePath.Contains("sectionGroups") ? EmptyGroupsJson : NotebooksJson,
                Encoding.UTF8, "application/json")
        });
        var tree = await new OneNoteClient(new FakeTokens(), stub).GetNotebookTreeAsync();
        var nb = Assert.Single(tree.Notebooks);
        Assert.Equal("General", nb.Name);
        Assert.Equal("Inbox", Assert.Single(nb.Sections).Name);
        Assert.Equal("Taxes 2026", Assert.Single(Assert.Single(nb.Groups).Sections).Name);
    }

    [Fact]
    public async Task CreatePagePostsMultipartWithPresentationAndParts()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Created)
        { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") });
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>",
            [new OneNoteRequestPart("img0", "image/png", [1, 2, 3])], []);
        var page = await new OneNoteClient(new FakeTokens(), stub).CreatePageAsync("s1", plan);

        Assert.Equal("p1", page.Id);
        Assert.Equal("onenote:https://x/p1", page.ClientUrl);
        var req = Assert.Single(stub.Requests);
        Assert.Equal(HttpMethod.Post, req.Method);
        Assert.Contains("/me/onenote/sections/s1/pages", req.RequestUri!.ToString());
        Assert.Equal("Bearer", req.Headers.Authorization!.Scheme);
        var body = await req.Content!.ReadAsStringAsync();
        Assert.Contains("name=Presentation", body.Replace("\"", ""));
        Assert.Contains("name=img0", body.Replace("\"", ""));
    }

    [Fact]
    public async Task AppendsSentAsPatchPerBatch()
    {
        var responses = new Queue<HttpResponseMessage>([
            new(HttpStatusCode.Created) { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") },
            new(HttpStatusCode.NoContent)]);
        var stub = new StubHttpHandler(_ => responses.Dequeue());
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"replace","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);
        await new OneNoteClient(new FakeTokens(), stub).CreatePageAsync("s1", plan);

        Assert.Equal(2, stub.Requests.Count);
        Assert.Equal(HttpMethod.Patch, stub.Requests[1].Method);
        Assert.Contains("/me/onenote/pages/p1/content", stub.Requests[1].RequestUri!.ToString());
    }

    [Fact]
    public async Task ErrorSurfacesStatusAndBody()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Forbidden)
        { Content = new StringContent("nope") });
        var ex = await Assert.ThrowsAsync<OneNoteApiException>(() =>
            new OneNoteClient(new FakeTokens(), stub).GetNotebookTreeAsync());
        Assert.Equal(403, ex.StatusCode);
        Assert.Contains("nope", ex.Message);
    }
}
