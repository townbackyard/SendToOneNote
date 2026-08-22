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
    public async Task FollowsNextLinkAcrossPages()
    {
        const string Page1Json = """
        {"value":[{"id":"n1","displayName":"NotebookOne","sections":[]}],
         "@odata.nextLink":"https://graph.microsoft.com/v1.0/me/onenote/notebooks?page2"}
        """;
        const string Page2Json = """
        {"value":[{"id":"n2","displayName":"NotebookTwo","sections":[]}]}
        """;
        var stub = new StubHttpHandler(req =>
        {
            // Check AbsolutePath (not the full URL) for "sectionGroups": the notebooks
            // request's own $expand query string contains that substring too.
            var json = req.RequestUri!.AbsolutePath.Contains("sectionGroups") ? EmptyGroupsJson
                : req.RequestUri!.ToString().Contains("page2") ? Page2Json
                : Page1Json;
            return new HttpResponseMessage(HttpStatusCode.OK)
            { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        });

        var tree = await new OneNoteClient(new FakeTokens(), stub).GetNotebookTreeAsync();

        Assert.Equal(2, tree.Notebooks.Count);
        var names = tree.Notebooks.Select(n => n.Name).ToList();
        Assert.Equal(2, names.Distinct().Count());
        Assert.Contains("NotebookOne", names);
        Assert.Contains("NotebookTwo", names);
        Assert.Contains(stub.Requests, r => r.RequestUri!.ToString().Contains("page2"));
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
            new(HttpStatusCode.OK) { Content = new StringContent("""{"id":"p1"}""", Encoding.UTF8, "application/json") },
            new(HttpStatusCode.NoContent)]);
        var stub = new StubHttpHandler(_ => responses.Dequeue());
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"replace","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);
        await new OneNoteClient(new FakeTokens(), stub, appendRetryBaseDelay: TimeSpan.Zero).CreatePageAsync("s1", plan);

        Assert.Equal(3, stub.Requests.Count); // create + addressability GET + PATCH
        Assert.Equal(HttpMethod.Get, stub.Requests[1].Method);
        Assert.Contains("/me/onenote/pages/p1", stub.Requests[1].RequestUri!.ToString());
        Assert.Equal(HttpMethod.Patch, stub.Requests[2].Method);
        Assert.Contains("/me/onenote/pages/p1/content", stub.Requests[2].RequestUri!.ToString());
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

    private const string NotIndexedYetJson =
        """{"error":{"code":"20102","message":"The specified resource ID does not exist."}}""";

    [Fact]
    public async Task AppendRetriesWhileNewPageIsNotYetAddressable()
    {
        // Live Graph returns 404/20102 until a freshly created page is indexed —
        // the addressability GET polls until 200 before the PATCH runs.
        var responses = new Queue<HttpResponseMessage>([
            new(HttpStatusCode.Created) { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") },
            new(HttpStatusCode.NotFound) { Content = new StringContent(NotIndexedYetJson) },
            new(HttpStatusCode.NotFound) { Content = new StringContent(NotIndexedYetJson) },
            new(HttpStatusCode.OK) { Content = new StringContent("""{"id":"p1"}""", Encoding.UTF8, "application/json") },
            new(HttpStatusCode.NoContent)]);
        var stub = new StubHttpHandler(_ => responses.Dequeue());
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"replace","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);

        var page = await new OneNoteClient(new FakeTokens(), stub, appendRetryBaseDelay: TimeSpan.Zero)
            .CreatePageAsync("s1", plan);

        Assert.Equal("p1", page.Id);
        Assert.Equal(5, stub.Requests.Count); // create + 2 poll misses + poll hit + PATCH
        Assert.Equal(HttpMethod.Patch, stub.Requests[^1].Method);
    }

    [Fact]
    public async Task TransientGatewayErrorsAreRetriedOnCreateAndAppend()
    {
        // Live Graph intermittently 504s on large multipart requests.
        var responses = new Queue<HttpResponseMessage>([
            new(HttpStatusCode.GatewayTimeout) { Content = new StringContent("""{"error":{"code":"UnknownError","message":""}}""") },
            new(HttpStatusCode.Created) { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") },
            new(HttpStatusCode.OK) { Content = new StringContent("""{"id":"p1"}""", Encoding.UTF8, "application/json") },
            new(HttpStatusCode.ServiceUnavailable) { Content = new StringContent("busy") },
            new(HttpStatusCode.Conflict) { Content = new StringContent("""{"error":{"code":"30103","message":"The user account has experienced too many simultaneous requests to the same location."}}""") },
            new(HttpStatusCode.NoContent)]);
        var stub = new StubHttpHandler(_ => responses.Dequeue());
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"append","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);

        var page = await new OneNoteClient(new FakeTokens(), stub, appendRetryBaseDelay: TimeSpan.Zero)
            .CreatePageAsync("s1", plan);

        Assert.Equal("p1", page.Id);
        Assert.Equal(6, stub.Requests.Count); // 504 create, retried create, poll GET, 503 append, 409/30103 append, success
    }

    [Fact]
    public async Task AppendGivesUpAfterRetryCapAndSurfacesError()
    {
        var stub = new StubHttpHandler(req =>
            req.Method == HttpMethod.Post
                ? new HttpResponseMessage(HttpStatusCode.Created)
                { Content = new StringContent(CreatedJson, Encoding.UTF8, "application/json") }
                : req.Method == HttpMethod.Get
                ? new HttpResponseMessage(HttpStatusCode.OK)
                { Content = new StringContent("""{"id":"p1"}""", Encoding.UTF8, "application/json") }
                : new HttpResponseMessage(HttpStatusCode.NotFound)
                { Content = new StringContent(NotIndexedYetJson) });
        var plan = new PagePlan("<html><head><title>t</title></head><body/></html>", [],
            [new AppendPlan("""[{"target":"#slot-img0","action":"replace","content":"<img src=\"name:img0\"/>"}]""",
                [new OneNoteRequestPart("img0", "image/png", [1])])]);

        var ex = await Assert.ThrowsAsync<OneNoteApiException>(() =>
            new OneNoteClient(new FakeTokens(), stub, appendRetryBaseDelay: TimeSpan.Zero)
                .CreatePageAsync("s1", plan));
        Assert.Equal(404, ex.StatusCode);
        Assert.Equal(10, stub.Requests.Count); // create + poll GET + 8 PATCH attempts
    }
}

