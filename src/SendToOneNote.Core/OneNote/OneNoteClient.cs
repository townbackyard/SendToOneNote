using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.OneNote;

public sealed record CreatedPage(string Id, string? ClientUrl, string? WebUrl);

public sealed class OneNoteApiException(int statusCode, string message) : Exception(message)
{
    public int StatusCode { get; } = statusCode;
}

public sealed class OneNoteClient
{
    private const string Base = "https://graph.microsoft.com/v1.0";
    // Large multipart pages (many images) can take well over 20s to become
    // addressable after creation; 8 attempts at linear 2s backoff waits ~56s.
    private const int MaxAppendAttempts = 8;
    private readonly ITokenProvider _tokens;
    private readonly HttpClient _http;
    private readonly TimeSpan _appendRetryBaseDelay;

    public OneNoteClient(ITokenProvider tokens, HttpMessageHandler? handler = null,
        TimeSpan? appendRetryBaseDelay = null)
    {
        _tokens = tokens;
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(100);
        _appendRetryBaseDelay = appendRetryBaseDelay ?? TimeSpan.FromSeconds(2);
    }

    public async Task<NotebookTree> GetNotebookTreeAsync(CancellationToken ct = default)
    {
        var url = $"{Base}/me/onenote/notebooks?$expand=sections($select=id,displayName)," +
                  "sectionGroups($expand=sections($select=id,displayName))&$select=id,displayName";
        var notebooks = new List<NotebookNode>();
        foreach (var el in await GetPagedValuesAsync(url, ct))
        {
            var groups = new List<GroupNode>();
            if (el.TryGetProperty("sectionGroups", out var sgs))
                foreach (var sg in sgs.EnumerateArray())
                    groups.Add(await BuildGroupAsync(sg, ct));
            notebooks.Add(new NotebookNode(
                el.GetProperty("id").GetString()!,
                el.GetProperty("displayName").GetString() ?? "(unnamed)",
                ReadSections(el), groups));
        }
        return new NotebookTree(notebooks, DateTimeOffset.UtcNow);
    }

    private async Task<GroupNode> BuildGroupAsync(JsonElement sg, CancellationToken ct)
    {
        var id = sg.GetProperty("id").GetString()!;
        var nested = new List<GroupNode>();
        var nestedUrl = $"{Base}/me/onenote/sectionGroups/{id}/sectionGroups" +
                        "?$expand=sections($select=id,displayName)&$select=id,displayName";
        foreach (var child in await GetPagedValuesAsync(nestedUrl, ct))
            nested.Add(await BuildGroupAsync(child, ct));
        return new GroupNode(id, sg.GetProperty("displayName").GetString() ?? "(unnamed)",
            ReadSections(sg), nested);
    }

    private static IReadOnlyList<SectionNode> ReadSections(JsonElement el)
    {
        if (!el.TryGetProperty("sections", out var secs)) return [];
        return secs.EnumerateArray().Select(s => new SectionNode(
            s.GetProperty("id").GetString()!,
            s.GetProperty("displayName").GetString() ?? "(unnamed)")).ToList();
    }

    private async Task<List<JsonElement>> GetPagedValuesAsync(string url, CancellationToken ct)
    {
        var all = new List<JsonElement>();
        string? next = url;
        while (next is not null)
        {
            using var doc = JsonDocument.Parse(await SendAsync(
                () => new HttpRequestMessage(HttpMethod.Get, next), ct));
            all.AddRange(doc.RootElement.GetProperty("value").EnumerateArray()
                .Select(e => e.Clone()));
            next = doc.RootElement.TryGetProperty("@odata.nextLink", out var link)
                ? link.GetString() : null;
        }
        return all;
    }

    // Flakes seen against the live service, especially on large multipart
    // requests. Retrying risks a rare duplicate if a timed-out request actually
    // landed — accepted: a duplicate beats a failed save, and a manual re-drag
    // carries the same risk.
    //   404/20102 — page created moments ago, not yet addressable
    //   409/30103 — per-location write throttle on rapid sequential writes
    private static bool IsRetryable(OneNoteApiException ex) =>
        ex.StatusCode is 429 or 502 or 503 or 504
        || (ex.StatusCode == 404 && ex.Message.Contains("20102"))
        || (ex.StatusCode == 409 && ex.Message.Contains("30103"));

    public async Task<CreatedPage> CreatePageAsync(string sectionId, PagePlan plan,
        CancellationToken ct = default)
    {
        string body;
        for (var attempt = 1; ; attempt++)
        {
            try
            {
                body = await SendAsync(() =>
                {
                    var content = new MultipartFormDataContent();
                    var pres = new StringContent(plan.PresentationXhtml, Encoding.UTF8, "application/xhtml+xml");
                    content.Add(pres, "Presentation");
                    foreach (var p in plan.Parts)
                        content.Add(MakeBinary(p), p.Name);
                    return new HttpRequestMessage(HttpMethod.Post,
                        $"{Base}/me/onenote/sections/{sectionId}/pages") { Content = content };
                }, ct);
                break;
            }
            catch (OneNoteApiException ex) when (attempt < 3 && IsRetryable(ex))
            {
                await Task.Delay(_appendRetryBaseDelay * attempt, ct);
            }
        }

        string id;
        string? clientUrl, webUrl;
        using (var doc = JsonDocument.Parse(body))
        {
            var root = doc.RootElement;
            id = root.GetProperty("id").GetString()!;
            clientUrl = Href(root, "oneNoteClientUrl");
            webUrl = Href(root, "oneNoteWebUrl");
        }
        var page = new CreatedPage(id, clientUrl, webUrl);

        // Large multipart pages can take minutes to become addressable
        // (404/20102) — poll a cheap GET before attempting any writes.
        if (plan.Appends.Count > 0)
            await WaitUntilPageAddressableAsync(page.Id, ct);

        foreach (var append in plan.Appends)
        {
            // A freshly created page can briefly 404 (error 20102) until the
            // service has indexed it — retry those with linear backoff.
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    await SendAsync(() =>
                    {
                        var content = new MultipartFormDataContent();
                        content.Add(new StringContent(append.CommandsJson, Encoding.UTF8, "application/json"),
                            "Commands");
                        foreach (var p in append.Parts)
                            content.Add(MakeBinary(p), p.Name);
                        return new HttpRequestMessage(HttpMethod.Patch,
                            $"{Base}/me/onenote/pages/{page.Id}/content") { Content = content };
                    }, ct);
                    break;
                }
                catch (OneNoteApiException ex) when (attempt < MaxAppendAttempts && IsRetryable(ex))
                {
                    await Task.Delay(_appendRetryBaseDelay * attempt, ct);
                }
            }
            // Pace sequential writes to the same page — rapid back-to-back
            // appends trip OneNote's 30103 per-location throttle.
            await Task.Delay(_appendRetryBaseDelay, ct);
        }
        return page;

        async Task WaitUntilPageAddressableAsync(string pageId, CancellationToken ct2)
        {
            const int maxPolls = 36; // × base delay (2s default) ≈ 72s + request time
            for (var poll = 1; ; poll++)
            {
                try
                {
                    await SendAsync(() => new HttpRequestMessage(HttpMethod.Get,
                        $"{Base}/me/onenote/pages/{pageId}?$select=id"), ct2);
                    return;
                }
                catch (OneNoteApiException ex) when (poll < maxPolls &&
                    ex.StatusCode == 404 && ex.Message.Contains("20102"))
                {
                    await Task.Delay(_appendRetryBaseDelay, ct2);
                }
            }
        }

        static string? Href(JsonElement doc, string name) =>
            doc.TryGetProperty("links", out var links) &&
            links.TryGetProperty(name, out var l) &&
            l.TryGetProperty("href", out var h) ? h.GetString() : null;

        static ByteArrayContent MakeBinary(OneNoteRequestPart p)
        {
            var c = new ByteArrayContent(p.Data);
            c.Headers.ContentType = new MediaTypeHeaderValue(p.ContentType);
            return c;
        }
    }

    private async Task<string> SendAsync(Func<HttpRequestMessage> makeRequest, CancellationToken ct)
    {
        var req = makeRequest();
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer",
            await _tokens.GetAccessTokenAsync(interactiveAllowed: false, ct));
        var resp = await _http.SendAsync(req, ct);
        var body = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
            throw new OneNoteApiException((int)resp.StatusCode, body);
        return body;
    }
}
