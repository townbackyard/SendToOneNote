using AngleSharp;
using AngleSharp.Html.Parser;
using AngleSharp.Xhtml;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data);

public sealed class ImageResolver
{
    private readonly HttpClient _http;

    public ImageResolver(HttpMessageHandler? handler = null)
    {
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var parser = new HtmlParser();
        var doc = await parser.ParseDocumentAsync(pageXhtml, ct);
        var imgs = doc.QuerySelectorAll("img").ToList();

        var resolved = new List<ResolvedImage>();
        var gate = new SemaphoreSlim(4);
        var work = imgs.Select(async img =>
        {
            var src = img.GetAttribute("src") ?? "";
            byte[]? data = null;
            string? contentType = null;

            if (src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
            {
                var cid = src[4..].Trim('<', '>');
                var match = inlineImages.FirstOrDefault(i =>
                    string.Equals(i.ContentId, cid, StringComparison.OrdinalIgnoreCase));
                if (match is not null) { data = match.Data; contentType = match.ContentType; }
            }
            else if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                     src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                await gate.WaitAsync(ct);
                try
                {
                    var resp = await _http.GetAsync(src, ct);
                    if (resp.IsSuccessStatusCode)
                    {
                        data = await resp.Content.ReadAsByteArrayAsync(ct);
                        contentType = resp.Content.Headers.ContentType?.MediaType ?? "image/png";
                    }
                }
                catch (Exception) when (!ct.IsCancellationRequested)
                {
                    // leave original src
                }
                finally { gate.Release(); }
            }

            return (img, data, contentType);
        }).ToList();

        var results = await Task.WhenAll(work);
        foreach (var (img, data, contentType) in results)
        {
            if (data is null || contentType is null) continue;
            var name = $"img{resolved.Count}";
            resolved.Add(new ResolvedImage(name, contentType, data));
            img.SetAttribute("src", $"name:{name}");
        }

        var xhtml = doc.ToHtml(XhtmlMarkupFormatter.Instance);
        return (xhtml, resolved);
    }
}
