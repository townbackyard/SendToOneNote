using System.Text.RegularExpressions;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Xhtml;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data, int Width = 0, int Height = 0);

public sealed class ImageResolver
{
    private static readonly Regex TrackingUrl = new(
        @"(?i)(open|track|pixel|beacon|spacer|blank)\.(gif|png)(\?|$)|/o\.gif", RegexOptions.Compiled);

    private readonly HttpClient _http;

    public ImageResolver(HttpMessageHandler? handler = null)
    {
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var r = await ResolveWithReportAsync(pageXhtml, inlineImages, ct);
        return (r.Xhtml, r.Images);
    }

    public async Task<ImageResolution> ResolveWithReportAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var doc = await new HtmlParser().ParseDocumentAsync(pageXhtml, ct);
        var imgs = doc.QuerySelectorAll("img").ToList();
        var gate = new SemaphoreSlim(4);

        var work = imgs.Select(async (img, index) =>
        {
            var src = img.GetAttribute("src") ?? "";
            var alt = img.GetAttribute("alt") ?? "";
            var source = src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ? "cid"
                : src.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? "remote" : "other";

            // Junk by markup alone — no download needed.
            if (source == "remote" && TrackingUrl.IsMatch(src))
                return (img, (byte[]?)null, (string?)null, Junk(index, src, alt, source, "tracking-pixel URL pattern"));
            if (TryInt(img.GetAttribute("width"), out var aw) && TryInt(img.GetAttribute("height"), out var ah) && aw <= 2 && ah <= 2)
                return (img, null, null, Junk(index, src, alt, source, $"declared size {aw}x{ah}"));

            byte[]? data = null; string? contentType = null;
            if (source == "cid")
            {
                var cid = src[4..].Trim('<', '>');
                var match = inlineImages.FirstOrDefault(i => string.Equals(i.ContentId, cid, StringComparison.OrdinalIgnoreCase));
                if (match is not null) { data = match.Data; contentType = match.ContentType; }
            }
            else if (source == "remote")
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
                catch (Exception) when (!ct.IsCancellationRequested) { /* leave original src */ }
                finally { gate.Release(); }
            }

            if (data is null || contentType is null)
                return (img, null, null, new ImageDecision(index, src, null, 0, 0, 0, alt, source, "left-as-url", "not downloadable / no matching cid"));

            var (w, h) = Dimensions(data);
            if (w > 0 && w <= 2 && h <= 2)
                return (img, null, null, Junk(index, src, alt, source, $"decoded size {w}x{h}"));

            return (img, data, contentType, new ImageDecision(index, src, null, data.Length, w, h, alt, source, "embedded", "ok"));
        }).ToList();

        var results = await Task.WhenAll(work);
        var resolved = new List<ResolvedImage>();
        var decisions = new List<ImageDecision>();
        foreach (var (img, data, contentType, decision) in results)
        {
            if (decision.Decision == "dropped-junk") { img.Remove(); decisions.Add(decision); continue; }
            if (data is null || contentType is null) { decisions.Add(decision); continue; }
            var name = $"img{resolved.Count}";
            resolved.Add(new ResolvedImage(name, contentType, data, decision.Width, decision.Height));
            img.SetAttribute("src", $"name:{name}");
            decisions.Add(decision with { PartName = name });
        }

        return new ImageResolution(doc.ToHtml(XhtmlMarkupFormatter.Instance), resolved, decisions);
    }

    private static ImageDecision Junk(int index, string src, string alt, string source, string reason) =>
        new(index, src, null, 0, 0, 0, alt, source, "dropped-junk", reason);

    private static bool TryInt(string? s, out int v) => int.TryParse(s?.Trim().TrimEnd('p', 'x'), out v);

    private static (int W, int H) Dimensions(byte[] data)
    {
        try
        {
            using var ms = new MemoryStream(data);
            using var image = System.Drawing.Image.FromStream(ms, useEmbeddedColorManagement: false, validateImageData: false);
            return (image.Width, image.Height);
        }
        catch (Exception) { return (0, 0); }
    }
}
