using System.Net;

namespace SendToOneNote.Tests;

public sealed class StubHttpHandler : HttpMessageHandler
{
    public List<HttpRequestMessage> Requests { get; } = [];
    private readonly Func<HttpRequestMessage, HttpResponseMessage> _responder;

    public StubHttpHandler(Func<HttpRequestMessage, HttpResponseMessage> responder)
        => _responder = responder;

    public static HttpResponseMessage Png(byte[] bytes) =>
        new(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(bytes)
            {
                Headers = { ContentType = new("image/png") }
            }
        };

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken ct)
    {
        Requests.Add(request);
        return Task.FromResult(_responder(request));
    }
}
