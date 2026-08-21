namespace SendToOneNote.Core.Email;

public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);

public sealed record ParsedEmail(
    string Subject,
    string From,
    string To,
    string? Cc,
    DateTimeOffset? SentDate,
    string? HtmlBody,
    string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<string> AttachmentNames);

public sealed class EmlParseException(string message, Exception? inner = null)
    : Exception(message, inner);
