namespace SendToOneNote.Core.Email;

public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);

/// <summary>A file the sender attached (not an inline cid: image). Bytes are decoded.</summary>
public sealed record EmailAttachment(string FileName, string ContentType, byte[] Data);

public sealed record ParsedEmail(
    string Subject,
    string From,
    string To,
    string? Cc,
    DateTimeOffset? SentDate,
    string? HtmlBody,
    string? TextBody,
    IReadOnlyList<InlineImage> InlineImages,
    IReadOnlyList<EmailAttachment> Attachments,
    IReadOnlyList<string> AttachedMessageNames)
{
    /// <summary>Header-row list: file attachments first, then attached messages (never embedded).</summary>
    public IReadOnlyList<string> AttachmentNames =>
        [.. Attachments.Select(a => a.FileName), .. AttachedMessageNames];
}

public sealed class EmlParseException(string message, Exception? inner = null)
    : Exception(message, inner);
