using System.Text.RegularExpressions;
using MimeKit;

namespace SendToOneNote.Core.Email;

public static class EmlParser
{
    public static ParsedEmail Parse(Stream emlStream)
    {
        MimeMessage msg;
        try
        {
            msg = MimeMessage.Load(emlStream);
        }
        catch (Exception ex)
        {
            throw new EmlParseException("Not a readable email file.", ex);
        }

        try
        {
            var html = msg.HtmlBody;
            var text = msg.TextBody;
            var from = msg.From.ToString();

            if (html is null && text is null && string.IsNullOrWhiteSpace(from))
                throw new EmlParseException("File has no email content (no body, no sender).");

            var inline = new List<InlineImage>();
            foreach (var part in (msg.BodyParts ?? []).OfType<MimePart>())
            {
                if (!string.Equals(part.ContentType.MediaType, "image", StringComparison.OrdinalIgnoreCase) ||
                    part.ContentId is null || part.Content is null)
                    continue;
                using var ms = new MemoryStream();
                part.Content.DecodeTo(ms);
                inline.Add(new InlineImage(
                    part.ContentId.Trim('<', '>'),
                    part.FileName ?? "image",
                    part.ContentType.MimeType,
                    ms.ToArray()));
            }

            // New Outlook's drag-out stamps a Content-ID on every part, so a Content-ID
            // alone does not make a part inline. Only an image the body references by
            // cid: is inline (captured above); everything else the sender marked as an
            // attachment is listed by name.
            var referencedCids = new HashSet<string>(
                Regex.Matches(html ?? "", @"cid:([^""'\s>)]+)", RegexOptions.IgnoreCase)
                    .Select(m => m.Groups[1].Value.Trim('<', '>')),
                StringComparer.OrdinalIgnoreCase);
            var attachments = new List<EmailAttachment>();
            var messageNames = new List<string>();
            foreach (var entity in msg.Attachments)
            {
                switch (entity)
                {
                    case MimePart p:
                        if (p.ContentId is not null && referencedCids.Contains(p.ContentId.Trim('<', '>')))
                            continue; // inline image referenced by the body, captured above
                        byte[] data = [];
                        if (p.Content is not null)
                        {
                            using var ms = new MemoryStream();
                            p.Content.DecodeTo(ms);
                            data = ms.ToArray();
                        }
                        // MimeKit defaults ContentType to text/plain (RFC 2045 §5.2) when the part has
                        // no Content-Type header at all, so a missing MIME type must be detected from
                        // the raw headers, not from ContentType.MimeType (which is never null/empty).
                        var mime = p.Headers.Contains(HeaderId.ContentType) ? p.ContentType?.MimeType : null;
                        attachments.Add(new EmailAttachment(
                            CleanName(p.FileName, "attachment"),
                            string.IsNullOrWhiteSpace(mime) ? "application/octet-stream" : mime,
                            data));
                        break;
                    case MessagePart mp:
                        // Nested .eml: listed by name, never embedded (no useful handler for new-Outlook users).
                        var subject = mp.Message?.Subject;
                        var rawName = !string.IsNullOrWhiteSpace(subject) ? subject : mp.ContentDisposition?.FileName;
                        messageNames.Add(CleanName(rawName, "attached message"));
                        break;
                }
            }

            DateTimeOffset? sent = msg.Date == DateTimeOffset.MinValue ? null : msg.Date;

            return new ParsedEmail(
                Subject: string.IsNullOrWhiteSpace(msg.Subject) ? "(no subject)" : msg.Subject,
                From: from,
                To: msg.To.ToString(),
                Cc: msg.Cc.Count > 0 ? msg.Cc.ToString() : null,
                SentDate: sent,
                HtmlBody: html,
                TextBody: text,
                InlineImages: inline,
                Attachments: attachments,
                AttachedMessageNames: messageNames);
        }
        catch (EmlParseException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new EmlParseException("Failed to extract email content.", ex);
        }
    }

    // A C0/C1 control character in a file name survives HtmlEncode and AngleSharp into the page
    // XHTML (Graph rejects ill-formed XHTML) and survives into the OneNote page XML on the desktop
    // path (OneNote rejects the page XML), so one odd name would fail the whole save on both backends.
    private static string CleanName(string? raw, string fallback)
    {
        var cleaned = raw is null ? "" : new string(raw.Where(c => !char.IsControl(c)).ToArray());
        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }
}
