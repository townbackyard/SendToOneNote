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

            var attachments = msg.Attachments.OfType<MimePart>()
                .Where(p => p.ContentId is null) // inline images already captured
                .Select(p => p.FileName ?? "attachment")
                .ToList();

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
                AttachmentNames: attachments);
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
}
