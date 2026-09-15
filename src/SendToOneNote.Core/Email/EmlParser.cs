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
            var attachments = msg.Attachments.OfType<MimePart>()
                .Where(p => !(p.ContentId is not null && referencedCids.Contains(p.ContentId.Trim('<', '>'))))
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
