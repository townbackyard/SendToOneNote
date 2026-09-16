using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

/// <summary>An attachment that will be embedded; PartName is "att1", "att2", … (Graph part name / desktop file index).</summary>
public sealed record PlannedAttachment(string PartName, EmailAttachment Attachment);

public sealed record OmittedAttachment(string FileName, long Bytes, string Reason);

public sealed record AttachmentPlan(IReadOnlyList<PlannedAttachment> Embedded, IReadOnlyList<OmittedAttachment> Omitted)
{
    public static AttachmentPlan Empty { get; } = new([], []);
}

/// <summary>Applies the IncludeAttachments switch and the per-file size cap. Pure; backend-agnostic.</summary>
public static class AttachmentPlanner
{
    public const string OverSizeLimit = "over the size limit";

    public static AttachmentPlan Plan(ParsedEmail email, bool includeAttachments, long maxAttachmentBytes)
    {
        if (!includeAttachments) return AttachmentPlan.Empty;

        var embedded = new List<PlannedAttachment>();
        var omitted = new List<OmittedAttachment>();
        foreach (var a in email.Attachments)
        {
            if (a.Data.LongLength > maxAttachmentBytes)
                omitted.Add(new OmittedAttachment(a.FileName, a.Data.LongLength, OverSizeLimit));
            else
                embedded.Add(new PlannedAttachment($"att{embedded.Count + 1}", a));
        }
        return new AttachmentPlan(embedded, omitted);
    }
}
