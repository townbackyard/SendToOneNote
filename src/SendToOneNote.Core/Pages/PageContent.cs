namespace SendToOneNote.Core.Pages;

/// <summary>Everything a backend needs to create one page: resolved XHTML (src="name:imgN",
/// data="name:attN"), the image parts, and the attachments to embed.</summary>
public sealed record PageContent(string Xhtml, IReadOnlyList<ResolvedImage> Images,
    IReadOnlyList<PlannedAttachment> Attachments);
