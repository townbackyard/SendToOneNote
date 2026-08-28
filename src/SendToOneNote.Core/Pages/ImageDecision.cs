namespace SendToOneNote.Core.Pages;

public sealed record ImageDecision(int Index, string Src, string? PartName, int Bytes, int Width, int Height, string Alt,
    string Source, string Decision, string Reason);

public sealed record ImageResolution(string Xhtml, IReadOnlyList<ResolvedImage> Images, IReadOnlyList<ImageDecision> Decisions);
