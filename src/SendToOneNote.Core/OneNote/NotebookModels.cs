namespace SendToOneNote.Core.OneNote;

public sealed record SectionNode(string Id, string Name);

public sealed record GroupNode(string Id, string Name,
    IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);

public sealed record NotebookNode(string Id, string Name,
    IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);

public sealed record NotebookTree(IReadOnlyList<NotebookNode> Notebooks, DateTimeOffset FetchedUtc);
