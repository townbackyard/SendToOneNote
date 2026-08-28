using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Backends;

public sealed class GraphBackend(OneNoteClient client) : IOneNoteBackend
{
    public string Name => "graph";

    public Task<NotebookTree> GetTreeAsync(CancellationToken ct = default) => client.GetNotebookTreeAsync(ct);

    public Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml, IReadOnlyList<ResolvedImage> images,
        CancellationToken ct = default) =>
        client.CreatePageAsync(sectionId, PagePlanner.Plan(pageXhtml, images), ct);
}
