using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Backends;

public interface IOneNoteBackend
{
    /// <summary>"desktop" or "graph" — used for logging, tooltip, and per-backend recents.</summary>
    string Name { get; }
    Task<NotebookTree> GetTreeAsync(CancellationToken ct = default);
    /// <param name="pageXhtml">Full page XHTML from PageXhtmlBuilder + ImageResolver (src="name:imgN").</param>
    Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml, IReadOnlyList<ResolvedImage> images,
        CancellationToken ct = default);
}
