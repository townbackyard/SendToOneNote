using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Backends;

public interface IOneNoteBackend
{
    /// <summary>"desktop" or "graph" — used for logging, tooltip, and per-backend recents.</summary>
    string Name { get; }
    Task<NotebookTree> GetTreeAsync(CancellationToken ct = default);
    /// <param name="content">Page XHTML from PageXhtmlBuilder + ImageResolver, plus image and attachment parts.</param>
    Task<CreatedPage> CreatePageAsync(string sectionId, PageContent content, CancellationToken ct = default);
}
