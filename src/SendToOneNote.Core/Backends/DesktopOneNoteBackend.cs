using System.Runtime.InteropServices;
using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Backends;

public sealed class DesktopOneNoteException(int hresult, string message) : Exception(message)
{
    public int HResultCode { get; } = hresult;
}

/// <summary>Writes pages into the local desktop OneNote via COM. Every COM call runs on the StaComWorker.</summary>
public sealed class DesktopOneNoteBackend : IOneNoteBackend
{
    private readonly StaComWorker _worker;
    private readonly Func<IApplication> _factory;
    private IApplication? _app; // created lazily ON THE WORKER THREAD

    public DesktopOneNoteBackend(StaComWorker worker, Func<IApplication>? applicationFactory = null)
    {
        _worker = worker;
        _factory = applicationFactory ?? CreateComApplication;
    }

    public string Name => "desktop";

    public Task<NotebookTree> GetTreeAsync(CancellationToken ct = default) =>
        _worker.RunAsync(() => Guard(() =>
        {
            App.GetHierarchy("", OneNoteConstants.HsSections, out var xml, OneNoteConstants.Xs2013);
            return HierarchyParser.Parse(xml);
        }));

    public Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml, IReadOnlyList<ResolvedImage> images,
        CancellationToken ct = default) =>
        _worker.RunAsync(() => Guard(() =>
        {
            var title = OneNotePageXmlBuilder.ExtractTitle(pageXhtml);
            var html = DataUriInliner.Inline(pageXhtml, images);
            App.CreateNewPage(sectionId, out var pageId, OneNoteConstants.NpsDefault);
            App.UpdatePageContent(OneNotePageXmlBuilder.Build(pageId, title, html),
                DateTime.MinValue, OneNoteConstants.Xs2013, false);
            App.GetHyperlinkToObject(pageId, "", out var link);
            return new CreatedPage(pageId, link, null);
        }));

    private IApplication App => _app ??= _factory();

    private static IApplication CreateComApplication()
    {
        var type = Type.GetTypeFromProgID(OneNoteConstants.ProgId)
            ?? throw new DesktopOneNoteException(0, "Desktop OneNote is not installed.");
        return (IApplication)Activator.CreateInstance(type)!;
    }

    private static T Guard<T>(Func<T> work)
    {
        try { return work(); }
        catch (COMException ex) { throw new DesktopOneNoteException(ex.HResult, Describe(ex)); }
    }

    // Best-known OneNote API HRESULTs (0x80042000 range); everything else is generic.
    private static string Describe(COMException ex) => (uint)ex.HResult switch
    {
        0x80042000 => "OneNote rejected the page XML as malformed.",
        0x80042001 => "OneNote rejected the page XML as invalid.",
        0x80042004 => "The chosen section no longer exists in OneNote.",
        0x80042005 => "The page could not be found after creation.",
        0x80042009 => "OneNote could not import the email's HTML.",
        0x8004200B => "The chosen section is read-only in OneNote.",
        0x8004200C => "The page is read-only in OneNote.",
        _ => $"Desktop OneNote error 0x{ex.HResult:X8}: {ex.Message}"
    };
}
