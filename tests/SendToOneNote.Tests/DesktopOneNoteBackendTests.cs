using System.Runtime.InteropServices;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class DesktopOneNoteBackendTests
{
    private const string Xhtml = "<html><head><title>Subject &amp; more</title></head><body><p>hi</p><img src=\"name:img0\"/></body></html>";

    [Fact]
    public async Task GetTreeParsesHierarchyFromCom()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var tree = await backend.GetTreeAsync();
        Assert.Equal("Alpha", Assert.Single(tree.Notebooks).Name);
        Assert.Equal("desktop", backend.Name);
    }

    [Fact]
    public async Task CreatePageWritesTitleHtmlBlockAndInlinedImages()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var images = new List<ResolvedImage> { new("img0", "image/png", [1, 2, 3]) };

        var page = await backend.CreatePageAsync("{S1}", Xhtml, images);

        Assert.Equal("{P1}{1}{B0}", page.Id);
        Assert.Equal(fake.Hyperlink, page.ClientUrl);
        Assert.Equal(("{S1}", OneNoteConstants.NpsDefault), Assert.Single(fake.CreatedPages));
        var xml = Assert.Single(fake.UpdatedPageXml);
        Assert.Contains("<![CDATA[Subject & more]]>", xml);
        Assert.Contains("one:HTMLBlock", xml);
        Assert.Contains("data:image/png;base64,AQID", xml);
        Assert.DoesNotContain("name:img0", xml);
    }

    [Fact]
    public async Task AllComCallsHappenOnTheWorkerThread()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var workerThread = await w.RunAsync(() => Environment.CurrentManagedThreadId);
        await backend.GetTreeAsync();
        Assert.Equal(workerThread, fake.ManagedThreadIdOfLastCall);
        await backend.CreatePageAsync("{S1}", Xhtml, []);
        Assert.Equal(workerThread, fake.ManagedThreadIdOfLastCall);
    }

    [Fact]
    public async Task ComExceptionBecomesDesktopOneNoteException()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("boom", unchecked((int)0x8004200B)) };
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var ex = await Assert.ThrowsAsync<DesktopOneNoteException>(() => backend.CreatePageAsync("{S1}", Xhtml, []));
        Assert.Equal(unchecked((int)0x8004200B), ex.HResultCode);
        Assert.Contains("read-only", ex.Message);
    }
}
