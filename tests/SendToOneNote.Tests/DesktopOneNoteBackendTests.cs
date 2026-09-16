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

        var page = await backend.CreatePageAsync("{S1}", new PageContent(Xhtml, images, []));

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
        await backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], []));
        Assert.Equal(workerThread, fake.ManagedThreadIdOfLastCall);
    }

    [Fact]
    public async Task ComExceptionBecomesDesktopOneNoteException()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("boom", unchecked((int)0x8004200B)) };
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var ex = await Assert.ThrowsAsync<DesktopOneNoteException>(() => backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], [])));
        Assert.Equal(unchecked((int)0x8004200B), ex.HResultCode);
        Assert.Contains("read-only", ex.Message);
    }

    [Theory]
    [InlineData(unchecked((int)0x80010108))] // RPC_E_DISCONNECTED
    [InlineData(unchecked((int)0x800706BA))] // HRESULT_FROM_WIN32(RPC_S_SERVER_UNAVAILABLE)
    public async Task ReactivatesAfterRpcDisconnection(int hresult)
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("gone", hresult) };
        var factoryCalls = 0;
        var backend = new DesktopOneNoteBackend(w, () => { factoryCalls++; return fake; });

        await Assert.ThrowsAsync<DesktopOneNoteException>(() => backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], [])));
        Assert.Equal(1, factoryCalls);

        fake.ThrowOnUpdate = null;
        var page = await backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], []));

        Assert.Equal(2, factoryCalls);
        Assert.Equal("{P1}{1}{B0}", page.Id);
    }

    [Fact]
    public async Task DoesNotReactivateOnNonConnectionComException()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("boom", unchecked((int)0x8004200B)) };
        var factoryCalls = 0;
        var backend = new DesktopOneNoteBackend(w, () => { factoryCalls++; return fake; });

        await Assert.ThrowsAsync<DesktopOneNoteException>(() => backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], [])));
        await Assert.ThrowsAsync<DesktopOneNoteException>(() => backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], [])));

        Assert.Equal(1, factoryCalls);
    }

    [Fact]
    public async Task AlreadyCancelledTokenSkipsComCalls()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(
            () => backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], []), cts.Token));
        Assert.Empty(fake.CreatedPages);
    }

    private static PlannedAttachment Att(string name, byte[] data) =>
        new("att1", new SendToOneNote.Core.Email.EmailAttachment(name, "application/pdf", data));

    private static string XhtmlWithAttachment(PlannedAttachment a) =>
        "<html><head><title>t</title></head><body>" + AttachmentMarkup.ObjectElement(a) + "<p>hi</p></body></html>";

    [Fact]
    public async Task AttachmentIsWrittenForTheComCallThenCleanedUpAndObjectMarkupStripped()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var a = Att("invoice.pdf", [0x25, 0x50, 0x44, 0x46]);

        await backend.CreatePageAsync("{S1}", new PageContent(XhtmlWithAttachment(a), [], [a]));

        var (path, bytes) = Assert.Single(fake.FilesSeenDuringUpdate);
        Assert.EndsWith("invoice.pdf", path);
        Assert.Equal(a.Attachment.Data, bytes);           // existed, with the right bytes, during the call
        Assert.False(File.Exists(path));                   // gone afterwards
        Assert.False(Directory.Exists(Path.GetDirectoryName(path)));
        var xml = Assert.Single(fake.UpdatedPageXml);
        Assert.Contains("one:InsertedFile", xml);
        Assert.Contains("preferredName=\"invoice.pdf\"", xml);
        Assert.DoesNotContain("<object", xml);
    }

    [Fact]
    public async Task TempFolderIsRemovedWhenTheComCallThrows()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication { ThrowOnUpdate = new COMException("boom", unchecked((int)0x8004200B)) };
        var backend = new DesktopOneNoteBackend(w, () => fake);
        var a = Att("x.pdf", [1]);
        var before = Directory.Exists(AttachmentTempFolder.DefaultRoot)
            ? Directory.GetDirectories(AttachmentTempFolder.DefaultRoot).ToHashSet() : [];

        await Assert.ThrowsAsync<DesktopOneNoteException>(() =>
            backend.CreatePageAsync("{S1}", new PageContent(XhtmlWithAttachment(a), [], [a])));

        var after = Directory.Exists(AttachmentTempFolder.DefaultRoot)
            ? Directory.GetDirectories(AttachmentTempFolder.DefaultRoot).ToHashSet() : [];
        Assert.Subset(before, after); // nothing new left behind
    }

    [Fact]
    public async Task NoAttachmentsProducesSingleOutlineXml()
    {
        using var w = new StaComWorker();
        var fake = new FakeOneNoteApplication();
        var backend = new DesktopOneNoteBackend(w, () => fake);
        await backend.CreatePageAsync("{S1}", new PageContent(Xhtml, [], []));
        Assert.DoesNotContain("one:InsertedFile", Assert.Single(fake.UpdatedPageXml));
        Assert.Empty(fake.FilesSeenDuringUpdate);
    }
}
