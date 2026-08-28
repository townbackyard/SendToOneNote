# Desktop OneNote Backend (v1.1) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Save emails into the user's desktop OneNote app via COM (instant, local, no sign-in) when it is installed, keeping the Microsoft Graph path as an automatic fallback; strip tracking-pixel/spacer images on both paths; rank and cap images on the Graph path; add an image-diagnostics dump.

**Architecture:** A new `IOneNoteBackend` seam at the end of the existing pipeline (watcher → EmlParser → PageXhtmlBuilder → ImageResolver → *backend*). `DesktopOneNoteBackend` talks to `OneNote.Application` through a hand-declared early-bound COM interface on a dedicated STA worker thread and writes pages as `one:HTMLBlock`s with base64-inlined images. `GraphBackend` wraps the existing `PagePlanner` + `OneNoteClient`. `BackendSelector` picks per the `Backend` setting (`auto` | `desktop` | `graph`).

**Tech Stack:** .NET 10 / C# / WPF (`net10.0-windows`), COM interop via `[ComImport]` (no new packages), AngleSharp (existing), System.Drawing.Common (existing), xUnit + Xunit.SkippableFact (existing).

**Spec:** `docs/superpowers/specs/2026-08-28-desktop-onenote-backend-design.md` (binding). v1 spec for the unchanged parts: `docs/superpowers/specs/2026-08-21-sendtoonenote-design.md`.

## Global Constraints

- .NET 10, TFM `net10.0-windows` (pinned in `Directory.Build.props`); `SendToOneNote.Core` contains no UI types.
- **No new NuGet packages.** COM is reached through a hand-declared `[ComImport]` interface; there is no interop package for OneNote on nuget.org.
- **All COM calls run on the single `StaComWorker` thread.** The RCW is created there and never touched from another thread.
- `Backend` setting: `"auto"` (default; try desktop first) | `"desktop"` (force COM) | `"graph"` (force cloud even when desktop OneNote is installed — owner's escape hatch).
- Desktop path: no `MsalTokenProvider`, no sign-in UI, no tree cache; toast-only after save (no `NavigateTo`).
- Images: tracking pixels/spacers dropped on both paths; desktop embeds every remaining image as a base64 data URI; Graph ranks by area and **drops** images beyond the existing caps (`PagePlanner.MaxBinaryPartsPerRequest = 30`, `MaxRequestBytes = 3_500_000`); Graph never produces append batches any more.
- Diagnostics dump: `<DropFolder>\Diagnostics\<eml-stem>-<yyyyMMdd-HHmmss>\images.csv` (+ image files), only when `ImageDiagnostics = true` (default false); never contains email text.
- Recents are per backend: `RecentSectionIds` (Graph, existing) and `RecentDesktopSectionIds` (new).
- Test fixtures are synthetic. The owner's real emails in `fixtures/local/` are gitignored — never commit, quote, or copy them. The synthetic hierarchy XML uses invented notebook names.
- Tests touching real OneNote or Graph are gated on `STN_INTEGRATION=1` and must SKIP (not fail) elsewhere.
- Commit style: `feat:` / `fix:` / `test:` / `docs:` prefixes, present tense; end every commit with the session's standard `Co-Authored-By` trailer.
- Owner-machine policy (`CLAUDE.md`): installing/running the app, Outlook drags, and deleting the spike folder are the owner's manual steps — propose, don't execute.
- Work on `main` (established for this repo); push after each task; CI (`.github/workflows/build.yml`) must stay green.

## Context for a fresh session (read before Task 1)

**Repo state at plan time:** `main` @ `8129c54`, CI green, 54 unit tests + 2 gated integration tests. Layout: `src/SendToOneNote.Core` (logic), `src/SendToOneNote` (WPF glue), `tests/SendToOneNote.Tests`, `fixtures/synthetic`, `docs/`, `agents/`. Read `AGENTS.md` and `agents/knowledge/architecture.md` first.

**Existing signatures this plan builds on (exact):**

```csharp
// SendToOneNote.Core.Email
public sealed record InlineImage(string ContentId, string FileName, string ContentType, byte[] Data);
public sealed record ParsedEmail(string Subject, string From, string To, string? Cc, DateTimeOffset? SentDate,
    string? HtmlBody, string? TextBody, IReadOnlyList<InlineImage> InlineImages, IReadOnlyList<string> AttachmentNames);
public static class EmlParser { public static ParsedEmail Parse(Stream emlStream); }

// SendToOneNote.Core.Pages
public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data);
public sealed class ImageResolver(HttpMessageHandler? handler = null)
{   public Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default); }
public static class PageXhtmlBuilder { public static string Build(ParsedEmail email); }  // <html><head><title>subject</title></head><body>…
public sealed record OneNoteRequestPart(string Name, string ContentType, byte[] Data);
public sealed record AppendPlan(string CommandsJson, IReadOnlyList<OneNoteRequestPart> Parts);
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts, IReadOnlyList<AppendPlan> Appends);
public static class PagePlanner { public const int MaxRequestBytes = 3_500_000; public const int MaxBinaryPartsPerRequest = 30;
    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images); }
public static class ImageShrinker { public static (byte[] Data, string ContentType) ShrinkIfNeeded(byte[] data, string contentType, int maxBytes); }

// SendToOneNote.Core.OneNote
public sealed record SectionNode(string Id, string Name);
public sealed record GroupNode(string Id, string Name, IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);
public sealed record NotebookNode(string Id, string Name, IReadOnlyList<SectionNode> Sections, IReadOnlyList<GroupNode> Groups);
public sealed record NotebookTree(IReadOnlyList<NotebookNode> Notebooks, DateTimeOffset FetchedUtc);
public sealed record CreatedPage(string Id, string? ClientUrl, string? WebUrl);
public sealed class OneNoteApiException(int statusCode, string message) : Exception(message) { public int StatusCode { get; } }
public sealed class OneNoteClient(ITokenProvider tokens, HttpMessageHandler? handler = null, TimeSpan? appendRetryBaseDelay = null)
{   public Task<NotebookTree> GetNotebookTreeAsync(CancellationToken ct = default);
    public Task<CreatedPage> CreatePageAsync(string sectionId, PagePlan plan, CancellationToken ct = default); }

// SendToOneNote.Core.Auth
public interface ITokenProvider { Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default); string? SignedInUser { get; } }
public sealed class MsalTokenProvider(string cacheDir, string? clientIdOverride = null, IntPtr parentWindow = default) : ITokenProvider;
public sealed class AuthRequiredException(string message) : Exception(message);

// SendToOneNote.Core.Storage
public sealed class AppSettings { string? DropFolder; string? ClientIdOverride; bool DeleteOnSuccess = true; List<string> RecentSectionIds = []; }
public sealed class JsonFileStore(string? rootDir = null) { string RootDir; AppSettings LoadSettings(); void SaveSettings(AppSettings);
    NotebookTree? LoadTreeCache(); void SaveTreeCache(NotebookTree); }

// SendToOneNote.Core.Picker
public sealed record PickerItem(string SectionId, string SectionName, string Path);
public sealed class SectionPickerViewModel(NotebookTree tree, IReadOnlyList<string> recentSectionIds)
{   IReadOnlyList<PickerItem> AllSections; IReadOnlyList<PickerItem> Filter(string query);
    static List<string> PushRecent(List<string> recents, string sectionId, int cap = 10); }

// SendToOneNote.Core.Watch
public sealed class DropFolderWatcher(string folder, TimeSpan? readinessTimeout = null) : IDisposable
{ event Action<string>? EmlReady; event Action<string>? NonEmlIgnored; event Action<string>? WatchError; void Start(); }

// src/SendToOneNote (WPF)
public sealed class SavePipeline(JsonFileStore store, OneNoteClient client, FileLog log)
{ event Action<string, string?>? Saved; event Action<string>? Failed; Task HandleEmlAsync(string path); }  // serialized by a SemaphoreSlim; File.Exists guards; Failed\ + .error.txt sidecar on failure
public sealed class TrayContext : IDisposable { void Run(); static void TryLogStatic(string, Exception?); }  // builds MsalTokenProvider, OneNoteClient, SavePipeline, DropFolderWatcher, TaskbarIcon (H.NotifyIcon, ForceCreate), Notify(title,msg,url,icon)
public partial class FirstRunWindow(ITokenProvider tokens, AppSettings settings) : Window { bool Completed; }  // sign-in button, folder picker, startup-shortcut checkbox
public sealed class FileLog(string dir) { void Info(string); void Error(string, Exception? = null); }
```

**Test conventions:** xUnit; `Fixtures.Open(name)` reads `fixtures/synthetic/`; `StubHttpHandler(Func<HttpRequestMessage,HttpResponseMessage>)` with `.Requests`; `StubHttpHandler.Png(bytes)`; gated tests use `[SkippableFact]` + `Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1", …)`. Run `dotnet test --filter <ClassName>` while iterating, full `dotnet test` before each commit. NuGet on the owner's machine has an unrelated private feed that 401s — irrelevant here (no packages added).

**Verified facts (spike, 2026-08-28, owner's machine, desktop OneNote M365 Click-to-Run):**
- `Type.GetTypeFromProgID("OneNote.Application")` resolves; `Activator.CreateInstance` ≈ 5 ms; casting the RCW to the interface below succeeds.
- `dynamic` and `Type.InvokeMember` both FAIL (`TYPE_E_LIBNOTREGISTERED`, Click-to-Run virtualized type library). Only vtable calls through the declared interface work.
- `GetHierarchy("", 3, out xml, 2)` ≈ 100 ms for 17 notebooks; `CreateNewPage` ≈ 2 ms; `UpdatePageContent` with an `HTMLBlock` ≈ 23 ms; `GetHyperlinkToObject` returns an `onenote:` URL.
- `HTMLBlock` renders tables/bold/italic/links/bullets; remote `<img src="https://…">` is downloaded at import and stored embedded (13,504 bytes for a logo); base64 data-URI `<img>` works; page synced to SharePoint within a minute.
- The spike lives at `D:\Projects\TownBackyard\SendToOneNote-spike-com` (outside the repo). Its `IApplication` declaration is reproduced verbatim in Task 1.

**Process:** same as v1 — subagent-driven execution with a per-task review; ledger under `.superpowers/sdd/<plan-basename>/`; one GitHub issue per task (mapping at the end of this plan); close each issue with its commit SHA when the task's review is clean.

---

### Task 1: COM interop declaration, STA worker, availability probe

**Files:**
- Create: `src/SendToOneNote.Core/Desktop/OneNoteInterop.cs`, `src/SendToOneNote.Core/Desktop/StaComWorker.cs`, `src/SendToOneNote.Core/Desktop/DesktopOneNoteProbe.cs`
- Test: `tests/SendToOneNote.Tests/StaComWorkerTests.cs`, `tests/SendToOneNote.Tests/DesktopOneNoteProbeTests.cs`

**Interfaces:**
- Consumes: nothing.
- Produces:

```csharp
namespace SendToOneNote.Core.Desktop;
public static class OneNoteConstants { const string ProgId = "OneNote.Application"; const int HsSections = 3; const int HsPages = 4; const int Xs2013 = 2; const int NpsDefault = 0; const int PiBinaryData = 1; }
[ComImport, Guid("452AC71A-B655-4967-A208-A4CC39DD7949"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IApplication { /* 16 methods in IDL order — see Step 1 */ }
public sealed class StaComWorker(string threadName = "SendToOneNote COM") : IDisposable
{ Task<T> RunAsync<T>(Func<T> work); Task RunAsync(Action work); }
public static class DesktopOneNoteProbe { public static bool IsAvailable(string progId = OneNoteConstants.ProgId); } // MUST be called on the STA worker
```

- [ ] **Step 1: Write the interop declaration (verbatim from the verified spike)**

`src/SendToOneNote.Core/Desktop/OneNoteInterop.cs`:

```csharp
using System.Runtime.InteropServices;

namespace SendToOneNote.Core.Desktop;

/// <summary>Enum values from the OneNote type library, as ints (we never load the typelib).</summary>
public static class OneNoteConstants
{
    public const string ProgId = "OneNote.Application";
    public const int HsSections = 3;    // HierarchyScope.hsSections
    public const int HsPages = 4;       // HierarchyScope.hsPages
    public const int Xs2013 = 2;        // XMLSchema.xs2013
    public const int NpsDefault = 0;    // NewPageStyle.npsDefault
    public const int PiBinaryData = 1;  // PageInfo.piBinaryData
    public static readonly string Namespace2013 = "http://schemas.microsoft.com/office/onenote/2013/onenote";
}

/// <summary>
/// OneNote 2013+ IApplication, hand-declared so calls go through the COM vtable.
/// Late binding (dynamic / InvokeMember) fails on Click-to-Run Office because the
/// type library lives in a virtualized registry hive (TYPE_E_LIBNOTREGISTERED).
/// Methods MUST stay in IDL order; only the prefix through GetHyperlinkToObject is declared.
/// Verified against the live app 2026-08-28.
/// </summary>
[ComImport, Guid("452AC71A-B655-4967-A208-A4CC39DD7949"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IApplication
{
    void GetHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID, int hsScope,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut, int xsSchema);
    void UpdateHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrChangesXmlIn, int xsSchema);
    void OpenHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrPath, [MarshalAs(UnmanagedType.BStr)] string bstrRelativeToObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrObjectID, int cftIfNotExist);
    void DeleteHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrObjectID, DateTime dateExpectedLastModified, bool deletePermanently);
    void CreateNewPage([MarshalAs(UnmanagedType.BStr)] string bstrSectionID, [MarshalAs(UnmanagedType.BStr)] out string pbstrPageID, int npsNewPageStyle);
    void CloseNotebook([MarshalAs(UnmanagedType.BStr)] string bstrNotebookID, bool force);
    void GetHierarchyParent([MarshalAs(UnmanagedType.BStr)] string bstrObjectID, [MarshalAs(UnmanagedType.BStr)] out string pbstrParentID);
    void GetPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] out string pbstrPageXMLOut,
        int pageInfoToExport, int xsSchema);
    void UpdatePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageChangesXmlIn, DateTime dateExpectedLastModified, int xsSchema, bool force);
    void GetBinaryPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] string bstrCallbackID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrBinaryObjectB64Out);
    void DeletePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        DateTime dateExpectedLastModified, bool force);
    void NavigateTo([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyObjectID, [MarshalAs(UnmanagedType.BStr)] string bstrObjectID, bool fNewWindow);
    void NavigateToUrl([MarshalAs(UnmanagedType.BStr)] string bstrUrl, bool fNewWindow);
    void Publish([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID, [MarshalAs(UnmanagedType.BStr)] string bstrTargetFilePath,
        int pfPublishFormat, [MarshalAs(UnmanagedType.BStr)] string bstrCLSIDofExporter);
    void OpenPackage([MarshalAs(UnmanagedType.BStr)] string bstrPathPackage, [MarshalAs(UnmanagedType.BStr)] string bstrPathDest,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrPathOut);
    void GetHyperlinkToObject([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID, [MarshalAs(UnmanagedType.BStr)] string bstrPageContentObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHyperlinkOut);
}
```

- [ ] **Step 2: Write failing tests for the worker and probe**

`tests/SendToOneNote.Tests/StaComWorkerTests.cs`:

```csharp
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class StaComWorkerTests
{
    [Fact]
    public async Task RunsWorkOnAnStaThread()
    {
        using var w = new StaComWorker();
        var state = await w.RunAsync(() => Thread.CurrentThread.GetApartmentState());
        Assert.Equal(ApartmentState.STA, state);
    }

    [Fact]
    public async Task AllWorkRunsOnTheSameThread()
    {
        using var w = new StaComWorker();
        var a = await w.RunAsync(() => Environment.CurrentManagedThreadId);
        var b = await w.RunAsync(() => Environment.CurrentManagedThreadId);
        Assert.Equal(a, b);
        Assert.NotEqual(Environment.CurrentManagedThreadId, a);
    }

    [Fact]
    public async Task ExceptionsPropagateToTheCaller()
    {
        using var w = new StaComWorker();
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            w.RunAsync<int>(() => throw new InvalidOperationException("boom")));
        Assert.Equal(7, await w.RunAsync(() => 7)); // worker survives a failed job
    }
}
```

`tests/SendToOneNote.Tests/DesktopOneNoteProbeTests.cs`:

```csharp
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class DesktopOneNoteProbeTests
{
    [Fact]
    public async Task UnknownProgIdIsNotAvailable()
    {
        using var w = new StaComWorker();
        Assert.False(await w.RunAsync(() => DesktopOneNoteProbe.IsAvailable("SendToOneNote.NoSuchProgId.Test")));
    }

    [Fact]
    public async Task RealProbeNeverThrows()
    {
        // CI runners have no OneNote (false); the owner's machine has it (true). Either is fine.
        using var w = new StaComWorker();
        var ex = await Record.ExceptionAsync(() => w.RunAsync(() => DesktopOneNoteProbe.IsAvailable()));
        Assert.Null(ex);
    }
}
```

Run: `dotnet test --filter "StaComWorkerTests|DesktopOneNoteProbeTests"`
Expected: FAIL (types missing).

- [ ] **Step 3: Implement the worker and probe**

`src/SendToOneNote.Core/Desktop/StaComWorker.cs`:

```csharp
using System.Collections.Concurrent;

namespace SendToOneNote.Core.Desktop;

/// <summary>
/// One dedicated STA thread that owns every COM call. The RCW must be created on
/// this thread and never touched from another; RunAsync is the only entry point.
/// </summary>
public sealed class StaComWorker : IDisposable
{
    private readonly BlockingCollection<Action> _queue = new();
    private readonly Thread _thread;

    public StaComWorker(string threadName = "SendToOneNote COM")
    {
        _thread = new Thread(() =>
        {
            foreach (var job in _queue.GetConsumingEnumerable()) job();
        })
        { IsBackground = true, Name = threadName };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> RunAsync<T>(Func<T> work)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try { tcs.SetResult(work()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task;
    }

    public Task RunAsync(Action work) => RunAsync(() => { work(); return true; });

    public void Dispose() => _queue.CompleteAdding();
}
```

`src/SendToOneNote.Core/Desktop/DesktopOneNoteProbe.cs`:

```csharp
using System.Runtime.InteropServices;

namespace SendToOneNote.Core.Desktop;

public static class DesktopOneNoteProbe
{
    /// <summary>
    /// True when desktop OneNote is installed AND exposes the 2013+ IApplication.
    /// Activates OneNote (launches it if not running). Call on the StaComWorker.
    /// </summary>
    public static bool IsAvailable(string progId = OneNoteConstants.ProgId)
    {
        var type = Type.GetTypeFromProgID(progId);
        if (type is null) return false;
        object? instance = null;
        try
        {
            instance = Activator.CreateInstance(type);
            return instance is IApplication; // QueryInterface for the declared IID
        }
        catch (Exception)
        {
            return false;
        }
        finally
        {
            if (instance is not null && Marshal.IsComObject(instance)) Marshal.ReleaseComObject(instance);
        }
    }
}
```

- [ ] **Step 4: Run tests**

Run: `dotnet test --filter "StaComWorkerTests|DesktopOneNoteProbeTests"` then full `dotnet test`
Expected: PASS (5 new); suite otherwise unchanged (54 passed, 2 skipped before this task).

- [ ] **Step 5: Commit**

```powershell
git add src tests
git commit -m "feat: OneNote COM interop declaration, STA worker, availability probe"
git push
```

---

### Task 2: HierarchyParser (OneNote hierarchy XML → NotebookTree)

**Files:**
- Create: `src/SendToOneNote.Core/Desktop/HierarchyParser.cs`
- Test: `tests/SendToOneNote.Tests/HierarchyParserTests.cs`

**Interfaces:**
- Consumes: `NotebookTree`/`NotebookNode`/`GroupNode`/`SectionNode` (existing), `OneNoteConstants.Namespace2013` (Task 1).
- Produces: `public static class HierarchyParser { public static NotebookTree Parse(string hierarchyXml); }`
- Rules: `one:Notebook` → `NotebookNode(ID, name)`; child `one:Section` → `SectionNode(ID, name)`; child `one:SectionGroup` → `GroupNode(ID, name, sections, nested groups)` recursively; skip any element with `isRecycleBin="true"` or `isInRecycleBin="true"`; skip sections with `isDeletedPages="true"`; a notebook with no remaining sections/groups is still listed (empty notebook is a valid pick target for a future "new section" feature — keep it).

- [ ] **Step 1: Write failing tests with a synthetic hierarchy**

`tests/SendToOneNote.Tests/HierarchyParserTests.cs`:

```csharp
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class HierarchyParserTests
{
    // Invented notebooks — never the owner's real hierarchy output.
    private const string Xml = """
    <?xml version="1.0"?>
    <one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
      <one:Notebook name="Alpha" ID="{N1}{1}{B0}" path="https://example.test/Alpha/">
        <one:Section name="Inbox" ID="{S1}{1}{B0}" path="https://example.test/Alpha/Inbox.one"/>
        <one:SectionGroup name="Taxes" ID="{G1}{1}{B0}">
          <one:Section name="Taxes 2026" ID="{S2}{1}{B0}"/>
          <one:SectionGroup name="Archive" ID="{G2}{1}{B0}">
            <one:Section name="Taxes 2025" ID="{S3}{1}{B0}"/>
          </one:SectionGroup>
        </one:SectionGroup>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{G9}{1}{B0}" isRecycleBin="true">
          <one:Section name="Deleted Pages" ID="{S9}{1}{B0}" isInRecycleBin="true" isDeletedPages="true"/>
        </one:SectionGroup>
      </one:Notebook>
      <one:Notebook name="Beta" ID="{N2}{1}{B0}">
        <one:Section name="Notes" ID="{S4}{1}{B0}"/>
      </one:Notebook>
    </one:Notebooks>
    """;

    [Fact]
    public void ParsesNotebooksSectionsAndNestedGroups()
    {
        var tree = HierarchyParser.Parse(Xml);
        Assert.Equal(2, tree.Notebooks.Count);
        var alpha = tree.Notebooks[0];
        Assert.Equal("Alpha", alpha.Name);
        Assert.Equal("{N1}{1}{B0}", alpha.Id);
        Assert.Equal("Inbox", Assert.Single(alpha.Sections).Name);
        var taxes = Assert.Single(alpha.Groups);
        Assert.Equal("Taxes 2026", Assert.Single(taxes.Sections).Name);
        Assert.Equal("Taxes 2025", Assert.Single(Assert.Single(taxes.Groups).Sections).Name);
        Assert.Equal("Notes", Assert.Single(tree.Notebooks[1].Sections).Name);
    }

    [Fact]
    public void SkipsRecycleBinGroupsAndSections()
    {
        var tree = HierarchyParser.Parse(Xml);
        var allGroupNames = tree.Notebooks.SelectMany(n => n.Groups).Select(g => g.Name);
        Assert.DoesNotContain("OneNote_RecycleBin", allGroupNames);
        var allSectionNames = tree.Notebooks.SelectMany(n => n.Sections).Select(s => s.Name);
        Assert.DoesNotContain("Deleted Pages", allSectionNames);
    }

    [Fact]
    public void FetchedUtcIsNow()
    {
        var tree = HierarchyParser.Parse(Xml);
        Assert.True((DateTimeOffset.UtcNow - tree.FetchedUtc) < TimeSpan.FromMinutes(1));
    }
}
```

Run: `dotnet test --filter HierarchyParserTests`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Desktop/HierarchyParser.cs`:

```csharp
using System.Xml.Linq;
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Core.Desktop;

public static class HierarchyParser
{
    private static readonly XNamespace One = OneNoteConstants.Namespace2013;

    public static NotebookTree Parse(string hierarchyXml)
    {
        var doc = XDocument.Parse(hierarchyXml);
        var notebooks = doc.Descendants(One + "Notebook")
            .Where(IsLive)
            .Select(nb => new NotebookNode(Id(nb), Name(nb), Sections(nb), Groups(nb)))
            .ToList();
        return new NotebookTree(notebooks, DateTimeOffset.UtcNow);
    }

    private static IReadOnlyList<SectionNode> Sections(XElement parent) =>
        parent.Elements(One + "Section")
            .Where(s => IsLive(s) && !Flag(s, "isDeletedPages"))
            .Select(s => new SectionNode(Id(s), Name(s)))
            .ToList();

    private static IReadOnlyList<GroupNode> Groups(XElement parent) =>
        parent.Elements(One + "SectionGroup")
            .Where(IsLive)
            .Select(g => new GroupNode(Id(g), Name(g), Sections(g), Groups(g)))
            .ToList();

    private static bool IsLive(XElement e) => !Flag(e, "isRecycleBin") && !Flag(e, "isInRecycleBin");
    private static bool Flag(XElement e, string attr) =>
        string.Equals(e.Attribute(attr)?.Value, "true", StringComparison.OrdinalIgnoreCase);
    private static string Id(XElement e) => e.Attribute("ID")?.Value ?? "";
    private static string Name(XElement e) => e.Attribute("name")?.Value ?? "(unnamed)";
}
```

- [ ] **Step 3: Run tests, commit**

Run: `dotnet test --filter HierarchyParserTests` → PASS (3/3); full `dotnet test` green.

```powershell
git add src tests
git commit -m "feat: parse OneNote hierarchy XML into NotebookTree"
git push
```

---

### Task 3: OneNotePageXmlBuilder + DataUriInliner

**Files:**
- Create: `src/SendToOneNote.Core/Desktop/OneNotePageXmlBuilder.cs`, `src/SendToOneNote.Core/Pages/DataUriInliner.cs`
- Test: `tests/SendToOneNote.Tests/OneNotePageXmlBuilderTests.cs`, `tests/SendToOneNote.Tests/DataUriInlinerTests.cs`

**Interfaces:**
- Consumes: `ResolvedImage` (existing), `OneNoteConstants.Namespace2013`.
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public static class DataUriInliner { public static string Inline(string xhtml, IReadOnlyList<ResolvedImage> images); } // src="name:imgN" → src="data:<ct>;base64,…"
namespace SendToOneNote.Core.Desktop;
public static class OneNotePageXmlBuilder
{   public static string Build(string pageId, string title, string html);   // one:Page with Title + HTMLBlock; CDATA-safe
    public static string ExtractTitle(string pageXhtml);                     // <title>…</title> HTML-decoded, else "(no subject)"
    public static string CdataEscape(string s); }                            // "]]>" → "]]]]><![CDATA[>"
```

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/DataUriInlinerTests.cs`:

```csharp
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class DataUriInlinerTests
{
    [Fact]
    public void RewritesEachPartNameToADataUri()
    {
        var images = new List<ResolvedImage>
        {
            new("img0", "image/png", [1, 2, 3]),
            new("img1", "image/jpeg", [4, 5]),
        };
        var xhtml = "<img alt=\"a\" src=\"name:img0\" width=\"10\"/><img src=\"name:img1\"/>";
        var result = DataUriInliner.Inline(xhtml, images);
        Assert.Contains("src=\"data:image/png;base64,AQID\"", result);
        Assert.Contains("src=\"data:image/jpeg;base64,BAU=\"", result);
        Assert.Contains("alt=\"a\"", result); // other attributes untouched
        Assert.DoesNotContain("name:img", result);
    }

    [Fact]
    public void DoesNotConfuseImg1WithImg10()
    {
        var images = new List<ResolvedImage> { new("img1", "image/png", [1]) };
        var result = DataUriInliner.Inline("<img src=\"name:img10\"/><img src=\"name:img1\"/>", images);
        Assert.Contains("src=\"name:img10\"", result);
        Assert.Contains("src=\"data:image/png;base64,AQ==\"", result);
    }
}
```

`tests/SendToOneNote.Tests/OneNotePageXmlBuilderTests.cs`:

```csharp
using System.Xml.Linq;
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class OneNotePageXmlBuilderTests
{
    [Fact]
    public void BuildsPageWithTitleAndHtmlBlock()
    {
        var xml = OneNotePageXmlBuilder.Build("{P1}{1}{B0}", "Hello <b>", "<html><body><p>hi</p></body></html>");
        var doc = XDocument.Parse(xml);
        XNamespace one = OneNoteConstants.Namespace2013;
        var page = doc.Root!;
        Assert.Equal(one + "Page", page.Name);
        Assert.Equal("{P1}{1}{B0}", page.Attribute("ID")!.Value);
        Assert.Equal("Hello <b>", page.Element(one + "Title")!.Element(one + "OE")!.Element(one + "T")!.Value);
        var block = page.Element(one + "Outline")!.Element(one + "OEChildren")!.Element(one + "HTMLBlock")!;
        Assert.Equal("<html><body><p>hi</p></body></html>", block.Element(one + "Data")!.Value);
    }

    [Fact]
    public void CdataTerminatorInsidePayloadIsEscaped()
    {
        var xml = OneNotePageXmlBuilder.Build("{P}", "t", "<p>x]]>y</p>");
        var doc = XDocument.Parse(xml); // must remain well-formed
        XNamespace one = OneNoteConstants.Namespace2013;
        Assert.Equal("<p>x]]>y</p>", doc.Descendants(one + "Data").Single().Value);
    }

    [Fact]
    public void ExtractTitleDecodesEntitiesAndFallsBack()
    {
        Assert.Equal("S & T", OneNotePageXmlBuilder.ExtractTitle("<html><head><title>S &amp; T</title></head><body/></html>"));
        Assert.Equal("(no subject)", OneNotePageXmlBuilder.ExtractTitle("<html><body/></html>"));
    }
}
```

Run: `dotnet test --filter "DataUriInlinerTests|OneNotePageXmlBuilderTests"`
Expected: FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Pages/DataUriInliner.cs`:

```csharp
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Pages;

public static class DataUriInliner
{
    /// <summary>Rewrites src="name:{part}" to a base64 data URI for every resolved image.</summary>
    public static string Inline(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        foreach (var img in images)
        {
            var pattern = $"src=\"name:{Regex.Escape(img.PartName)}\"";
            var replacement = $"src=\"data:{img.ContentType};base64,{Convert.ToBase64String(img.Data)}\"";
            xhtml = Regex.Replace(xhtml, pattern, replacement.Replace("$", "$$"));
        }
        return xhtml;
    }
}
```

`src/SendToOneNote.Core/Desktop/OneNotePageXmlBuilder.cs`:

```csharp
using System.Net;
using System.Text.RegularExpressions;

namespace SendToOneNote.Core.Desktop;

public static class OneNotePageXmlBuilder
{
    public static string Build(string pageId, string title, string html) =>
        $"""
        <?xml version="1.0"?>
        <one:Page xmlns:one="{OneNoteConstants.Namespace2013}" ID="{pageId}">
          <one:Title><one:OE><one:T><![CDATA[{CdataEscape(title)}]]></one:T></one:OE></one:Title>
          <one:Outline><one:OEChildren>
            <one:HTMLBlock><one:Data><![CDATA[{CdataEscape(html)}]]></one:Data></one:HTMLBlock>
          </one:OEChildren></one:Outline>
        </one:Page>
        """;

    public static string CdataEscape(string s) => s.Replace("]]>", "]]]]><![CDATA[>");

    public static string ExtractTitle(string pageXhtml)
    {
        var m = Regex.Match(pageXhtml, "<title>(.*?)</title>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
        var raw = m.Success ? WebUtility.HtmlDecode(m.Groups[1].Value).Trim() : "";
        return raw.Length == 0 ? "(no subject)" : raw;
    }
}
```

- [ ] **Step 3: Run tests, commit**

Run: `dotnet test --filter "DataUriInlinerTests|OneNotePageXmlBuilderTests"` → PASS (5/5); full suite green.

```powershell
git add src tests
git commit -m "feat: OneNote page XML builder and data-URI image inliner"
git push
```

---

### Task 4: IOneNoteBackend, GraphBackend, DesktopOneNoteBackend

**Files:**
- Create: `src/SendToOneNote.Core/Backends/IOneNoteBackend.cs`, `src/SendToOneNote.Core/Backends/GraphBackend.cs`, `src/SendToOneNote.Core/Backends/DesktopOneNoteBackend.cs`
- Test: `tests/SendToOneNote.Tests/DesktopOneNoteBackendTests.cs`, `tests/SendToOneNote.Tests/GraphBackendTests.cs`, `tests/SendToOneNote.Tests/FakeOneNoteApplication.cs`

**Interfaces:**
- Consumes: Tasks 1–3; `OneNoteClient`, `PagePlanner`, `CreatedPage` (existing).
- Produces:

```csharp
namespace SendToOneNote.Core.Backends;
public interface IOneNoteBackend
{   string Name { get; }  // "desktop" | "graph"
    Task<NotebookTree> GetTreeAsync(CancellationToken ct = default);
    Task<CreatedPage> CreatePageAsync(string sectionId, string pageXhtml, IReadOnlyList<ResolvedImage> images, CancellationToken ct = default); }
public sealed class GraphBackend(OneNoteClient client) : IOneNoteBackend;
public sealed class DesktopOneNoteException(int hresult, string message) : Exception(message) { public int HResultCode { get; } }
public sealed class DesktopOneNoteBackend(StaComWorker worker, Func<IApplication>? applicationFactory = null) : IOneNoteBackend;
```

- [ ] **Step 1: Write the fake COM application for tests**

`tests/SendToOneNote.Tests/FakeOneNoteApplication.cs`:

```csharp
using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

/// <summary>Managed implementation of the COM interface — records calls, returns canned data.</summary>
public sealed class FakeOneNoteApplication : IApplication
{
    public string HierarchyXml { get; set; } = """
    <one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
      <one:Notebook name="Alpha" ID="{N1}"><one:Section name="Inbox" ID="{S1}"/></one:Notebook>
    </one:Notebooks>
    """;
    public string NextPageId { get; set; } = "{P1}{1}{B0}";
    public string Hyperlink { get; set; } = "onenote:https://example.test/Alpha/Inbox.one#p1";
    public List<(string SectionId, int Style)> CreatedPages { get; } = [];
    public List<string> UpdatedPageXml { get; } = [];
    public int ManagedThreadIdOfLastCall { get; private set; }
    public Exception? ThrowOnUpdate { get; set; }

    public void GetHierarchy(string bstrStartNodeID, int hsScope, out string pbstrHierarchyXmlOut, int xsSchema)
    { Touch(); pbstrHierarchyXmlOut = HierarchyXml; }
    public void CreateNewPage(string bstrSectionID, out string pbstrPageID, int npsNewPageStyle)
    { Touch(); CreatedPages.Add((bstrSectionID, npsNewPageStyle)); pbstrPageID = NextPageId; }
    public void UpdatePageContent(string bstrPageChangesXmlIn, DateTime dateExpectedLastModified, int xsSchema, bool force)
    { Touch(); if (ThrowOnUpdate is not null) throw ThrowOnUpdate; UpdatedPageXml.Add(bstrPageChangesXmlIn); }
    public void GetHyperlinkToObject(string bstrHierarchyID, string bstrPageContentObjectID, out string pbstrHyperlinkOut)
    { Touch(); pbstrHyperlinkOut = Hyperlink; }

    private void Touch() => ManagedThreadIdOfLastCall = Environment.CurrentManagedThreadId;

    // Unused members of the interface:
    public void UpdateHierarchy(string a, int b) => throw new NotImplementedException();
    public void OpenHierarchy(string a, string b, out string c, int d) => throw new NotImplementedException();
    public void DeleteHierarchy(string a, DateTime b, bool c) => throw new NotImplementedException();
    public void CloseNotebook(string a, bool b) => throw new NotImplementedException();
    public void GetHierarchyParent(string a, out string b) => throw new NotImplementedException();
    public void GetPageContent(string a, out string b, int c, int d) => throw new NotImplementedException();
    public void GetBinaryPageContent(string a, string b, out string c) => throw new NotImplementedException();
    public void DeletePageContent(string a, string b, DateTime c, bool d) => throw new NotImplementedException();
    public void NavigateTo(string a, string b, bool c) => throw new NotImplementedException();
    public void NavigateToUrl(string a, bool b) => throw new NotImplementedException();
    public void Publish(string a, string b, int c, string d) => throw new NotImplementedException();
    public void OpenPackage(string a, string b, out string c) => throw new NotImplementedException();
}
```

- [ ] **Step 2: Write failing backend tests**

`tests/SendToOneNote.Tests/DesktopOneNoteBackendTests.cs`:

```csharp
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
```

`tests/SendToOneNote.Tests/GraphBackendTests.cs`:

```csharp
using System.Net;
using System.Text;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Tests;

file sealed class FakeTokens : ITokenProvider
{
    public string? SignedInUser => "test@example.com";
    public Task<string> GetAccessTokenAsync(bool interactiveAllowed, CancellationToken ct = default) => Task.FromResult("T");
}

public class GraphBackendTests
{
    [Fact]
    public async Task CreatePagePlansAndPostsThroughOneNoteClient()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(HttpStatusCode.Created)
        {
            Content = new StringContent("""{"id":"p1","links":{"oneNoteClientUrl":{"href":"onenote:x"}}}""", Encoding.UTF8, "application/json")
        });
        var backend = new GraphBackend(new OneNoteClient(new FakeTokens(), stub));
        var page = await backend.CreatePageAsync("s1",
            "<html><head><title>t</title></head><body><img src=\"name:img0\"/></body></html>",
            [new("img0", "image/png", StubPng())]);
        Assert.Equal("p1", page.Id);
        Assert.Equal("graph", backend.Name);
        var req = Assert.Single(stub.Requests);
        Assert.Contains("/me/onenote/sections/s1/pages", req.RequestUri!.ToString());
    }

    private static byte[] StubPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
}
```

Run: `dotnet test --filter "DesktopOneNoteBackendTests|GraphBackendTests"`
Expected: FAIL.

- [ ] **Step 3: Implement**

`src/SendToOneNote.Core/Backends/IOneNoteBackend.cs`:

```csharp
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
```

`src/SendToOneNote.Core/Backends/GraphBackend.cs`:

```csharp
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
```

`src/SendToOneNote.Core/Backends/DesktopOneNoteBackend.cs`:

```csharp
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
```

- [ ] **Step 4: Run tests, commit**

Run: `dotnet test --filter "DesktopOneNoteBackendTests|GraphBackendTests"` → PASS (5/5); full suite green.

```powershell
git add src tests
git commit -m "feat: IOneNoteBackend with desktop (COM) and Graph implementations"
git push
```

---

### Task 5: Settings additions + BackendSelector

**Files:**
- Modify: `src/SendToOneNote.Core/Storage/AppSettings.cs`
- Create: `src/SendToOneNote.Core/Backends/BackendSelector.cs`
- Test: `tests/SendToOneNote.Tests/BackendSelectorTests.cs`, extend `tests/SendToOneNote.Tests/StorageTests.cs`

**Interfaces:**
- Consumes: `AppSettings`, `JsonFileStore` (existing).
- Produces:

```csharp
// AppSettings additions
public string Backend { get; set; } = "auto";              // "auto" | "desktop" | "graph"
public bool ImageDiagnostics { get; set; }                 // default false
public List<string> RecentDesktopSectionIds { get; set; } = [];

namespace SendToOneNote.Core.Backends;
public enum BackendKind { Desktop, Graph }
public sealed record BackendChoice(BackendKind Kind, string Reason);
public static class BackendSelector
{   // desktopAvailable is evaluated only when needed (auto/desktop). Throws InvalidOperationException when "desktop" is forced but unavailable.
    public static BackendChoice Choose(string? setting, Func<bool> desktopAvailable); }
```

- [ ] **Step 1: Write failing tests**

Append to `tests/SendToOneNote.Tests/StorageTests.cs` (inside the class):

```csharp
    [Fact]
    public void NewSettingsHaveSafeDefaultsAndRoundTrip()
    {
        var store = new JsonFileStore(_dir);
        var s = store.LoadSettings();
        Assert.Equal("auto", s.Backend);
        Assert.False(s.ImageDiagnostics);
        Assert.Empty(s.RecentDesktopSectionIds);
        s.Backend = "graph"; s.ImageDiagnostics = true; s.RecentDesktopSectionIds.Add("{S1}");
        store.SaveSettings(s);
        var s2 = new JsonFileStore(_dir).LoadSettings();
        Assert.Equal("graph", s2.Backend);
        Assert.True(s2.ImageDiagnostics);
        Assert.Equal(["{S1}"], s2.RecentDesktopSectionIds);
    }

    [Fact]
    public void OldSettingsFileWithoutNewKeysStillLoads()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(Path.Combine(_dir, "settings.json"), """{"DropFolder":"C:\\Drop","DeleteOnSuccess":true,"RecentSectionIds":["a"]}""");
        var s = new JsonFileStore(_dir).LoadSettings();
        Assert.Equal("auto", s.Backend);
        Assert.Equal(["a"], s.RecentSectionIds);
    }
```

`tests/SendToOneNote.Tests/BackendSelectorTests.cs`:

```csharp
using SendToOneNote.Core.Backends;

namespace SendToOneNote.Tests;

public class BackendSelectorTests
{
    [Fact]
    public void AutoPrefersDesktopWhenAvailable() =>
        Assert.Equal(BackendKind.Desktop, BackendSelector.Choose("auto", () => true).Kind);

    [Fact]
    public void AutoFallsBackToGraphWhenDesktopMissing() =>
        Assert.Equal(BackendKind.Graph, BackendSelector.Choose("auto", () => false).Kind);

    [Fact]
    public void GraphForcedSkipsDesktopWithoutProbing()
    {
        var probed = false;
        var choice = BackendSelector.Choose("graph", () => { probed = true; return true; });
        Assert.Equal(BackendKind.Graph, choice.Kind);
        Assert.False(probed);
    }

    [Fact]
    public void DesktopForcedButUnavailableThrows() =>
        Assert.Throws<InvalidOperationException>(() => BackendSelector.Choose("desktop", () => false));

    [Fact]
    public void NullOrUnknownSettingBehavesLikeAuto()
    {
        Assert.Equal(BackendKind.Desktop, BackendSelector.Choose(null, () => true).Kind);
        Assert.Equal(BackendKind.Graph, BackendSelector.Choose("banana", () => false).Kind);
    }
}
```

Run: `dotnet test --filter "BackendSelectorTests|StorageTests"` → FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Storage/AppSettings.cs` — add three properties:

```csharp
namespace SendToOneNote.Core.Storage;

public sealed class AppSettings
{
    public string? DropFolder { get; set; }
    public string? ClientIdOverride { get; set; }
    public bool DeleteOnSuccess { get; set; } = true;
    public List<string> RecentSectionIds { get; set; } = [];

    /// <summary>"auto" (try desktop OneNote first) | "desktop" (force COM) | "graph" (force cloud).</summary>
    public string Backend { get; set; } = "auto";
    /// <summary>Write &lt;DropFolder&gt;\Diagnostics\&lt;email&gt;\images.csv (+ images) per save.</summary>
    public bool ImageDiagnostics { get; set; }
    /// <summary>Recents for the desktop backend — section IDs differ from Graph's.</summary>
    public List<string> RecentDesktopSectionIds { get; set; } = [];
}
```

`src/SendToOneNote.Core/Backends/BackendSelector.cs`:

```csharp
namespace SendToOneNote.Core.Backends;

public enum BackendKind { Desktop, Graph }

public sealed record BackendChoice(BackendKind Kind, string Reason);

public static class BackendSelector
{
    public static BackendChoice Choose(string? setting, Func<bool> desktopAvailable)
    {
        switch ((setting ?? "auto").Trim().ToLowerInvariant())
        {
            case "graph":
                return new(BackendKind.Graph, "Backend=graph (forced)");
            case "desktop":
                return desktopAvailable()
                    ? new(BackendKind.Desktop, "Backend=desktop (forced)")
                    : throw new InvalidOperationException(
                        "Backend is set to \"desktop\" but desktop OneNote is not available on this machine. " +
                        "Install desktop OneNote or set Backend to \"auto\" or \"graph\" in settings.json.");
            default:
                return desktopAvailable()
                    ? new(BackendKind.Desktop, "auto: desktop OneNote detected")
                    : new(BackendKind.Graph, "auto: desktop OneNote not available");
        }
    }
}
```

- [ ] **Step 3: Run tests, commit**

Run: `dotnet test --filter "BackendSelectorTests|StorageTests"` → PASS; full suite green.

```powershell
git add src tests
git commit -m "feat: Backend/ImageDiagnostics settings and BackendSelector"
git push
```

---

### Task 6: Junk-image filter + image decisions in ImageResolver

**Files:**
- Modify: `src/SendToOneNote.Core/Pages/ImageResolver.cs`
- Create: `src/SendToOneNote.Core/Pages/ImageDecision.cs`
- Test: extend `tests/SendToOneNote.Tests/ImageResolverTests.cs`

**Interfaces:**
- Consumes: `ResolvedImage`, `InlineImage`, `ImageShrinker` (System.Drawing already referenced by Core).
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data, int Width = 0, int Height = 0); // dims added, defaults keep existing positional callers working
public sealed record ImageDecision(int Index, string Src, string? PartName, int Bytes, int Width, int Height, string Alt,
    string Source /* cid|remote|other */, string Decision /* embedded|dropped-junk|left-as-url */, string Reason);
public sealed record ImageResolution(string Xhtml, IReadOnlyList<ResolvedImage> Images, IReadOnlyList<ImageDecision> Decisions);
// ImageResolver: new method; the old ResolveAsync stays as a thin wrapper returning (Xhtml, Images)
public Task<ImageResolution> ResolveWithReportAsync(string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default);
```

- Junk rules (any → `dropped-junk`, `<img>` removed from the DOM): remote URL matching `(?i)(open|track|pixel|beacon|spacer|blank)\.(gif|png)(\?|$)` or containing `/o.gif`; HTML `width` and `height` attributes both present and both ≤ 2; decoded dimensions ≤ 2×2 (via `System.Drawing.Image.FromStream`; undecodable → dimensions 0, not junk).

- [ ] **Step 1: Write failing tests**

Append to `tests/SendToOneNote.Tests/ImageResolverTests.cs` (inside the class; reuse its `PngBytes` 1×1 constant):

```csharp
    [Fact]
    public async Task TrackingPixelUrlIsDroppedAndRemovedFromHtml()
    {
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/open.gif?u=1\"/><p>text</p></body></html>", []);
        Assert.Empty(r.Images);
        Assert.DoesNotContain("<img", r.Xhtml);
        var d = Assert.Single(r.Decisions);
        Assert.Equal("dropped-junk", d.Decision);
        Assert.Contains("tracking", d.Reason);
    }

    [Fact]
    public async Task OneByOneDecodedImageIsDropped()
    {
        // PngBytes is a 1x1 PNG served from an innocent-looking URL.
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(PngBytes));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/logo.png\"/></body></html>", []);
        Assert.Empty(r.Images);
        Assert.Equal("dropped-junk", Assert.Single(r.Decisions).Decision);
    }

    [Fact]
    public async Task TinyWidthHeightAttributesAreDroppedWithoutDownloading()
    {
        var stub = new StubHttpHandler(_ => throw new InvalidOperationException("must not download"));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/s.png\" width=\"1\" height=\"1\"/></body></html>", []);
        Assert.Empty(r.Images);
        Assert.Equal("dropped-junk", Assert.Single(r.Decisions).Decision);
        Assert.Empty(stub.Requests);
    }

    [Fact]
    public async Task RealImageIsEmbeddedWithDimensionsAndDecision()
    {
        using var bmp = new System.Drawing.Bitmap(40, 30);
        using var ms = new MemoryStream();
        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        var stub = new StubHttpHandler(_ => StubHttpHandler.Png(ms.ToArray()));
        var r = await new ImageResolver(stub).ResolveWithReportAsync(
            "<html><head><title>t</title></head><body><img alt=\"hero\" src=\"https://x.example/hero.png\"/></body></html>", []);
        var img = Assert.Single(r.Images);
        Assert.Equal((40, 30), (img.Width, img.Height));
        var d = Assert.Single(r.Decisions);
        Assert.Equal("embedded", d.Decision);
        Assert.Equal("img0", d.PartName);
        Assert.Equal("hero", d.Alt);
        Assert.Equal("remote", d.Source);
    }

    [Fact]
    public async Task LegacyResolveAsyncStillReturnsTuple()
    {
        var stub = new StubHttpHandler(_ => new HttpResponseMessage(System.Net.HttpStatusCode.NotFound));
        var (xhtml, images) = await new ImageResolver(stub).ResolveAsync(
            "<html><head><title>t</title></head><body><img src=\"https://x.example/gone.png\"/></body></html>", []);
        Assert.Empty(images);
        Assert.Contains("gone.png", xhtml);
    }
```

Run: `dotnet test --filter ImageResolverTests` → FAIL (new members missing).

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Pages/ImageDecision.cs`:

```csharp
namespace SendToOneNote.Core.Pages;

public sealed record ImageDecision(int Index, string Src, string? PartName, int Bytes, int Width, int Height, string Alt,
    string Source, string Decision, string Reason);

public sealed record ImageResolution(string Xhtml, IReadOnlyList<ResolvedImage> Images, IReadOnlyList<ImageDecision> Decisions);
```

Replace `src/SendToOneNote.Core/Pages/ImageResolver.cs` in full:

```csharp
using System.Text.RegularExpressions;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using AngleSharp.Xhtml;
using SendToOneNote.Core.Email;

namespace SendToOneNote.Core.Pages;

public sealed record ResolvedImage(string PartName, string ContentType, byte[] Data, int Width = 0, int Height = 0);

public sealed class ImageResolver
{
    private static readonly Regex TrackingUrl = new(
        @"(?i)(open|track|pixel|beacon|spacer|blank)\.(gif|png)(\?|$)|/o\.gif", RegexOptions.Compiled);

    private readonly HttpClient _http;

    public ImageResolver(HttpMessageHandler? handler = null)
    {
        _http = handler is null ? new HttpClient() : new HttpClient(handler);
        _http.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task<(string Xhtml, IReadOnlyList<ResolvedImage> Images)> ResolveAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var r = await ResolveWithReportAsync(pageXhtml, inlineImages, ct);
        return (r.Xhtml, r.Images);
    }

    public async Task<ImageResolution> ResolveWithReportAsync(
        string pageXhtml, IReadOnlyList<InlineImage> inlineImages, CancellationToken ct = default)
    {
        var doc = await new HtmlParser().ParseDocumentAsync(pageXhtml, ct);
        var imgs = doc.QuerySelectorAll("img").ToList();
        var gate = new SemaphoreSlim(4);

        var work = imgs.Select(async (img, index) =>
        {
            var src = img.GetAttribute("src") ?? "";
            var alt = img.GetAttribute("alt") ?? "";
            var source = src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase) ? "cid"
                : src.StartsWith("http", StringComparison.OrdinalIgnoreCase) ? "remote" : "other";

            // Junk by markup alone — no download needed.
            if (source == "remote" && TrackingUrl.IsMatch(src))
                return (img, (byte[]?)null, (string?)null, Junk(index, src, alt, source, "tracking-pixel URL pattern"));
            if (TryInt(img.GetAttribute("width"), out var aw) && TryInt(img.GetAttribute("height"), out var ah) && aw <= 2 && ah <= 2)
                return (img, null, null, Junk(index, src, alt, source, $"declared size {aw}x{ah}"));

            byte[]? data = null; string? contentType = null;
            if (source == "cid")
            {
                var cid = src[4..].Trim('<', '>');
                var match = inlineImages.FirstOrDefault(i => string.Equals(i.ContentId, cid, StringComparison.OrdinalIgnoreCase));
                if (match is not null) { data = match.Data; contentType = match.ContentType; }
            }
            else if (source == "remote")
            {
                await gate.WaitAsync(ct);
                try
                {
                    var resp = await _http.GetAsync(src, ct);
                    if (resp.IsSuccessStatusCode)
                    {
                        data = await resp.Content.ReadAsByteArrayAsync(ct);
                        contentType = resp.Content.Headers.ContentType?.MediaType ?? "image/png";
                    }
                }
                catch (Exception) when (!ct.IsCancellationRequested) { /* leave original src */ }
                finally { gate.Release(); }
            }

            if (data is null || contentType is null)
                return (img, null, null, new ImageDecision(index, src, null, 0, 0, 0, alt, source, "left-as-url", "not downloadable / no matching cid"));

            var (w, h) = Dimensions(data);
            if (w > 0 && w <= 2 && h <= 2)
                return (img, null, null, Junk(index, src, alt, source, $"decoded size {w}x{h}"));

            return (img, data, contentType, new ImageDecision(index, src, null, data.Length, w, h, alt, source, "embedded", "ok"));
        }).ToList();

        var results = await Task.WhenAll(work);
        var resolved = new List<ResolvedImage>();
        var decisions = new List<ImageDecision>();
        foreach (var (img, data, contentType, decision) in results)
        {
            if (decision.Decision == "dropped-junk") { img.Remove(); decisions.Add(decision); continue; }
            if (data is null || contentType is null) { decisions.Add(decision); continue; }
            var name = $"img{resolved.Count}";
            resolved.Add(new ResolvedImage(name, contentType, data, decision.Width, decision.Height));
            img.SetAttribute("src", $"name:{name}");
            decisions.Add(decision with { PartName = name });
        }

        return new ImageResolution(doc.ToHtml(XhtmlMarkupFormatter.Instance), resolved, decisions);
    }

    private static ImageDecision Junk(int index, string src, string alt, string source, string reason) =>
        new(index, src, null, 0, 0, 0, alt, source, "dropped-junk", reason);

    private static bool TryInt(string? s, out int v) => int.TryParse(s?.Trim().TrimEnd('p', 'x'), out v);

    private static (int W, int H) Dimensions(byte[] data)
    {
        try
        {
            using var ms = new MemoryStream(data);
            using var image = System.Drawing.Image.FromStream(ms, useEmbeddedColorManagement: false, validateImageData: false);
            return (image.Width, image.Height);
        }
        catch (Exception) { return (0, 0); }
    }
}
```

- [ ] **Step 3: Run tests, commit**

Run: `dotnet test --filter ImageResolverTests` → PASS (all, including the five pre-existing); full suite green. Note the existing `RemoteImageDownloadedAndRewritten` test serves a 1×1 PNG from a non-tracking URL — it now expects that image to be **dropped as junk**. Update that test's assertions to `Assert.Empty(images)` and `Assert.DoesNotContain("<img", xhtml)`; the `CidImageResolvedFromInlineParts` test likewise uses the 1×1 — change its inline image to a 40×30 bitmap (generate as in `RealImageIsEmbeddedWithDimensionsAndDecision`) so it still asserts embedding.

```powershell
git add src tests
git commit -m "feat: junk-image filter and per-image decisions in ImageResolver"
git push
```

---

### Task 7: Graph-path image ranking (embed important, drop minor; no appends)

**Files:**
- Modify: `src/SendToOneNote.Core/Pages/PagePlanner.cs`, `tests/SendToOneNote.Tests/PagePlannerTests.cs`, `tests/SendToOneNote.Tests/IntegrationSmokeTests.cs`
- Modify: `agents/knowledge/architecture.md` (planner description)

**Interfaces:**
- Consumes: `ResolvedImage` with `Width/Height` (Task 6).
- Produces: `PagePlan` gains `public IReadOnlyList<string> DroppedPartNames { get; init; } = [];` (positional ctor unchanged). `Plan(...)` now: shrink → drop oversize (as today) → **rank** by score = area (Width×Height, else decoded, else Data.Length) × 4 if among the first 3 images in document order → take in rank order while `Parts.Count < 30` and bytes fit the 3.5 MB budget → **drop** the rest (their `<img>` removed from the XHTML, names in `DroppedPartNames`) → `Appends` is always empty. `OneNoteClient`'s append/poll code remains (unused by the planner, still unit-tested).

- [ ] **Step 1: Rewrite the planner tests**

Replace `OverflowImagesBecomeSlotsAndAppends` and `OverflowReplacementSurvivesImgAttributes` in `tests/SendToOneNote.Tests/PagePlannerTests.cs` with:

```csharp
    private static ResolvedImage Img(int i, int w, int h) => new($"img{i}", "image/png", Png, w, h);

    [Fact]
    public void ImagesBeyondPartCapAreDroppedNotAppended()
    {
        var images = Enumerable.Range(0, 33).Select(i => Img(i, 100, 100)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(33), images);
        Assert.Equal(30, plan.Parts.Count);
        Assert.Empty(plan.Appends);
        Assert.Equal(3, plan.DroppedPartNames.Count);
        foreach (var name in plan.DroppedPartNames)
            Assert.DoesNotContain($"name:{name}", plan.PresentationXhtml);
        Assert.DoesNotContain("data-id=\"slot-", plan.PresentationXhtml);
    }

    [Fact]
    public void LargerImagesWinWhenOverCap()
    {
        // 31 images: index 5 is huge, everything else tiny — the huge one must survive.
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i == 5 ? 800 : 10, i == 5 ? 600 : 10)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.Contains(plan.Parts, p => p.Name == "img5");
        Assert.Single(plan.DroppedPartNames);
    }

    [Fact]
    public void EarlyImagesGetABoost()
    {
        // 31 equal-size images: the first three (logo/banner position) must never be the ones dropped.
        var images = Enumerable.Range(0, 31).Select(i => Img(i, 50, 50)).ToList();
        var plan = PagePlanner.Plan(XhtmlWith(31), images);
        Assert.DoesNotContain("img0", plan.DroppedPartNames);
        Assert.DoesNotContain("img1", plan.DroppedPartNames);
        Assert.DoesNotContain("img2", plan.DroppedPartNames);
    }

    [Fact]
    public void DroppedImageTagRemovedRegardlessOfAttributes()
    {
        var imgs = string.Join("", Enumerable.Range(0, 31).Select(i =>
            $"<img alt=\"pic {i}\" width=\"600\" src=\"name:img{i}\" style=\"border:0\" />"));
        var xhtml = $"<html><head><title>t</title></head><body>{imgs}</body></html>";
        var images = Enumerable.Range(0, 31).Select(i => Img(i, i == 30 ? 1 : 50, i == 30 ? 1 : 50)).ToList();
        var plan = PagePlanner.Plan(xhtml, images);
        Assert.Equal(["img30"], plan.DroppedPartNames);
        Assert.DoesNotContain("pic 30", plan.PresentationXhtml);
        Assert.Contains("pic 29", plan.PresentationXhtml);
    }
```

Also in `tests/SendToOneNote.Tests/IntegrationSmokeTests.cs` (`CreatesRealPageInScratchSection`): replace the `Assert.True(plan.Appends.Count >= 1, ...)` with `Assert.Empty(plan.Appends); Assert.Equal(30, plan.Parts.Count);` and rename the page title to `"Integration smoke (30 images, one POST)"`.

Run: `dotnet test --filter PagePlannerTests` → FAIL.

- [ ] **Step 2: Implement the ranking planner**

Replace `PagePlanner.Plan` and remove `ToAppend` in `src/SendToOneNote.Core/Pages/PagePlanner.cs`:

```csharp
public sealed record PagePlan(string PresentationXhtml, IReadOnlyList<OneNoteRequestPart> Parts,
    IReadOnlyList<AppendPlan> Appends)
{
    public IReadOnlyList<string> DroppedPartNames { get; init; } = [];
}

public static class PagePlanner
{
    public const int MaxRequestBytes = 3_500_000;
    public const int MaxBinaryPartsPerRequest = 30;
    private const int EarlyImageCount = 3;   // logo/banner slots get a ranking boost
    private const int EarlyImageBoost = 4;

    private static Regex ImgTagRegex(string partName) =>
        new($"<img\\b[^>]*src=\"name:{Regex.Escape(partName)}\"[^>]*/>");

    /// <summary>
    /// Graph path only: everything fits in ONE request or is dropped. Images are ranked by
    /// rendered area (early ones boosted); the desktop backend never calls this.
    /// </summary>
    public static PagePlan Plan(string xhtml, IReadOnlyList<ResolvedImage> images)
    {
        var dropped = new List<string>();

        var shrunk = images.Select(i =>
        {
            var (data, ct) = ImageShrinker.ShrinkIfNeeded(i.Data, i.ContentType, MaxRequestBytes / 2);
            return i with { Data = data, ContentType = ct };
        }).ToList();

        var kept = new List<ResolvedImage>();
        foreach (var img in shrunk)
        {
            if (img.Data.Length > MaxRequestBytes - 4096)
            {
                xhtml = ImgTagRegex(img.PartName).Replace(xhtml, "<p style=\"color:#999999\">[image omitted: too large]</p>");
                dropped.Add(img.PartName);
            }
            else kept.Add(img);
        }

        var ranked = kept
            .Select((img, docIndex) => (img, score: Score(img) * (docIndex < EarlyImageCount ? EarlyImageBoost : 1)))
            .OrderByDescending(x => x.score)
            .Select(x => x.img)
            .ToList();

        var selected = new HashSet<string>();
        long budget = MaxRequestBytes - Encoding.UTF8.GetByteCount(xhtml) - 4096;
        foreach (var img in ranked)
        {
            if (selected.Count >= MaxBinaryPartsPerRequest || img.Data.Length > budget) continue;
            selected.Add(img.PartName);
            budget -= img.Data.Length;
        }

        foreach (var img in kept.Where(i => !selected.Contains(i.PartName)))
        {
            xhtml = ImgTagRegex(img.PartName).Replace(xhtml, "");
            dropped.Add(img.PartName);
        }

        var parts = kept.Where(i => selected.Contains(i.PartName))   // document order for stability
            .Select(i => new OneNoteRequestPart(i.PartName, i.ContentType, i.Data)).ToList();
        return new PagePlan(xhtml, parts, []) { DroppedPartNames = dropped };
    }

    private static long Score(ResolvedImage img)
    {
        if (img.Width > 0 && img.Height > 0) return (long)img.Width * img.Height;
        try
        {
            using var ms = new MemoryStream(img.Data);
            using var image = System.Drawing.Image.FromStream(ms, false, false);
            return (long)image.Width * image.Height;
        }
        catch (Exception) { return img.Data.Length; }
    }
}
```

Keep the `using System.Text; using System.Text.RegularExpressions;` and the `OneNoteRequestPart`/`AppendPlan` records; delete the old slot/append code and its comments.

- [ ] **Step 3: Update the architecture doc**

In `agents/knowledge/architecture.md`, replace the "Overflow images: …" bullet with: `- Graph path never appends any more: images are ranked by area (first three boosted) and embedded in one request up to the caps; the rest are dropped (removed from the XHTML) per the v1.1 spec. OneNoteClient keeps its append/poll support (tested, unused by the planner).`

- [ ] **Step 4: Run tests, commit**

Run: `dotnet test` → all green (the OneNoteClient append tests still pass — they construct `AppendPlan`s directly).

```powershell
git add src tests agents
git commit -m "feat: rank images on the Graph path; drop beyond caps instead of appending"
git push
```

---

### Task 8: Image diagnostics dump

**Files:**
- Create: `src/SendToOneNote.Core/Pages/ImageDiagnosticsWriter.cs`
- Test: `tests/SendToOneNote.Tests/ImageDiagnosticsWriterTests.cs`

**Interfaces:**
- Consumes: `ImageDecision`, `ResolvedImage` (Task 6), `PagePlan.DroppedPartNames` (Task 7).
- Produces:

```csharp
namespace SendToOneNote.Core.Pages;
public static class ImageDiagnosticsWriter
{   // Creates <dropFolder>\Diagnostics\<emlStem>-<yyyyMMdd-HHmmss>\ with images.csv and imgN.<ext>; returns the folder path.
    public static string Write(string dropFolder, string emlStem, IReadOnlyList<ImageDecision> decisions,
        IReadOnlyList<ResolvedImage> images, IReadOnlyList<string> droppedMinorPartNames, DateTime now); }
```

- CSV header exactly: `index,src,part,bytes,width,height,alt,source,decision,reason`; decision for a part in `droppedMinorPartNames` becomes `dropped-minor`; values are CSV-quoted (double quotes doubled). Extension from content type: `image/jpeg`→`jpg`, `image/png`→`png`, `image/gif`→`gif`, `image/webp`→`webp`, else `bin`.

- [ ] **Step 1: Write failing tests**

`tests/SendToOneNote.Tests/ImageDiagnosticsWriterTests.cs`:

```csharp
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class ImageDiagnosticsWriterTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-diag-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void WritesCsvAndImageFiles()
    {
        var decisions = new List<ImageDecision>
        {
            new(0, "https://x.example/a.png", "img0", 3, 40, 30, "hero, \"big\"", "remote", "embedded", "ok"),
            new(1, "https://x.example/open.gif", null, 0, 0, 0, "", "remote", "dropped-junk", "tracking-pixel URL pattern"),
            new(2, "https://x.example/b.jpg", "img1", 2, 5, 5, "", "remote", "embedded", "ok"),
        };
        var images = new List<ResolvedImage> { new("img0", "image/png", [1, 2, 3], 40, 30), new("img1", "image/jpeg", [4, 5], 5, 5) };

        var folder = ImageDiagnosticsWriter.Write(_dir, "The robotics race", decisions, images, ["img1"],
            new DateTime(2026, 8, 28, 9, 5, 7));

        Assert.Equal(Path.Combine(_dir, "Diagnostics", "The robotics race-20260828-090507"), folder);
        var lines = File.ReadAllLines(Path.Combine(folder, "images.csv"));
        Assert.Equal("index,src,part,bytes,width,height,alt,source,decision,reason", lines[0]);
        Assert.Contains("\"hero, \"\"big\"\"\"", lines[1]);            // CSV quoting
        Assert.Contains(",dropped-junk,", lines[2]);
        Assert.Contains(",dropped-minor,", lines[3]);                  // planner drop overrides "embedded"
        Assert.True(File.Exists(Path.Combine(folder, "img0.png")));
        Assert.True(File.Exists(Path.Combine(folder, "img1.jpg")));
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(folder, "img0.png")));
    }

    [Fact]
    public void SanitizesInvalidFolderCharactersInStem()
    {
        var folder = ImageDiagnosticsWriter.Write(_dir, "Re: fwd/ \"quoted\"?", [], [], [], new DateTime(2026, 1, 1));
        Assert.True(Directory.Exists(folder));
        Assert.DoesNotContain("/", Path.GetFileName(folder));
    }
}
```

Run: `dotnet test --filter ImageDiagnosticsWriterTests` → FAIL.

- [ ] **Step 2: Implement**

`src/SendToOneNote.Core/Pages/ImageDiagnosticsWriter.cs`:

```csharp
using System.Text;

namespace SendToOneNote.Core.Pages;

public static class ImageDiagnosticsWriter
{
    public static string Write(string dropFolder, string emlStem, IReadOnlyList<ImageDecision> decisions,
        IReadOnlyList<ResolvedImage> images, IReadOnlyList<string> droppedMinorPartNames, DateTime now)
    {
        var stem = string.Concat(emlStem.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c)).Trim();
        var folder = Path.Combine(dropFolder, "Diagnostics", $"{stem}-{now:yyyyMMdd-HHmmss}");
        Directory.CreateDirectory(folder);

        var minor = new HashSet<string>(droppedMinorPartNames);
        var sb = new StringBuilder("index,src,part,bytes,width,height,alt,source,decision,reason\n");
        foreach (var d in decisions)
        {
            var decision = d.PartName is not null && minor.Contains(d.PartName) ? "dropped-minor" : d.Decision;
            sb.Append(string.Join(",", new[]
            {
                d.Index.ToString(), Q(d.Src), Q(d.PartName ?? ""), d.Bytes.ToString(), d.Width.ToString(), d.Height.ToString(),
                Q(d.Alt), d.Source, decision, Q(d.Reason)
            })).Append('\n');
        }
        File.WriteAllText(Path.Combine(folder, "images.csv"), sb.ToString(), Encoding.UTF8);

        foreach (var img in images)
            File.WriteAllBytes(Path.Combine(folder, $"{img.PartName}.{Ext(img.ContentType)}"), img.Data);

        return folder;
    }

    private static string Q(string s) => $"\"{s.Replace("\"", "\"\"")}\"";

    private static string Ext(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/jpeg" => "jpg", "image/png" => "png", "image/gif" => "gif", "image/webp" => "webp", _ => "bin"
    };
}
```

- [ ] **Step 3: Run tests, commit**

Run: `dotnet test --filter ImageDiagnosticsWriterTests` → PASS; full suite green.

```powershell
git add src tests
git commit -m "feat: per-email image diagnostics dump (CSV + image files)"
git push
```

---

### Task 9: App wiring — backend selection, no-sign-in first-run, pipeline on IOneNoteBackend

**Files:**
- Modify: `src/SendToOneNote/SavePipeline.cs`, `src/SendToOneNote/TrayContext.cs`, `src/SendToOneNote/FirstRunWindow.xaml`, `src/SendToOneNote/FirstRunWindow.xaml.cs`

**Interfaces:**
- Consumes: everything above.
- Produces: the runnable v1.1 app. `SavePipeline(JsonFileStore store, IOneNoteBackend backend, FileLog log)`; `FirstRunWindow(ITokenProvider? tokens, AppSettings settings)` (null tokens = desktop mode, no sign-in UI).

No automated UI tests (by design); each step ends with a manual check the owner performs.

- [ ] **Step 1: SavePipeline on the backend seam**

Edit `src/SendToOneNote/SavePipeline.cs`:

1. Constructor becomes `public sealed class SavePipeline(JsonFileStore store, IOneNoteBackend backend, FileLog log)`; add `using SendToOneNote.Core.Backends;`.
2. Tree retrieval — replace the cache block with:

```csharp
            NotebookTree tree;
            if (backend.Name == "desktop")
            {
                tree = await backend.GetTreeAsync();            // ~100 ms local; always fresh
            }
            else
            {
                var cached = store.LoadTreeCache();
                tree = cached ?? await RefreshTreeAsync();
                if (cached is not null)                        // cache hit: refresh for next time
                    _ = RefreshTreeAsync().ContinueWith(t =>
                        log.Error("Background notebook refresh failed", t.Exception),
                        TaskContinuationOptions.OnlyOnFaulted);
            }
```

   and `RefreshTreeAsync` calls `backend.GetTreeAsync()`.
3. Recents — pick the list by backend:

```csharp
            var settings = store.LoadSettings();
            var recents = backend.Name == "desktop" ? settings.RecentDesktopSectionIds : settings.RecentSectionIds;
            var vm = new SectionPickerViewModel(tree, recents);
```

   and after the save: `if (backend.Name == "desktop") settings.RecentDesktopSectionIds = SectionPickerViewModel.PushRecent(recents, pick.SectionId); else settings.RecentSectionIds = SectionPickerViewModel.PushRecent(recents, pick.SectionId);`
4. Resolve + diagnostics + save — replace the `xhtml…CreatePageAsync` lines with:

```csharp
            var xhtml = PageXhtmlBuilder.Build(email);
            var resolution = await _images.ResolveWithReportAsync(xhtml, email.InlineImages);
            var page = await backend.CreatePageAsync(pick.SectionId, resolution.Xhtml, resolution.Images);
            if (settings.ImageDiagnostics)
            {
                try
                {
                    var droppedMinor = backend.Name == "graph"
                        ? PagePlanner.Plan(resolution.Xhtml, resolution.Images).DroppedPartNames
                        : [];
                    var folder = ImageDiagnosticsWriter.Write(Path.GetDirectoryName(path)!,
                        Path.GetFileNameWithoutExtension(path), resolution.Decisions, resolution.Images, droppedMinor, DateTime.Now);
                    log.Info($"Image diagnostics written to {folder}");
                }
                catch (Exception diagEx) { log.Error("Image diagnostics failed (save was fine)", diagEx); }
            }
```

5. Failure mapping — add before the generic fallback: `DesktopOneNoteException d => $"Desktop OneNote: {d.Message}",`.

Build: `dotnet build` — expect TrayContext to fail to compile until Step 3 (fine).

- [ ] **Step 2: FirstRunWindow desktop mode**

`src/SendToOneNote/FirstRunWindow.xaml`: give the intro `TextBlock` `x:Name="IntroText"`, and wrap the sign-in button + status in a `<StackPanel x:Name="SignInPanel">`.

`src/SendToOneNote/FirstRunWindow.xaml.cs`: constructor takes `ITokenProvider? tokens`; when null:

```csharp
        if (_tokens is null)
        {
            SignInPanel.Visibility = Visibility.Collapsed;
            IntroText.Text = "Saving to your desktop OneNote — no sign-in needed. Choose the folder you'll drag emails into.";
        }
```

   and `Finish_Click` only requires sign-in when `_tokens is not null && _tokens.SignedInUser is null`. `SignIn_Click` guards `_tokens is null` (return).

- [ ] **Step 3: TrayContext backend selection**

Edit `src/SendToOneNote/TrayContext.cs`:

1. Fields: `private readonly StaComWorker _comWorker = new();` and `private IOneNoteBackend? _backend;` (`using SendToOneNote.Core.Backends; using SendToOneNote.Core.Desktop;`).
2. In `Run()`, replace the token-provider construction and everything through `_pipeline = …` with:

```csharp
        var settings = _store.LoadSettings();
        BackendChoice choice;
        try
        {
            choice = BackendSelector.Choose(settings.Backend,
                () => _comWorker.RunAsync(() => DesktopOneNoteProbe.IsAvailable()).GetAwaiter().GetResult());
        }
        catch (InvalidOperationException ex)
        {
            MessageBox.Show(ex.Message, "SendToOneNote");
            Application.Current.Shutdown();
            return;
        }
        _log.Info($"Backend: {choice.Kind} — {choice.Reason}");

        ITokenProvider? tokens = null;
        if (choice.Kind == BackendKind.Desktop)
        {
            _backend = new DesktopOneNoteBackend(_comWorker);
        }
        else
        {
            _tokens = new MsalTokenProvider(_store.RootDir, settings.ClientIdOverride);
            tokens = _tokens;
            _backend = new GraphBackend(new OneNoteClient(_tokens));
        }

        if (settings.DropFolder is null)
        {
            var first = new FirstRunWindow(tokens, settings);
            first.ShowDialog();
            if (!first.Completed) { Application.Current.Shutdown(); return; }
            _store.SaveSettings(settings);
        }

        _pipeline = new SavePipeline(_store, _backend, _log);
```

3. Tray icon tooltip: `ToolTipText = choice.Kind == BackendKind.Desktop ? "SendToOneNote — desktop OneNote" : "SendToOneNote — cloud (Graph)"`.
4. Menu: add the "Sign in again" item only `if (_tokens is not null)`.
5. `Dispose()`: also `_comWorker.Dispose();`.

Build: `dotnet build` clean; `dotnet test` green (no test changes).

- [ ] **Step 4: Manual verification (owner)**

`dotnet run --project src/SendToOneNote` with a fresh settings file (rename `%APPDATA%\SendToOneNote\settings.json` temporarily): first-run shows NO sign-in button and the desktop note; tray tooltip says "desktop OneNote"; drag a plain email → picker lists all notebooks (including local-only) within a blink → page appears in OneNote immediately; drag an image-heavy newsletter → same, images present, no tracking pixel; set `"Backend": "graph"` in settings.json, restart → tooltip says cloud, sign-in flow as before, save works; set `"ImageDiagnostics": true` → `Diagnostics\<email>\images.csv` appears with sensible decisions. Restore the original settings file afterwards.

- [ ] **Step 5: Commit**

```powershell
git add src
git commit -m "feat: desktop OneNote backend wired into the app; no sign-in when local"
git push
```

---

### Task 10: Gated desktop smoke test, docs, checklist, release prep

**Files:**
- Create: `tests/SendToOneNote.Tests/DesktopIntegrationSmokeTests.cs`
- Modify: `README.md`, `docs/e2e-checklist.md`, `agents/knowledge/architecture.md`, `agents/knowledge/background.md`, `agents/rules/csharp.md`, `AGENTS.md`, `docs/superpowers/specs/2026-08-28-desktop-onenote-backend-design.md` (Status → Accepted)

**Interfaces:**
- Consumes: Tasks 1–9.
- Produces: the public face of v1.1 and the owner's verification path.

- [ ] **Step 1: Gated desktop smoke test**

`tests/SendToOneNote.Tests/DesktopIntegrationSmokeTests.cs`:

```csharp
using System.Xml.Linq;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class DesktopIntegrationSmokeTests
{
    private const string RedPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

    [SkippableFact]
    public async Task CreatesPageInDesktopOneNoteAndEmbedsImages()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1", "Set STN_INTEGRATION=1 to run against desktop OneNote.");
        using var worker = new StaComWorker();
        Skip.If(!await worker.RunAsync(() => DesktopOneNoteProbe.IsAvailable()), "Desktop OneNote not installed.");

        var backend = new DesktopOneNoteBackend(worker);
        var tree = await backend.GetTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections).FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        // One remote image (OneNote fetches + embeds) and one already-resolved image (base64-inlined by us).
        var xhtml = "<html><head><title>Desktop smoke (remote + inline)</title></head><body>" +
                    "<p>remote:</p><img src=\"https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png\"/>" +
                    "<p>inline:</p><img src=\"name:img0\" width=\"64\" height=\"64\"/></body></html>";
        var images = new List<ResolvedImage> { new("img0", "image/png", Convert.FromBase64String(RedPng), 1, 1) };

        var page = await backend.CreatePageAsync(scratch!.Id, xhtml, images);
        Assert.NotEmpty(page.Id);
        Assert.StartsWith("onenote:", page.ClientUrl);

        // Read back with binary data: both images must be stored embedded.
        var stored = await worker.RunAsync(() =>
        {
            var app = (IApplication)Activator.CreateInstance(Type.GetTypeFromProgID(OneNoteConstants.ProgId)!)!;
            app.GetPageContent(page.Id, out var xml, OneNoteConstants.PiBinaryData, OneNoteConstants.Xs2013);
            XNamespace one = OneNoteConstants.Namespace2013;
            return XDocument.Parse(xml).Descendants(one + "Image")
                .Select(i => i.Element(one + "Data")?.Value.Length ?? 0).ToList();
        });
        Assert.Equal(2, stored.Count);
        Assert.All(stored, len => Assert.True(len > 0));
    }
}
```

Run: `dotnet test --filter DesktopIntegrationSmokeTests` → SKIPPED without the env var. The owner runs `STN_INTEGRATION=1 dotnet test --filter DesktopIntegrationSmokeTests` (Git Bash syntax) and expects PASS with a new page in the test section.

- [ ] **Step 2: README**

Update `README.md`:
- Under the intro, add **"How it saves"**: desktop OneNote installed → instant local save, no sign-in, OneNote syncs; otherwise → Microsoft Graph with sign-in. Mention `settings.json` → `"Backend": "graph"` to force the cloud path (and `"desktop"` to force local).
- Install step 2 becomes conditional: "If you have desktop OneNote, there is no sign-in — just pick the folder."
- Privacy: add "With desktop OneNote, nothing is sent anywhere by this app; OneNote's own sync is the only network activity. Tracking pixels in emails are stripped before saving, so senders don't get a read receipt from your machine."
- Fidelity: note the desktop path renders through OneNote's own HTML importer (closer to the classic button); the sanitization caveat now applies to the Graph path only.
- Troubleshooting: `"ImageDiagnostics": true` writes `Diagnostics\<email>\images.csv` next to the drop folder; `"Backend": "desktop"` failing at startup means desktop OneNote isn't installed (the Store "OneNote for Windows 10" doesn't count).

- [ ] **Step 3: E2E checklist and agent docs**

Append to `docs/e2e-checklist.md`:

```markdown
- [ ] Fresh first-run on a machine with desktop OneNote shows no sign-in step
- [ ] Image-heavy newsletter saves within seconds and is visible in OneNote immediately (no "Page Not Yet Available")
- [ ] Saved page contains no tracking pixel (check page for stray 1x1 images)
- [ ] `"Backend": "graph"` forces the cloud path (tooltip says cloud; sign-in appears; save works)
- [ ] `"ImageDiagnostics": true` writes Diagnostics\<email>\images.csv with sensible decisions
- [ ] Picker shows local-only notebooks on the desktop path
```

`agents/knowledge/architecture.md`: add a **Backends** section (the seam, both implementations, `StaComWorker` rule, selection, per-backend recents, no cache on desktop) and add `Core/Desktop/` + `Core/Backends/` to the layout tree. `agents/knowledge/background.md`: add "Why the legacy button was fast" (COM to local cache) and the Click-to-Run type-library fact. `agents/rules/csharp.md`: add "All OneNote COM calls go through `StaComWorker`; never `dynamic`/`InvokeMember` on the OneNote RCW (Click-to-Run typelib is invisible)". `AGENTS.md`: quick-reference row for the desktop smoke test; stack-constraints bullet for the `Backend` setting; add the new knowledge files to the index if any were added.

- [ ] **Step 4: Mark the spec accepted, commit, push**

Change the spec's `Status:` line to `Accepted 2026-08-28 (implemented by docs/superpowers/plans/2026-08-28-desktop-onenote-backend.md)`.

```powershell
git add tests README.md docs agents AGENTS.md
git commit -m "docs: v1.1 README/checklist/agent docs; gated desktop OneNote smoke test"
git push
```

- [ ] **Step 5: Owner steps (manual, not automated)**

1. Run `docs/e2e-checklist.md` including the new items.
2. Delete `D:\Projects\TownBackyard\SendToOneNote-spike-com` (throwaway; its interface now lives in `OneNoteInterop.cs`).
3. Decide the release: tag `v0.2.0` (`git tag v0.2.0 && git push origin v0.2.0` — the release workflow publishes the zip). Tags are owner decisions.

---

## Task dependency order

1 → 2 → 3 → 4 (Core desktop stack) · 5 anytime after 1 · 6 → 7 → 8 (image pipeline, independent of 1–5) · 9 after 4, 5, 8 · 10 last. The owner can run the gated desktop smoke (Task 10 Step 1) as soon as Task 4 lands by copying that test in early — useful for catching COM surprises before the app wiring.

## GitHub issues mapping

Create one issue per task (Tasks 1–10), titled `v1.1 Task N: <name>`, labeled `v1.1` (create the label), body = the task's one-line scope + dependencies + link to this plan file. Plus one `manual` issue assigned to the owner: "v1.1 E2E run, spike cleanup, v0.2.0 tag" (Task 10 Step 5). Close each task issue with its commit SHA once its review is clean. Issue #20 (v2 picker polish) and #16–#19 remain the v2 backlog.
