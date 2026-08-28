using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Storage;

namespace SendToOneNote.Tests;

public class StorageTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-tests-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void SettingsRoundTrip()
    {
        var store = new JsonFileStore(_dir);
        var s = store.LoadSettings();          // defaults on first run
        Assert.True(s.DeleteOnSuccess);
        s.DropFolder = @"C:\Drop";
        s.RecentSectionIds.Add("sec1");
        store.SaveSettings(s);
        var s2 = new JsonFileStore(_dir).LoadSettings();
        Assert.Equal(@"C:\Drop", s2.DropFolder);
        Assert.Equal(["sec1"], s2.RecentSectionIds);
    }

    [Fact]
    public void CorruptSettingsFallBackToDefaults()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(Path.Combine(_dir, "settings.json"), "{not json!!");
        var s = new JsonFileStore(_dir).LoadSettings();
        Assert.Null(s.DropFolder);
        Assert.True(s.DeleteOnSuccess);
    }

    [Fact]
    public void TreeCacheRoundTrip()
    {
        var store = new JsonFileStore(_dir);
        Assert.Null(store.LoadTreeCache());
        var tree = new NotebookTree(
            [new NotebookNode("n1", "General",
                [new SectionNode("s1", "Inbox")],
                [new GroupNode("g1", "Taxes", [new SectionNode("s2", "Taxes 2026")], [])])],
            DateTimeOffset.UtcNow);
        store.SaveTreeCache(tree);
        var loaded = new JsonFileStore(_dir).LoadTreeCache();
        Assert.NotNull(loaded);
        Assert.Equal("Taxes 2026", loaded!.Notebooks[0].Groups[0].Sections[0].Name);
    }

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
}
