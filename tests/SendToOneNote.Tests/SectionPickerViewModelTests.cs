using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Picker;

namespace SendToOneNote.Tests;

public class SectionPickerViewModelTests
{
    private static NotebookTree Tree() => new(
        [
            new NotebookNode("n1", "General",
                [new SectionNode("s1", "Inbox"), new SectionNode("s2", "Quick Notes")], []),
            new NotebookNode("n2", "Projects",
                [new SectionNode("s3", "AirBnb")],
                [new GroupNode("g1", "Taxes", [new SectionNode("s4", "Taxes 2026")], [])])
        ], DateTimeOffset.UtcNow);

    [Fact]
    public void FlattensWithPaths()
    {
        var vm = new SectionPickerViewModel(Tree(), []);
        Assert.Equal(4, vm.AllSections.Count);
        var taxes = vm.AllSections.Single(s => s.SectionId == "s4");
        Assert.Equal("Projects » Taxes", taxes.Path);
    }

    [Fact]
    public void EmptyQueryPutsRecentsFirst()
    {
        var vm = new SectionPickerViewModel(Tree(), ["s4", "s1"]);
        var items = vm.Filter("");
        Assert.Equal("s4", items[0].SectionId);
        Assert.Equal("s1", items[1].SectionId);
        Assert.Equal(4, items.Count); // no duplicates
    }

    [Fact]
    public void QueryMatchesNameAndPathCaseInsensitive()
    {
        var vm = new SectionPickerViewModel(Tree(), []);
        Assert.Equal(["s4"], vm.Filter("taxes 2").Select(i => i.SectionId));
        Assert.Equal(["s4"], vm.Filter("projects » tax").Select(i => i.SectionId).Distinct());
        Assert.Contains("s3", vm.Filter("PROJECT").Select(i => i.SectionId)); // path match
    }

    [Fact]
    public void PushRecentDedupesAndCaps()
    {
        var r = SectionPickerViewModel.PushRecent(["a", "b"], "b");
        Assert.Equal(["b", "a"], r);
        var many = Enumerable.Range(0, 12).Select(i => $"s{i}").ToList();
        var capped = SectionPickerViewModel.PushRecent(many, "new");
        Assert.Equal(10, capped.Count);
        Assert.Equal("new", capped[0]);
    }
}
