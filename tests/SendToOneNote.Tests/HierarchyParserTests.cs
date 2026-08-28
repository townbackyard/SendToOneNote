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
