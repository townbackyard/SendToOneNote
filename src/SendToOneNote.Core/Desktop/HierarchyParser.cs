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
