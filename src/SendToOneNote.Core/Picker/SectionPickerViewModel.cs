using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Core.Picker;

public sealed record PickerItem(string SectionId, string SectionName, string Path);

public sealed class SectionPickerViewModel
{
    private readonly IReadOnlyList<string> _recents;
    public IReadOnlyList<PickerItem> AllSections { get; }

    public SectionPickerViewModel(NotebookTree tree, IReadOnlyList<string> recentSectionIds)
    {
        _recents = recentSectionIds;
        var items = new List<PickerItem>();
        foreach (var nb in tree.Notebooks)
        {
            items.AddRange(nb.Sections.Select(s => new PickerItem(s.Id, s.Name, nb.Name)));
            foreach (var g in nb.Groups) Walk(g, nb.Name, items);
        }
        AllSections = items;

        static void Walk(GroupNode g, string path, List<PickerItem> items)
        {
            var p = $"{path} » {g.Name}";
            items.AddRange(g.Sections.Select(s => new PickerItem(s.Id, s.Name, p)));
            foreach (var child in g.Groups) Walk(child, p, items);
        }
    }

    public IReadOnlyList<PickerItem> Filter(string query)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            var recents = _recents
                .Select(id => AllSections.FirstOrDefault(s => s.SectionId == id))
                .Where(s => s is not null).Cast<PickerItem>().ToList();
            return [.. recents, .. AllSections.Where(s => !_recents.Contains(s.SectionId))];
        }
        var q = query.Trim();
        return AllSections.Where(s =>
            s.SectionName.Contains(q, StringComparison.OrdinalIgnoreCase) ||
            $"{s.Path} » {s.SectionName}".Contains(q, StringComparison.OrdinalIgnoreCase) ||
            s.Path.Contains(q, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    public static List<string> PushRecent(List<string> recents, string sectionId, int cap = 10)
    {
        var r = new List<string> { sectionId };
        r.AddRange(recents.Where(x => x != sectionId));
        return r.Take(cap).ToList();
    }
}
