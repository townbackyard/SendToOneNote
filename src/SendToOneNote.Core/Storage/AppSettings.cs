namespace SendToOneNote.Core.Storage;

public sealed class AppSettings
{
    public string? DropFolder { get; set; }
    public string? ClientIdOverride { get; set; }
    public bool DeleteOnSuccess { get; set; } = true;
    public List<string> RecentSectionIds { get; set; } = [];

    /// <summary>"auto" (try desktop OneNote first) | "desktop" (force COM) | "graph" (force cloud).</summary>
    public string Backend { get; set; } = "auto";
    /// <summary>Write <DropFolder>\Diagnostics\<email>\images.csv (+ images) per save.</summary>
    public bool ImageDiagnostics { get; set; }
    /// <summary>Recents for the desktop backend — section IDs differ from Graph's.</summary>
    public List<string> RecentDesktopSectionIds { get; set; } = [];
}
