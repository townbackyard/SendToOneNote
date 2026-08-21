namespace SendToOneNote.Core.Storage;

public sealed class AppSettings
{
    public string? DropFolder { get; set; }
    public string? ClientIdOverride { get; set; }
    public bool DeleteOnSuccess { get; set; } = true;
    public List<string> RecentSectionIds { get; set; } = [];
}
