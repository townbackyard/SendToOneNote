using System.Text.Json;
using SendToOneNote.Core.OneNote;

namespace SendToOneNote.Core.Storage;

public sealed class JsonFileStore
{
    private static readonly JsonSerializerOptions Opts = new() { WriteIndented = true };
    public string RootDir { get; }

    public JsonFileStore(string? rootDir = null)
    {
        RootDir = rootDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SendToOneNote");
        Directory.CreateDirectory(RootDir);
    }

    public AppSettings LoadSettings() => Load<AppSettings>("settings.json") ?? new AppSettings();
    public void SaveSettings(AppSettings s) => Save("settings.json", s);
    public NotebookTree? LoadTreeCache() => Load<NotebookTree>("cache.json");
    public void SaveTreeCache(NotebookTree t) => Save("cache.json", t);

    private T? Load<T>(string file) where T : class
    {
        var path = Path.Combine(RootDir, file);
        if (!File.Exists(path)) return null;
        try { return JsonSerializer.Deserialize<T>(File.ReadAllText(path), Opts); }
        catch (JsonException) { return null; }
    }

    private void Save<T>(string file, T value)
    {
        var path = Path.Combine(RootDir, file);
        var tmp = path + ".tmp";
        File.WriteAllText(tmp, JsonSerializer.Serialize(value, Opts));
        File.Move(tmp, path, overwrite: true);
    }
}
