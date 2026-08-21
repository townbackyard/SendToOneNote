using System.Collections.Concurrent;

namespace SendToOneNote.Core.Watch;

public sealed class DropFolderWatcher : IDisposable
{
    private readonly FileSystemWatcher _fsw;
    private readonly ConcurrentDictionary<string, byte> _inFlight = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, byte> _ignoredOnce = new(StringComparer.OrdinalIgnoreCase);

    public event Action<string>? EmlReady;
    public event Action<string>? NonEmlIgnored;

    public DropFolderWatcher(string folder)
    {
        Directory.CreateDirectory(folder);
        _fsw = new FileSystemWatcher(folder) { IncludeSubdirectories = false };
        _fsw.Created += (_, e) => Handle(e.FullPath);
        _fsw.Renamed += (_, e) => Handle(e.FullPath);
    }

    public void Start() => _fsw.EnableRaisingEvents = true;

    private void Handle(string path)
    {
        if (Directory.Exists(path)) return;
        if (!path.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
        {
            if (_ignoredOnce.TryAdd(path, 0)) NonEmlIgnored?.Invoke(path);
            return;
        }
        if (!_inFlight.TryAdd(path, 0)) return;
        _ = Task.Run(async () =>
        {
            try
            {
                if (await FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(30)))
                    EmlReady?.Invoke(path);
            }
            finally { _inFlight.TryRemove(path, out _); }
        });
    }

    public void Dispose() => _fsw.Dispose();
}
