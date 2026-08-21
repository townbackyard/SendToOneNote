using System.Collections.Concurrent;

namespace SendToOneNote.Core.Watch;

public sealed class DropFolderWatcher : IDisposable
{
    private readonly FileSystemWatcher _fsw;
    private readonly ConcurrentDictionary<string, byte> _inFlight = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, byte> _ignoredOnce = new(StringComparer.OrdinalIgnoreCase);
    private readonly TimeSpan _readinessTimeout;

    public event Action<string>? EmlReady;
    public event Action<string>? NonEmlIgnored;
    public event Action<string>? WatchError;

    public DropFolderWatcher(string folder, TimeSpan? readinessTimeout = null)
    {
        _readinessTimeout = readinessTimeout ?? TimeSpan.FromSeconds(30);
        Directory.CreateDirectory(folder);
        _fsw = new FileSystemWatcher(folder) { IncludeSubdirectories = false, InternalBufferSize = 65536 };
        _fsw.Created += (_, e) => Handle(e.FullPath);
        _fsw.Renamed += (_, e) => Handle(e.FullPath);
        _fsw.Error += (_, e) => WatchError?.Invoke($"File watcher error: {e.GetException()?.Message}");
    }

    public void Start() => _fsw.EnableRaisingEvents = true;

    private void Handle(string path)
    {
        if (Directory.Exists(path)) return;
        if (!path.EndsWith(".eml", StringComparison.OrdinalIgnoreCase))
        {
            if (_ignoredOnce.Count > 256) _ignoredOnce.Clear();
            if (_ignoredOnce.TryAdd(path, 0)) NonEmlIgnored?.Invoke(path);
            return;
        }
        if (!_inFlight.TryAdd(path, 0)) return;
        _ = Task.Run(async () =>
        {
            try
            {
                if (await FileReadiness.WaitUntilUnlockedAsync(path, _readinessTimeout))
                    EmlReady?.Invoke(path);
                else
                    WatchError?.Invoke(
                        $"Timed out waiting for {Path.GetFileName(path)} to finish copying — it was left in the drop folder.");
            }
            catch (Exception ex) { WatchError?.Invoke($"{path}: {ex.Message}"); }
            finally { _inFlight.TryRemove(path, out _); }
        });
    }

    public void Dispose() => _fsw.Dispose();
}
