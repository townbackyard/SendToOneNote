using System.IO;

namespace SendToOneNote.Logging;

public sealed class FileLog(string dir)
{
    private readonly object _gate = new();
    private bool _pruned;

    public void Info(string msg) => Write("INFO", msg);
    public void Error(string msg, Exception? ex = null) =>
        Write("ERROR", ex is null ? msg : $"{msg} :: {ex}");

    private void Write(string level, string msg)
    {
        lock (_gate)
        {
            Directory.CreateDirectory(dir);
            PruneOldLogsOnce();
            File.AppendAllText(
                Path.Combine(dir, $"stn-{DateTime.Now:yyyyMMdd}.log"),
                $"{DateTime.Now:HH:mm:ss} [{level}] {msg}{Environment.NewLine}");
        }
    }

    // Best-effort startup pruning of old log files — runs once per FileLog instance,
    // on the first write (after the log directory is guaranteed to exist).
    private void PruneOldLogsOnce()
    {
        if (_pruned) return;
        _pruned = true;
        try
        {
            var cutoff = DateTime.Now.AddDays(-30);
            foreach (var file in Directory.EnumerateFiles(dir, "stn-*.log"))
            {
                if (File.GetLastWriteTime(file) < cutoff)
                    File.Delete(file);
            }
        }
        catch { /* logging must never throw */ }
    }
}
