using System.IO;

namespace SendToOneNote.Logging;

public sealed class FileLog(string dir)
{
    private readonly object _gate = new();

    public void Info(string msg) => Write("INFO", msg);
    public void Error(string msg, Exception? ex = null) =>
        Write("ERROR", ex is null ? msg : $"{msg} :: {ex}");

    private void Write(string level, string msg)
    {
        lock (_gate)
        {
            Directory.CreateDirectory(dir);
            File.AppendAllText(
                Path.Combine(dir, $"stn-{DateTime.Now:yyyyMMdd}.log"),
                $"{DateTime.Now:HH:mm:ss} [{level}] {msg}{Environment.NewLine}");
        }
    }
}
