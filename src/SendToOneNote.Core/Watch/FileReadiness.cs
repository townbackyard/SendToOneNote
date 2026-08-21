namespace SendToOneNote.Core.Watch;

public static class FileReadiness
{
    public static async Task<bool> WaitUntilUnlockedAsync(string path, TimeSpan timeout,
        CancellationToken ct = default)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            try
            {
                using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None);
                return true;
            }
            catch (FileNotFoundException) { return false; }
            catch (IOException) { await Task.Delay(250, ct); }
        }
        return false;
    }
}
