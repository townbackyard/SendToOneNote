using SendToOneNote.Core.Watch;

namespace SendToOneNote.Tests;

public class DropFolderWatcherTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "stn-watch-" + Guid.NewGuid());
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public async Task RaisesEmlReadyOnceForDroppedFile()
    {
        using var w = new DropFolderWatcher(_dir);
        var hits = new List<string>();
        var signal = new TaskCompletionSource();
        w.EmlReady += p => { lock (hits) hits.Add(p); signal.TrySetResult(); };
        w.Start();

        var path = Path.Combine(_dir, "mail.eml");
        await File.WriteAllTextAsync(path, "From: a@b.c\r\n\r\nhi");
        await Task.WhenAny(signal.Task, Task.Delay(5000));
        await Task.Delay(500); // absorb any duplicate events

        Assert.Equal([path], hits);
    }

    [Fact]
    public async Task IgnoresNonEmlWithNotice()
    {
        using var w = new DropFolderWatcher(_dir);
        string? ignored = null;
        var signal = new TaskCompletionSource();
        w.NonEmlIgnored += p => { ignored = p; signal.TrySetResult(); };
        w.EmlReady += _ => throw new InvalidOperationException("must not fire");
        w.Start();

        await File.WriteAllTextAsync(Path.Combine(_dir, "note.txt"), "hi");
        await Task.WhenAny(signal.Task, Task.Delay(5000));
        Assert.EndsWith("note.txt", ignored);
    }

    [Fact]
    public async Task WaitsForLockedFileToBeReleased()
    {
        Directory.CreateDirectory(_dir);
        var path = Path.Combine(_dir, "locked.eml");
        await File.WriteAllTextAsync(path, "x");
        Task<bool> wait;
        using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            wait = FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(10));
            await Task.Delay(600);
            Assert.False(wait.IsCompleted);
        }
        Assert.True(await wait);
    }

    [Fact]
    public async Task TimesOutOnPermanentLock()
    {
        Directory.CreateDirectory(_dir);
        var path = Path.Combine(_dir, "stuck.eml");
        await File.WriteAllTextAsync(path, "x");
        using var _ = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None);
        Assert.False(await FileReadiness.WaitUntilUnlockedAsync(path, TimeSpan.FromSeconds(1)));
    }

    [Fact]
    public async Task RaisesWatchErrorOnReadinessTimeout()
    {
        // A 1s readiness timeout keeps this test fast instead of waiting on the 30s default.
        using var w = new DropFolderWatcher(_dir, TimeSpan.FromSeconds(1));
        string? error = null;
        var emlReadyFired = false;
        var signal = new TaskCompletionSource();
        w.WatchError += msg => { error = msg; signal.TrySetResult(); };
        w.EmlReady += _ => emlReadyFired = true;
        w.Start();

        var path = Path.Combine(_dir, "stuck.eml");
        using (File.Open(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            await Task.WhenAny(signal.Task, Task.Delay(10_000));
        }

        Assert.NotNull(error);
        Assert.Contains("stuck.eml", error);
        Assert.False(emlReadyFired);
    }

    [Fact]
    public async Task ThrowingSubscriberRaisesWatchErrorAndDoesNotCrash()
    {
        using var w = new DropFolderWatcher(_dir);
        string? error = null;
        var signal = new TaskCompletionSource();
        w.EmlReady += _ => throw new InvalidOperationException("subscriber boom");
        w.WatchError += msg => { error = msg; signal.TrySetResult(); };
        w.Start();

        await File.WriteAllTextAsync(Path.Combine(_dir, "boom.eml"), "From: a@b.c\r\n\r\nhi");
        await Task.WhenAny(signal.Task, Task.Delay(5000));
        Assert.Contains("subscriber boom", error);
    }
}
