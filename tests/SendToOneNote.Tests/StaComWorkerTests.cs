using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class StaComWorkerTests
{
    [Fact]
    public async Task RunsWorkOnAnStaThread()
    {
        using var w = new StaComWorker();
        var state = await w.RunAsync(() => Thread.CurrentThread.GetApartmentState());
        Assert.Equal(ApartmentState.STA, state);
    }

    [Fact]
    public async Task AllWorkRunsOnTheSameThread()
    {
        using var w = new StaComWorker();
        var a = await w.RunAsync(() => Environment.CurrentManagedThreadId);
        var b = await w.RunAsync(() => Environment.CurrentManagedThreadId);
        Assert.Equal(a, b);
        Assert.NotEqual(Environment.CurrentManagedThreadId, a);
    }

    [Fact]
    public async Task ExceptionsPropagateToTheCaller()
    {
        using var w = new StaComWorker();
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            w.RunAsync<int>(() => throw new InvalidOperationException("boom")));
        Assert.Equal(7, await w.RunAsync(() => 7)); // worker survives a failed job
    }
}
