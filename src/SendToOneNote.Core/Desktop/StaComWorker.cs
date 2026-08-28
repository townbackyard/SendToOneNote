using System.Collections.Concurrent;

namespace SendToOneNote.Core.Desktop;

/// <summary>
/// One dedicated STA thread that owns every COM call. The RCW must be created on
/// this thread and never touched from another; RunAsync is the only entry point.
/// </summary>
public sealed class StaComWorker : IDisposable
{
    private readonly BlockingCollection<Action> _queue = new();
    private readonly Thread _thread;

    public StaComWorker(string threadName = "SendToOneNote COM")
    {
        _thread = new Thread(() =>
        {
            foreach (var job in _queue.GetConsumingEnumerable()) job();
        })
        { IsBackground = true, Name = threadName };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    public Task<T> RunAsync<T>(Func<T> work)
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try { tcs.SetResult(work()); }
            catch (Exception ex) { tcs.SetException(ex); }
        });
        return tcs.Task;
    }

    public Task RunAsync(Action work) => RunAsync(() => { work(); return true; });

    public void Dispose() => _queue.CompleteAdding();
}
