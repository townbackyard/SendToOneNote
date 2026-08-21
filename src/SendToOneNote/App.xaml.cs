using System.Threading;
using System.Windows;

namespace SendToOneNote;

public partial class App : Application
{
    private TrayContext? _tray;
    private Mutex? _mutex;
    private bool _ownsMutex;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _mutex = new Mutex(initiallyOwned: true, @"Local\SendToOneNote.SingleInstance", out var createdNew);
        _ownsMutex = createdNew;
        if (!createdNew)
        {
            MessageBox.Show("SendToOneNote is already running — check the system tray.", "SendToOneNote");
            Shutdown();
            return;
        }

        DispatcherUnhandledException += (_, ex) =>
        {
            TrayContext.TryLogStatic("Unhandled UI exception", ex.Exception);
            ex.Handled = true;
        };
        TaskScheduler.UnobservedTaskException += (_, ex) =>
        {
            TrayContext.TryLogStatic("Unobserved task exception", ex.Exception);
            ex.SetObserved();
        };

        try
        {
            _tray = new TrayContext();
            _tray.Run();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"SendToOneNote failed to start: {ex.Message}", "SendToOneNote");
            Shutdown();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _tray?.Dispose();
        if (_ownsMutex) _mutex?.ReleaseMutex();
        _mutex?.Dispose();
        base.OnExit(e);
    }
}
