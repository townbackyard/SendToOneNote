using System.Windows;

namespace SendToOneNote;

public partial class App : Application
{
    private TrayContext? _tray;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _tray = new TrayContext();
        _tray.Run();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _tray?.Dispose();
        base.OnExit(e);
    }
}
