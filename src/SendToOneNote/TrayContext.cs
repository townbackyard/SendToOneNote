using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media;
using H.NotifyIcon;
using H.NotifyIcon.Core;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Storage;
using SendToOneNote.Core.Watch;
using SendToOneNote.Logging;

namespace SendToOneNote;

public sealed class TrayContext : IDisposable
{
    private readonly JsonFileStore _store = new();
    private readonly FileLog _log;
    private TaskbarIcon? _icon;
    private DropFolderWatcher? _watcher;
    private MsalTokenProvider? _tokens;
    private SavePipeline? _pipeline;
    private string? _lastUrl;

    public TrayContext() => _log = new FileLog(Path.Combine(_store.RootDir, "logs"));

    public void Run()
    {
        var settings = _store.LoadSettings();
        _tokens = new MsalTokenProvider(_store.RootDir, settings.ClientIdOverride);

        if (settings.DropFolder is null)
        {
            var first = new FirstRunWindow(_tokens, settings);
            first.ShowDialog();
            if (!first.Completed) { Application.Current.Shutdown(); return; }
            _store.SaveSettings(settings);
        }

        var client = new OneNoteClient(_tokens);
        _pipeline = new SavePipeline(_store, client, _log);
        _pipeline.Saved += (msg, url) => Notify("SendToOneNote", msg, url, NotificationIcon.Info);
        _pipeline.Failed += msg => Notify("SendToOneNote — failed", msg, null, NotificationIcon.Error);

        _watcher = new DropFolderWatcher(settings.DropFolder!);
        _watcher.EmlReady += p => _ = _pipeline.HandleEmlAsync(p);
        _watcher.NonEmlIgnored += p => _log.Info($"Ignored non-eml: {p}");
        _watcher.WatchError += msg => _log.Error(msg);
        _watcher.Start();

        _icon = new TaskbarIcon
        {
            ToolTipText = "SendToOneNote",
            IconSource = new GeneratedIconSource
            {
                Text = "N",
                Background = Brushes.SteelBlue,
                Foreground = Brushes.White
            }
        };
        // Subscribe once — the handler always opens whatever page the most recent
        // notification pointed at (null means the last toast carried no page URL).
        _icon.TrayBalloonTipClicked += (_, _) =>
        {
            if (_lastUrl is { } url)
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
        };

        var menu = new System.Windows.Controls.ContextMenu();
        AddItem(menu, "Open drop folder", () =>
            Process.Start(new ProcessStartInfo("explorer.exe", $"\"{settings.DropFolder}\"") { UseShellExecute = true }));
        AddItem(menu, "Sign in again", async () =>
        {
            try { await _tokens.GetAccessTokenAsync(interactiveAllowed: true); }
            catch (Exception ex) { _log.Error("Interactive sign-in failed", ex); }
        });
        AddItem(menu, "Exit", () => Application.Current.Shutdown());
        _icon.ContextMenu = menu;
        _log.Info("Started");
    }

    private static void AddItem(System.Windows.Controls.ContextMenu menu, string header, Action act)
    {
        var mi = new System.Windows.Controls.MenuItem { Header = header };
        mi.Click += (_, _) => act();
        menu.Items.Add(mi);
    }

    private void Notify(string title, string message, string? url, NotificationIcon icon)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            _lastUrl = url;
            _icon?.ShowNotification(title, message, icon);
        });
    }

    public void Dispose()
    {
        _icon?.Dispose();
        _watcher?.Dispose();
        _tokens = null;
    }
}
