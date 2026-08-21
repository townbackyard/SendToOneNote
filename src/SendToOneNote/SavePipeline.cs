using System.IO;
using System.Windows;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;
using SendToOneNote.Core.Picker;
using SendToOneNote.Core.Storage;
using SendToOneNote.Logging;

namespace SendToOneNote;

public sealed class SavePipeline(
    JsonFileStore store, OneNoteClient client, FileLog log)
{
    // One HttpClient-backed resolver reused across emails — avoids the
    // HttpClient-per-call anti-pattern of constructing a new ImageResolver per email.
    private readonly ImageResolver _images = new();

    public event Action<string, string?>? Saved;   // (message, onenoteClientUrl)
    public event Action<string>? Failed;           // message

    public async Task HandleEmlAsync(string path)
    {
        // A late-duplicate watcher event for an already-processed (and deleted) file
        // must be a silent no-op, not an error toast.
        if (!File.Exists(path)) return;

        try
        {
            ParsedEmail email;
            await using (var s = File.OpenRead(path))
                email = EmlParser.Parse(s);

            var tree = store.LoadTreeCache() ?? await RefreshTreeAsync();
            // Background refresh for next time — logged instead of silently swallowed.
            _ = RefreshTreeAsync().ContinueWith(t =>
                log.Error("Background notebook refresh failed", t.Exception),
                TaskContinuationOptions.OnlyOnFaulted);

            var settings = store.LoadSettings();
            var vm = new SectionPickerViewModel(tree, settings.RecentSectionIds);

            PickerItem? pick = null;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var win = new PickerWindow(vm, email.Subject);
                if (win.ShowDialog() == true) pick = win.Selected;
            });
            if (pick is null) { log.Info($"Cancelled: {path}"); return; }

            var xhtml = PageXhtmlBuilder.Build(email);
            var (normalized, images) = await _images.ResolveAsync(xhtml, email.InlineImages);
            var plan = PagePlanner.Plan(normalized, images);
            var page = await client.CreatePageAsync(pick.SectionId, plan);

            settings.RecentSectionIds =
                SectionPickerViewModel.PushRecent(settings.RecentSectionIds, pick.SectionId);
            store.SaveSettings(settings);

            if (settings.DeleteOnSuccess) File.Delete(path);
            log.Info($"Saved '{email.Subject}' to {pick.SectionName}");
            Saved?.Invoke($"Saved to {pick.SectionName}", page.ClientUrl);
        }
        catch (Exception ex)
        {
            log.Error($"Failed for {path}", ex);
            MoveToFailed(path);
            Failed?.Invoke(ex switch
            {
                EmlParseException => "That file isn't a readable email.",
                AuthRequiredException => "Sign-in required — open SendToOneNote from the tray.",
                OneNoteApiException o => $"OneNote API error {o.StatusCode}.",
                _ => "Unexpected error — see log."
            });
        }
    }

    private async Task<NotebookTree> RefreshTreeAsync()
    {
        var tree = await client.GetNotebookTreeAsync();
        store.SaveTreeCache(tree);
        return tree;
    }

    private void MoveToFailed(string path)
    {
        try
        {
            var failed = Path.Combine(Path.GetDirectoryName(path)!, "Failed");
            Directory.CreateDirectory(failed);
            File.Move(path, Path.Combine(failed, Path.GetFileName(path)), overwrite: true);
        }
        catch (Exception moveEx) { log.Error($"Could not move {path} to Failed", moveEx); }
    }
}
