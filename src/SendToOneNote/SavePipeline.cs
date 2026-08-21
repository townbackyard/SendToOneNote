using System.IO;
using System.Net.Http;
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

    // Serializes the whole pipeline so concurrent multi-drops queue in arrival order
    // instead of racing on the picker dialog and settings.json read-modify-write.
    private readonly SemaphoreSlim _gate = new(1, 1);

    public event Action<string, string?>? Saved;   // (message, onenoteClientUrl)
    public event Action<string>? Failed;           // message

    public async Task HandleEmlAsync(string path)
    {
        // A late-duplicate watcher event for an already-processed (and deleted) file
        // must be a silent no-op, not an error toast.
        if (!File.Exists(path)) return;

        await _gate.WaitAsync();
        try
        {
            // The file may have been processed while this call was queued behind the gate.
            if (!File.Exists(path)) return;

            ParsedEmail email;
            await using (var s = File.OpenRead(path))
                email = EmlParser.Parse(s);

            var tree = store.LoadTreeCache();
            if (tree is null)
            {
                tree = await RefreshTreeAsync();
            }
            else
            {
                // Background refresh for next time — logged instead of silently swallowed.
                _ = RefreshTreeAsync().ContinueWith(t =>
                    log.Error("Background notebook refresh failed", t.Exception),
                    TaskContinuationOptions.OnlyOnFaulted);
            }

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
                HttpRequestException => "You appear to be offline. The email was moved to the Failed folder — drag it back into the drop folder to retry.",
                _ => "Unexpected error — see log."
            });
        }
        finally
        {
            _gate.Release();
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
