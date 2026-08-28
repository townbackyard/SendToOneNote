using System.IO;
using System.Net.Http;
using System.Windows;
using SendToOneNote.Core.Auth;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;
using SendToOneNote.Core.Picker;
using SendToOneNote.Core.Storage;
using SendToOneNote.Logging;

namespace SendToOneNote;

public sealed class SavePipeline(
    JsonFileStore store, IOneNoteBackend backend, FileLog log)
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

            NotebookTree tree;
            if (backend.Name == "desktop")
            {
                tree = await backend.GetTreeAsync();            // ~100 ms local; always fresh
            }
            else
            {
                var cached = store.LoadTreeCache();
                tree = cached ?? await RefreshTreeAsync();
                if (cached is not null)                        // cache hit: refresh for next time
                    _ = RefreshTreeAsync().ContinueWith(t =>
                        log.Error("Background notebook refresh failed", t.Exception),
                        TaskContinuationOptions.OnlyOnFaulted);
            }

            var settings = store.LoadSettings();
            var recents = backend.Name == "desktop" ? settings.RecentDesktopSectionIds : settings.RecentSectionIds;
            var vm = new SectionPickerViewModel(tree, recents);

            PickerItem? pick = null;
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var win = new PickerWindow(vm, email.Subject);
                if (win.ShowDialog() == true) pick = win.Selected;
            });
            if (pick is null) { log.Info($"Cancelled: {path}"); return; }

            var xhtml = PageXhtmlBuilder.Build(email);
            var resolution = await _images.ResolveWithReportAsync(xhtml, email.InlineImages);
            var page = await backend.CreatePageAsync(pick.SectionId, resolution.Xhtml, resolution.Images);
            if (settings.ImageDiagnostics)
            {
                try
                {
                    // Only the Graph path can drop images (the 3.5 MB request cap);
                    // the desktop backend embeds every part.
                    IReadOnlyList<string> droppedMinor = backend.Name == "graph"
                        ? PagePlanner.Plan(resolution.Xhtml, resolution.Images).DroppedPartNames
                        : [];
                    var folder = ImageDiagnosticsWriter.Write(Path.GetDirectoryName(path)!,
                        Path.GetFileNameWithoutExtension(path), resolution.Decisions, resolution.Images, droppedMinor, DateTime.Now);
                    log.Info($"Image diagnostics written to {folder}");
                }
                catch (Exception diagEx) { log.Error("Image diagnostics failed (save was fine)", diagEx); }
            }

            if (backend.Name == "desktop")
                settings.RecentDesktopSectionIds = SectionPickerViewModel.PushRecent(recents, pick.SectionId);
            else
                settings.RecentSectionIds = SectionPickerViewModel.PushRecent(recents, pick.SectionId);
            store.SaveSettings(settings);

            if (settings.DeleteOnSuccess) File.Delete(path);
            log.Info($"Saved '{email.Subject}' to {pick.SectionName}");
            // A broken notifier must never convert a completed save into the
            // failure path — the page exists and the .eml is already handled.
            try { Saved?.Invoke($"Saved to {pick.SectionName}", page.ClientUrl); }
            catch (Exception notifyEx) { log.Error("Success notification failed (the email WAS saved)", notifyEx); }
        }
        catch (Exception ex)
        {
            log.Error($"Failed for {path}", ex);
            var reason = ex switch
            {
                EmlParseException => "That file isn't a readable email.",
                AuthRequiredException => "Sign-in required — open SendToOneNote from the tray.",
                OneNoteApiException o => $"OneNote API error {o.StatusCode}.",
                HttpRequestException => "You appear to be offline. The email was moved to the Failed folder — drag it back into the drop folder to retry.",
                DesktopOneNoteException d => $"Desktop OneNote: {d.Message}",
                _ => "Unexpected error — see log."
            };
            MoveToFailed(path, reason, ex);
            try { Failed?.Invoke(reason); }
            catch (Exception notifyEx) { log.Error("Failure notification failed", notifyEx); }
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task<NotebookTree> RefreshTreeAsync()
    {
        var tree = await backend.GetTreeAsync();
        store.SaveTreeCache(tree);
        return tree;
    }

    private void MoveToFailed(string path, string reason, Exception ex)
    {
        try
        {
            // Already gone (e.g. deleted after a successful save whose notification
            // then failed, or a duplicate event): nothing to move, nothing to explain.
            if (!File.Exists(path)) return;
            var failed = Path.Combine(Path.GetDirectoryName(path)!, "Failed");
            Directory.CreateDirectory(failed);
            var dest = Path.Combine(failed, Path.GetFileName(path));
            File.Move(path, dest, overwrite: true);
            // Sidecar so the user can see WHY without hunting for the app log.
            File.WriteAllText(dest + ".error.txt",
                $"""
                {DateTime.Now:yyyy-MM-dd HH:mm:ss}  SendToOneNote could not save this email.

                {reason}

                To retry: drag the .eml file back into the drop folder.

                Technical detail:
                {ex}

                Full log: %APPDATA%\SendToOneNote\logs
                """);
        }
        catch (Exception moveEx) { log.Error($"Could not move {path} to Failed", moveEx); }
    }
}
