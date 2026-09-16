using SendToOneNote.Core.Pages;

namespace SendToOneNote.Core.Desktop;

/// <summary>A file OneNote should embed: pathSource is read at import; preferredName is what the page shows.</summary>
public sealed record InsertedFile(string PathSource, string PreferredName);

/// <summary>
/// one:InsertedFile needs the bytes on disk. This writes each attachment to a per-save folder under
/// %TEMP%\SendToOneNote and deletes the folder on Dispose — OneNote copies the bytes into the notebook
/// during UpdatePageContent, so the files are not needed afterwards.
/// </summary>
public sealed class AttachmentTempFolder : IDisposable
{
    public static string DefaultRoot { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SendToOneNote");

    public string? Path { get; }
    public IReadOnlyList<InsertedFile> Files { get; }

    private AttachmentTempFolder(string? path, IReadOnlyList<InsertedFile> files) { Path = path; Files = files; }

    public static AttachmentTempFolder Create(IReadOnlyList<PlannedAttachment> attachments, string? root = null)
    {
        if (attachments.Count == 0) return new AttachmentTempFolder(null, []);

        var folder = System.IO.Path.Combine(root ?? DefaultRoot, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var files = new List<InsertedFile>();
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var a in attachments)
            {
                var name = SanitiseFileName(a.Attachment.FileName);
                var onDisk = name;
                for (var n = 2; !used.Add(onDisk); n++)
                    onDisk = System.IO.Path.GetFileNameWithoutExtension(name) + $" ({n})" + System.IO.Path.GetExtension(name);
                var path = System.IO.Path.Combine(folder, onDisk);
                File.WriteAllBytes(path, a.Attachment.Data);
                files.Add(new InsertedFile(path, a.Attachment.FileName));
            }
        }
        catch
        {
            TryDelete(folder);
            throw;
        }
        return new AttachmentTempFolder(folder, files);
    }

    public static string SanitiseFileName(string name)
    {
        var invalid = System.IO.Path.GetInvalidFileNameChars();
        var cleaned = new string(name.Select(c => invalid.Contains(c) ? '_' : c).ToArray()).Trim().TrimEnd('.');
        return cleaned.Length == 0 || cleaned.All(c => c == '.' || c == '_') ? "attachment" : cleaned;
    }

    public void Dispose() { if (Path is not null) TryDelete(Path); }

    // Best effort: a leftover temp folder must never turn a successful save into a failure.
    private static void TryDelete(string folder)
    {
        try { if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
