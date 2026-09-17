using System.Drawing;
using System.Runtime.InteropServices;

namespace SendToOneNote;

/// <summary>
/// The app's icon, embedded per build configuration by the project file (purple for Release,
/// orange with a "D" badge for Debug) so a running debug build is never mistaken for the
/// installed one. The same .ico is the exe's application icon; regenerate both with
/// <c>tools/make-icons.ps1</c>.
/// </summary>
internal static class AppIcon
{
    private const string ResourceName = "SendToOneNote.app.ico";
    private const int SmCxSmIcon = 49; // GetSystemMetrics index: small-icon width in pixels at the current DPI

#if DEBUG
    public const string TooltipSuffix = " (debug)";
#else
    public const string TooltipSuffix = "";
#endif

    /// <summary>
    /// Loads the frame that matches the notification area's size (16 px at 100%, 24 px at 150%, …)
    /// so the shell shows a frame drawn for that size instead of scaling a large one down.
    /// The caller owns the returned icon.
    /// </summary>
    public static Icon LoadForTray()
    {
        using var stream = typeof(AppIcon).Assembly.GetManifestResourceStream(ResourceName)
            ?? throw new InvalidOperationException($"Embedded icon '{ResourceName}' is missing from the build.");
        var size = GetSystemMetrics(SmCxSmIcon);
        if (size <= 0) size = 16;
        return new Icon(stream, size, size);
    }

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);
}
