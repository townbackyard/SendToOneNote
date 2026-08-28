namespace SendToOneNote.Core.Backends;

public enum BackendKind { Desktop, Graph }

public sealed record BackendChoice(BackendKind Kind, string Reason);

public static class BackendSelector
{
    public static BackendChoice Choose(string? setting, Func<bool> desktopAvailable)
    {
        switch ((setting ?? "auto").Trim().ToLowerInvariant())
        {
            case "graph":
                return new(BackendKind.Graph, "Backend=graph (forced)");
            case "desktop":
                return desktopAvailable()
                    ? new(BackendKind.Desktop, "Backend=desktop (forced)")
                    : throw new InvalidOperationException(
                        "Backend is set to \"desktop\" but desktop OneNote is not available on this machine. " +
                        "Install desktop OneNote or set Backend to \"auto\" or \"graph\" in settings.json.");
            default:
                return desktopAvailable()
                    ? new(BackendKind.Desktop, "auto: desktop OneNote detected")
                    : new(BackendKind.Graph, "auto: desktop OneNote not available");
        }
    }
}
