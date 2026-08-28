using System.Runtime.InteropServices;

namespace SendToOneNote.Core.Desktop;

public static class DesktopOneNoteProbe
{
    /// <summary>
    /// True when desktop OneNote is installed AND exposes the 2013+ IApplication.
    /// Activates OneNote (launches it if not running). Call on the StaComWorker.
    /// </summary>
    public static bool IsAvailable(string progId = OneNoteConstants.ProgId)
    {
        var type = Type.GetTypeFromProgID(progId);
        if (type is null) return false;
        object? instance = null;
        try
        {
            instance = Activator.CreateInstance(type);
            return instance is IApplication; // QueryInterface for the declared IID
        }
        catch (Exception)
        {
            return false;
        }
        finally
        {
            if (instance is not null && Marshal.IsComObject(instance)) Marshal.ReleaseComObject(instance);
        }
    }
}
