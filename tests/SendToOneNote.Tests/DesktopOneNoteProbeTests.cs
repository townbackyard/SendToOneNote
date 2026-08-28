using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

public class DesktopOneNoteProbeTests
{
    [Fact]
    public async Task UnknownProgIdIsNotAvailable()
    {
        using var w = new StaComWorker();
        Assert.False(await w.RunAsync(() => DesktopOneNoteProbe.IsAvailable("SendToOneNote.NoSuchProgId.Test")));
    }

    [Fact]
    public async Task RealProbeNeverThrows()
    {
        // CI runners have no OneNote (false); the owner's machine has it (true). Either is fine.
        using var w = new StaComWorker();
        var ex = await Record.ExceptionAsync(() => w.RunAsync(() => DesktopOneNoteProbe.IsAvailable()));
        Assert.Null(ex);
    }
}
