using SendToOneNote.Core.Backends;

namespace SendToOneNote.Tests;

public class BackendSelectorTests
{
    [Fact]
    public void AutoPrefersDesktopWhenAvailable() =>
        Assert.Equal(BackendKind.Desktop, BackendSelector.Choose("auto", () => true).Kind);

    [Fact]
    public void AutoFallsBackToGraphWhenDesktopMissing() =>
        Assert.Equal(BackendKind.Graph, BackendSelector.Choose("auto", () => false).Kind);

    [Fact]
    public void GraphForcedSkipsDesktopWithoutProbing()
    {
        var probed = false;
        var choice = BackendSelector.Choose("graph", () => { probed = true; return true; });
        Assert.Equal(BackendKind.Graph, choice.Kind);
        Assert.False(probed);
    }

    [Fact]
    public void DesktopForcedButUnavailableThrows() =>
        Assert.Throws<InvalidOperationException>(() => BackendSelector.Choose("desktop", () => false));

    [Fact]
    public void NullOrUnknownSettingBehavesLikeAuto()
    {
        Assert.Equal(BackendKind.Desktop, BackendSelector.Choose(null, () => true).Kind);
        Assert.Equal(BackendKind.Graph, BackendSelector.Choose("banana", () => false).Kind);
    }
}
