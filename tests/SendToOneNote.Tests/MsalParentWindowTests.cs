using SendToOneNote.Core.Auth;

namespace SendToOneNote.Tests;

public class MsalParentWindowTests
{
    [Fact]
    public void ConfiguredHandleWinsWhenNonZero()
    {
        var configured = new IntPtr(123);
        Assert.Equal(configured, MsalTokenProvider.ResolveParentWindow(configured));
    }

    [Fact]
    public void ZeroHandleResolvesToARealWindow()
    {
        // WAM refuses IntPtr.Zero; the resolver must fall back to the foreground
        // window or, failing that, the desktop window — never return zero.
        Assert.NotEqual(IntPtr.Zero, MsalTokenProvider.ResolveParentWindow(IntPtr.Zero));
    }
}
