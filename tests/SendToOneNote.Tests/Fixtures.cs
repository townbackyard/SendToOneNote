namespace SendToOneNote.Tests;

public static class Fixtures
{
    public static string Dir { get; } = FindDir();

    public static Stream Open(string name) =>
        File.OpenRead(Path.Combine(Dir, name));

    private static string FindDir()
    {
        var d = AppContext.BaseDirectory;
        while (d is not null && !Directory.Exists(Path.Combine(d, "fixtures", "synthetic")))
            d = Path.GetDirectoryName(d);
        return Path.Combine(d ?? throw new DirectoryNotFoundException("fixtures/synthetic not found"),
            "fixtures", "synthetic");
    }
}

public class FixtureTests
{
    [Theory]
    [InlineData("html-remote-images.eml")]
    [InlineData("plain-text-receipt.eml")]
    [InlineData("inline-cid-image.eml")]
    [InlineData("malformed.eml")]
    public void FixtureExists(string name) =>
        Assert.True(File.Exists(Path.Combine(Fixtures.Dir, name)));
}
