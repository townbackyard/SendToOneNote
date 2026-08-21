using SendToOneNote.Core.Email;

namespace SendToOneNote.Tests;

public class LocalFixtureTests
{
    public static IEnumerable<object[]> LocalEmls()
    {
        var dir = Path.Combine(Path.GetDirectoryName(Fixtures.Dir)!, "local");
        if (!Directory.Exists(dir)) yield break;
        foreach (var f in Directory.GetFiles(dir, "*.eml"))
            yield return new object[] { f };
    }

    [Theory]
    [MemberData(nameof(LocalEmls))]
    public void ParsesRealEmailsWithoutThrowing(string path)
    {
        using var s = File.OpenRead(path);
        var e = EmlParser.Parse(s);
        Assert.False(string.IsNullOrWhiteSpace(e.Subject));
        Assert.True(e.HtmlBody is not null || e.TextBody is not null);
    }
}
