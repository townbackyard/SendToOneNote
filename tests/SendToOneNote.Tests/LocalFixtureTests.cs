using SendToOneNote.Core.Email;

namespace SendToOneNote.Tests;

public class LocalFixtureTests
{
    public static IEnumerable<object?[]> LocalEmls()
    {
        var dir = Path.Combine(Path.GetDirectoryName(Fixtures.Dir)!, "local");
        var files = Directory.Exists(dir) ? Directory.GetFiles(dir, "*.eml") : [];
        if (files.Length == 0) { yield return new object?[] { null }; yield break; }
        foreach (var f in files) yield return new object?[] { f };
    }

    [SkippableTheory]
    [MemberData(nameof(LocalEmls))]
    public void ParsesRealEmailsWithoutThrowing(string? path)
    {
        Skip.If(path is null, "fixtures/local not present (CI or fresh clone)");
        using var s = File.OpenRead(path!);
        var e = EmlParser.Parse(s);
        Assert.False(string.IsNullOrWhiteSpace(e.Subject));
        Assert.True(e.HtmlBody is not null || e.TextBody is not null);
    }
}
