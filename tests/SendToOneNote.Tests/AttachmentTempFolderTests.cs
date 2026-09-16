using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class AttachmentTempFolderTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "stn-tests", Guid.NewGuid().ToString("N"));

    private static PlannedAttachment Att(int n, string name, byte[]? data = null) =>
        new($"att{n}", new EmailAttachment(name, "application/octet-stream", data ?? [1, 2, 3]));

    [Fact]
    public void WritesEachFileAndDeletesTheFolderOnDispose()
    {
        string? folder;
        using (var temp = AttachmentTempFolder.Create([Att(1, "a.pdf", [9, 8]), Att(2, "b.txt")], _root))
        {
            folder = temp.Path;
            Assert.NotNull(folder);
            Assert.StartsWith(_root, folder);
            Assert.Equal(2, temp.Files.Count);
            Assert.Equal("a.pdf", temp.Files[0].PreferredName);
            Assert.Equal(new byte[] { 9, 8 }, File.ReadAllBytes(temp.Files[0].PathSource));
            Assert.Equal(Path.Combine(folder, "b.txt"), temp.Files[1].PathSource);
        }
        Assert.False(Directory.Exists(folder));
    }

    [Fact]
    public void NoAttachmentsCreatesNothing()
    {
        using var temp = AttachmentTempFolder.Create([], _root);
        Assert.Null(temp.Path);
        Assert.Empty(temp.Files);
        Assert.False(Directory.Exists(_root));
    }

    [Theory]
    [InlineData("in:va|id?.pdf", "in_va_id_.pdf")]
    [InlineData("   ", "attachment")]
    [InlineData("..", "attachment")]
    [InlineData("plain.docx", "plain.docx")]
    public void SanitisesNames(string input, string expected) =>
        Assert.Equal(expected, AttachmentTempFolder.SanitiseFileName(input));

    [Fact]
    public void DuplicateNamesGetNumericSuffixButKeepPreferredName()
    {
        using var temp = AttachmentTempFolder.Create([Att(1, "same.pdf"), Att(2, "same.pdf")], _root);
        Assert.Equal("same.pdf", Path.GetFileName(temp.Files[0].PathSource));
        Assert.Equal("same (2).pdf", Path.GetFileName(temp.Files[1].PathSource));
        Assert.Equal("same.pdf", temp.Files[1].PreferredName);
    }

    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }
}
