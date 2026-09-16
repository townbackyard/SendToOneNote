using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class AttachmentPlannerTests
{
    private static ParsedEmail Email(params EmailAttachment[] attachments) =>
        new("s", "a@b.c", "d@e.f", null, null, "<p>hi</p>", null, [], attachments, ["Fwd: nested"]);

    private static EmailAttachment File(string name, int bytes) =>
        new(name, "application/octet-stream", new byte[bytes]);

    [Fact]
    public void DisabledYieldsEmptyPlan()
    {
        var plan = AttachmentPlanner.Plan(Email(File("a.pdf", 10)), includeAttachments: false, maxAttachmentBytes: 1000);
        Assert.Same(AttachmentPlan.Empty, plan);
        Assert.Empty(plan.Embedded);
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void NumbersEmbeddedAttachmentsInEmailOrder()
    {
        var plan = AttachmentPlanner.Plan(Email(File("a.pdf", 10), File("b.docx", 20)), true, 1000);
        Assert.Equal(["att1", "att2"], plan.Embedded.Select(p => p.PartName));
        Assert.Equal(["a.pdf", "b.docx"], plan.Embedded.Select(p => p.Attachment.FileName));
        Assert.Empty(plan.Omitted);
    }

    [Fact]
    public void ExactlyAtCapIsEmbeddedOneByteOverIsOmitted()
    {
        var plan = AttachmentPlanner.Plan(Email(File("ok.pdf", 100), File("big.pdf", 101)), true, 100);
        Assert.Equal(["att1"], plan.Embedded.Select(p => p.PartName));
        var omitted = Assert.Single(plan.Omitted);
        Assert.Equal("big.pdf", omitted.FileName);
        Assert.Equal(101, omitted.Bytes);
        Assert.Equal(AttachmentPlanner.OverSizeLimit, omitted.Reason);
    }

    [Fact]
    public void NumberingSkipsOmittedFiles()
    {
        // att numbers count only embedded files, so parts are contiguous.
        var plan = AttachmentPlanner.Plan(Email(File("big.pdf", 500), File("small.pdf", 5)), true, 100);
        Assert.Equal(["att1"], plan.Embedded.Select(p => p.PartName));
        Assert.Equal("small.pdf", plan.Embedded[0].Attachment.FileName);
    }
}
