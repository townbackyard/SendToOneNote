using System.Xml.Linq;
using SendToOneNote.Core.Backends;
using SendToOneNote.Core.Desktop;
using SendToOneNote.Core.Email;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class DesktopIntegrationSmokeTests
{
    private const string RedPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";

    [SkippableFact]
    public async Task CreatesPageInDesktopOneNoteAndEmbedsImages()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1", "Set STN_INTEGRATION=1 to run against desktop OneNote.");
        using var worker = new StaComWorker();
        Skip.If(!await worker.RunAsync(() => DesktopOneNoteProbe.IsAvailable()), "Desktop OneNote not installed.");

        var backend = new DesktopOneNoteBackend(worker);
        var tree = await backend.GetTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections).FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        // One remote image (OneNote fetches + embeds) and one already-resolved image (base64-inlined by us).
        var xhtml = "<html><head><title>Desktop smoke (remote + inline)</title></head><body>" +
                    "<p>remote:</p><img src=\"https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_272x92dp.png\"/>" +
                    "<p>inline:</p><img src=\"name:img0\" width=\"64\" height=\"64\"/></body></html>";
        var images = new List<ResolvedImage> { new("img0", "image/png", Convert.FromBase64String(RedPng), 1, 1) };

        var page = await backend.CreatePageAsync(scratch!.Id, new PageContent(xhtml, images, []));
        Assert.NotEmpty(page.Id);
        Assert.StartsWith("onenote:", page.ClientUrl);

        // Read back with binary data: both images must be stored embedded.
        var stored = await worker.RunAsync(() =>
        {
            var app = (IApplication)Activator.CreateInstance(Type.GetTypeFromProgID(OneNoteConstants.ProgId)!)!;
            app.GetPageContent(page.Id, out var xml, OneNoteConstants.PiBinaryData, OneNoteConstants.Xs2013);
            XNamespace one = OneNoteConstants.Namespace2013;
            return XDocument.Parse(xml).Descendants(one + "Image")
                .Select(i => i.Element(one + "Data")?.Value.Length ?? 0).ToList();
        });
        Assert.Equal(2, stored.Count);
        Assert.All(stored, len => Assert.True(len > 0));
    }

    [SkippableFact]
    public async Task EmbedsPdfAttachmentAsInsertedFile()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1", "Set STN_INTEGRATION=1 to run against desktop OneNote.");
        using var worker = new StaComWorker();
        Skip.If(!await worker.RunAsync(() => DesktopOneNoteProbe.IsAvailable()), "Desktop OneNote not installed.");
        var backend = new DesktopOneNoteBackend(worker);
        var tree = await backend.GetTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections).FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        // The real production pieces, on the synthetic PDF fixture.
        var email = EmlParser.Parse(Fixtures.Open("pdf-attachment.eml"));
        var plan = AttachmentPlanner.Plan(email, true, 26_214_400);
        var resolution = await new ImageResolver().ResolveWithReportAsync(PageXhtmlBuilder.Build(email, plan), email.InlineImages);
        var content = new PageContent(resolution.Xhtml, resolution.Images, plan.Embedded);

        var page = await backend.CreatePageAsync(scratch!.Id, content);
        Assert.NotEmpty(page.Id);

        var insertedFiles = await worker.RunAsync(() =>
        {
            var app = (IApplication)Activator.CreateInstance(Type.GetTypeFromProgID(OneNoteConstants.ProgId)!)!;
            app.GetPageContent(page.Id, out var xml, OneNoteConstants.PiBinaryData, OneNoteConstants.Xs2013);
            XNamespace one = OneNoteConstants.Namespace2013;
            return XDocument.Parse(xml).Descendants(one + "InsertedFile")
                .Select(f => f.Attribute("preferredName")?.Value).ToList();
        });
        Assert.Equal(["invoice.pdf"], insertedFiles);

        // Temp folder cleaned up (nothing of ours left under %TEMP%\SendToOneNote).
        if (Directory.Exists(AttachmentTempFolder.DefaultRoot))
            Assert.Empty(Directory.GetDirectories(AttachmentTempFolder.DefaultRoot));
    }
}
