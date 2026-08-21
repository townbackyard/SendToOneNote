using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class IntegrationSmokeTests
{
    // Minimal 1x1 PNG, reused per-image so each part carries its own byte[] instance.
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [SkippableFact]
    public async Task CreatesRealPageInScratchSection()
    {
        Skip.If(Environment.GetEnvironmentVariable("STN_INTEGRATION") != "1",
            "Set STN_INTEGRATION=1 to run against the real Graph API.");
        var tokens = new MsalTokenProvider(Path.Combine(Path.GetTempPath(), "stn-int"));
        var token = await tokens.GetAccessTokenAsync(interactiveAllowed: true);
        Assert.NotEmpty(token);

        var client = new OneNoteClient(tokens);
        var tree = await client.GetNotebookTreeAsync();
        var scratch = tree.Notebooks.SelectMany(n => n.Sections)
            .FirstOrDefault(s => s.Name == "SendToOneNote Test");
        Skip.If(scratch is null, "Create a section named 'SendToOneNote Test' first.");

        // Build the PagePlan via the real production pieces: 7 images through
        // PagePlanner.Plan forces at least one PATCH append batch, since only 5
        // binary parts fit in the initial create request (MaxBinaryPartsPerRequest).
        var imgTags = string.Concat(Enumerable.Range(0, 7).Select(i => $"<img src=\"name:img{i}\"/>"));
        var xhtml =
            $"<html><head><title>Integration smoke (images+append)</title></head><body>{imgTags}</body></html>";
        var images = Enumerable.Range(0, 7)
            .Select(i => new ResolvedImage($"img{i}", "image/png", (byte[])PngBytes.Clone()))
            .ToList();
        var plan = PagePlanner.Plan(xhtml, images);
        Assert.True(plan.Appends.Count >= 1,
            "7 images with a 5-part-per-request cap should force at least one append batch.");

        var page = await client.CreatePageAsync(scratch!.Id, plan);
        Assert.NotEmpty(page.Id);
    }
}
