using SendToOneNote.Core.Auth;
using SendToOneNote.Core.OneNote;
using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class IntegrationSmokeTests
{
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

        var page = await client.CreatePageAsync(scratch!.Id, new PagePlan(
            "<html><head><title>Integration smoke</title></head><body><p>hello</p></body></html>",
            [], []));
        Assert.NotEmpty(page.Id);
    }
}
