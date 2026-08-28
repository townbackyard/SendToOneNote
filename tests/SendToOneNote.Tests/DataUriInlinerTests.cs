using SendToOneNote.Core.Pages;

namespace SendToOneNote.Tests;

public class DataUriInlinerTests
{
    [Fact]
    public void RewritesEachPartNameToADataUri()
    {
        var images = new List<ResolvedImage>
        {
            new("img0", "image/png", [1, 2, 3]),
            new("img1", "image/jpeg", [4, 5]),
        };
        var xhtml = "<img alt=\"a\" src=\"name:img0\" width=\"10\"/><img src=\"name:img1\"/>";
        var result = DataUriInliner.Inline(xhtml, images);
        Assert.Contains("src=\"data:image/png;base64,AQID\"", result);
        Assert.Contains("src=\"data:image/jpeg;base64,BAU=\"", result);
        Assert.Contains("alt=\"a\"", result); // other attributes untouched
        Assert.DoesNotContain("name:img", result);
    }

    [Fact]
    public void DoesNotConfuseImg1WithImg10()
    {
        var images = new List<ResolvedImage> { new("img1", "image/png", [1]) };
        var result = DataUriInliner.Inline("<img src=\"name:img10\"/><img src=\"name:img1\"/>", images);
        Assert.Contains("src=\"name:img10\"", result);
        Assert.Contains("src=\"data:image/png;base64,AQ==\"", result);
    }
}
