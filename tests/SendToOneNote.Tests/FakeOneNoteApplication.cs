using SendToOneNote.Core.Desktop;

namespace SendToOneNote.Tests;

/// <summary>Managed implementation of the COM interface — records calls, returns canned data.</summary>
public sealed class FakeOneNoteApplication : IApplication
{
    public string HierarchyXml { get; set; } = """
    <one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
      <one:Notebook name="Alpha" ID="{N1}"><one:Section name="Inbox" ID="{S1}"/></one:Notebook>
    </one:Notebooks>
    """;
    public string NextPageId { get; set; } = "{P1}{1}{B0}";
    public string Hyperlink { get; set; } = "onenote:https://example.test/Alpha/Inbox.one#p1";
    public List<(string SectionId, int Style)> CreatedPages { get; } = [];
    public List<string> UpdatedPageXml { get; } = [];
    public int ManagedThreadIdOfLastCall { get; private set; }
    public Exception? ThrowOnUpdate { get; set; }

    public void GetHierarchy(string bstrStartNodeID, int hsScope, out string pbstrHierarchyXmlOut, int xsSchema)
    { Touch(); pbstrHierarchyXmlOut = HierarchyXml; }
    public void CreateNewPage(string bstrSectionID, out string pbstrPageID, int npsNewPageStyle)
    { Touch(); CreatedPages.Add((bstrSectionID, npsNewPageStyle)); pbstrPageID = NextPageId; }
    public void UpdatePageContent(string bstrPageChangesXmlIn, DateTime dateExpectedLastModified, int xsSchema, bool force)
    { Touch(); if (ThrowOnUpdate is not null) throw ThrowOnUpdate; UpdatedPageXml.Add(bstrPageChangesXmlIn); }
    public void GetHyperlinkToObject(string bstrHierarchyID, string bstrPageContentObjectID, out string pbstrHyperlinkOut)
    { Touch(); pbstrHyperlinkOut = Hyperlink; }

    private void Touch() => ManagedThreadIdOfLastCall = Environment.CurrentManagedThreadId;

    // Unused members of the interface:
    public void UpdateHierarchy(string a, int b) => throw new NotImplementedException();
    public void OpenHierarchy(string a, string b, out string c, int d) => throw new NotImplementedException();
    public void DeleteHierarchy(string a, DateTime b, bool c) => throw new NotImplementedException();
    public void CloseNotebook(string a, bool b) => throw new NotImplementedException();
    public void GetHierarchyParent(string a, out string b) => throw new NotImplementedException();
    public void GetPageContent(string a, out string b, int c, int d) => throw new NotImplementedException();
    public void GetBinaryPageContent(string a, string b, out string c) => throw new NotImplementedException();
    public void DeletePageContent(string a, string b, DateTime c, bool d) => throw new NotImplementedException();
    public void NavigateTo(string a, string b, bool c) => throw new NotImplementedException();
    public void NavigateToUrl(string a, bool b) => throw new NotImplementedException();
    public void Publish(string a, string b, int c, string d) => throw new NotImplementedException();
    public void OpenPackage(string a, string b, out string c) => throw new NotImplementedException();
}
