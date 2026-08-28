using System.Runtime.InteropServices;

namespace SendToOneNote.Core.Desktop;

/// <summary>Enum values from the OneNote type library, as ints (we never load the typelib).</summary>
public static class OneNoteConstants
{
    public const string ProgId = "OneNote.Application";
    public const int HsSections = 3;    // HierarchyScope.hsSections
    public const int HsPages = 4;       // HierarchyScope.hsPages
    public const int Xs2013 = 2;        // XMLSchema.xs2013
    public const int NpsDefault = 0;    // NewPageStyle.npsDefault
    public const int PiBinaryData = 1;  // PageInfo.piBinaryData
    public static readonly string Namespace2013 = "http://schemas.microsoft.com/office/onenote/2013/onenote";
}

/// <summary>
/// OneNote 2013+ IApplication, hand-declared so calls go through the COM vtable.
/// Late binding (dynamic / InvokeMember) fails on Click-to-Run Office because the
/// type library lives in a virtualized registry hive (TYPE_E_LIBNOTREGISTERED).
/// Methods MUST stay in IDL order; only the prefix through GetHyperlinkToObject is declared.
/// Verified against the live app 2026-08-28.
/// </summary>
[ComImport, Guid("452AC71A-B655-4967-A208-A4CC39DD7949"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IApplication
{
    void GetHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrStartNodeID, int hsScope,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHierarchyXmlOut, int xsSchema);
    void UpdateHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrChangesXmlIn, int xsSchema);
    void OpenHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrPath, [MarshalAs(UnmanagedType.BStr)] string bstrRelativeToObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrObjectID, int cftIfNotExist);
    void DeleteHierarchy([MarshalAs(UnmanagedType.BStr)] string bstrObjectID, DateTime dateExpectedLastModified, bool deletePermanently);
    void CreateNewPage([MarshalAs(UnmanagedType.BStr)] string bstrSectionID, [MarshalAs(UnmanagedType.BStr)] out string pbstrPageID, int npsNewPageStyle);
    void CloseNotebook([MarshalAs(UnmanagedType.BStr)] string bstrNotebookID, bool force);
    void GetHierarchyParent([MarshalAs(UnmanagedType.BStr)] string bstrObjectID, [MarshalAs(UnmanagedType.BStr)] out string pbstrParentID);
    void GetPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] out string pbstrPageXMLOut,
        int pageInfoToExport, int xsSchema);
    void UpdatePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageChangesXmlIn, DateTime dateExpectedLastModified, int xsSchema, bool force);
    void GetBinaryPageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] string bstrCallbackID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrBinaryObjectB64Out);
    void DeletePageContent([MarshalAs(UnmanagedType.BStr)] string bstrPageID, [MarshalAs(UnmanagedType.BStr)] string bstrObjectID,
        DateTime dateExpectedLastModified, bool force);
    void NavigateTo([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyObjectID, [MarshalAs(UnmanagedType.BStr)] string bstrObjectID, bool fNewWindow);
    void NavigateToUrl([MarshalAs(UnmanagedType.BStr)] string bstrUrl, bool fNewWindow);
    void Publish([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID, [MarshalAs(UnmanagedType.BStr)] string bstrTargetFilePath,
        int pfPublishFormat, [MarshalAs(UnmanagedType.BStr)] string bstrCLSIDofExporter);
    void OpenPackage([MarshalAs(UnmanagedType.BStr)] string bstrPathPackage, [MarshalAs(UnmanagedType.BStr)] string bstrPathDest,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrPathOut);
    void GetHyperlinkToObject([MarshalAs(UnmanagedType.BStr)] string bstrHierarchyID, [MarshalAs(UnmanagedType.BStr)] string bstrPageContentObjectID,
        [MarshalAs(UnmanagedType.BStr)] out string pbstrHyperlinkOut);
}
