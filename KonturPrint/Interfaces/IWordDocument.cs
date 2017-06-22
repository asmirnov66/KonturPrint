using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("3A02D24D-960E-4F57-84AC-999557ED304B")]
    [ComVisible(true)]
    public interface IWordDocument
    {
        IWordDocumentBookmarks Bookmarks { get; }
        IWordDocumentTables Tables { get; }
        IWordDocumentHeadersFooters Footers { get; }
        IWordDocumentHeadersFooters Headers { get; }
        IPrintObject PageEnumerator { get; }
    }
}