using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("4C7CA327-3A0A-43E7-B572-5B4540CBEDB8")]
    [ComVisible(true)]
    public interface IWordDocumentTableRows
    {
        int Count { get; }

        IWordDocumentTableRow Add();
        IWordDocumentTableRow AddEmpty();
        IWordDocumentTableRow Copy();
        IWordDocumentTableRow CopyRow(int index);
        IWordDocumentTableRow Item(int index);
    }
}