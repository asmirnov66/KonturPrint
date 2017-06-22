using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("AB66B08B-876D-485B-BECD-6B84621A1F61")]
    [ComVisible(true)]
    public interface IWordDocumentTables
    {
        int Count { get; }

        bool Exists(string name);
        IWordDocumentTable Item(int index);
        IWordDocumentTable Item(string index);
        IWordDocumentTable Add(IWordDocumentTable table);
    }
}