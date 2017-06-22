using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("17C2FA5F-966E-4422-A1A4-73107ACC7802")]
    [ComVisible(true)]
    public interface IWordDocumentBookmarks
    {
        int Count { get; }
        IWordDocumentBookmark Item(string index);
        IWordDocumentBookmark Item(int index);
        bool Exists(string name);
    }
}