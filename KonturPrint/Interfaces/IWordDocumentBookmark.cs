using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("9DD2B90B-DD47-4F43-9E1B-A3E3D2A41E29")]
    [ComVisible(true)]
    public interface IWordDocumentBookmark
    {
        string Name { get; }
        string Text { get; set; }
        bool Select(string name);
        IWordDocumentTable Table { get; }
    }
}