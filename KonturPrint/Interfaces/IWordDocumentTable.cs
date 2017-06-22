using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("C4771D40-419B-41C6-8CF0-2CEB406F92E4")]
    [ComVisible(true)]
    public interface IWordDocumentTable
    {
        string Name { get; }
        object Table { get; }
        object TableProperties { get; set; }
        IWordDocumentTableRows Rows { get; }

        IWordDocumentTableCell Cell(int rowNum, int colNum);
        bool Select(string name);
        IWordDocumentTable GetCopy(string name = "");
    }
}