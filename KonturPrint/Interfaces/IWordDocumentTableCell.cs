using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    [Guid("D7F0A320-60FB-4E10-BF53-AC84B9875755")]
    [ComVisible(true)]
    public interface IWordDocumentTableCell
    {
        object CellProperties { get; set; }
        string Text { get; set; }

        bool Select(int rowNum, int colNum);
    }
}