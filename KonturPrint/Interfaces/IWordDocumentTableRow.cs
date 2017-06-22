using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KonturPrint.Interfaces
{
    [Guid("076FDE90-89F3-401E-ACD7-D04961ABD43D")]
    [ComVisible(true)]
    public interface IWordDocumentTableRow
    {
        TableRow Row { get; }

        TableRowProperties RowProperties { get; set; }
    }
}