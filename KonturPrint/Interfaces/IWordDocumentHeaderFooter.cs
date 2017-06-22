using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;

namespace KonturPrint.Interfaces
{
    [Guid("A3230F7D-2C58-4F19-A8CD-77B05D7BDE25")]
    [ComVisible(true)]
    public interface IWordDocumentHeaderFooter
    {
        OpenXmlElement XmlElement { get; }
        IWordDocumentTables Tables { get; }
    }
}