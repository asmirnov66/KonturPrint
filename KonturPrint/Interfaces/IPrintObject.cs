using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;

namespace KonturPrint.Interfaces
{
    [Guid("AC70EA74-F442-4241-85D8-C051AE33F278")]
    [ComVisible(true)]
    public interface IPrintObject
    {
        OpenXmlElement XmlElement { get; }

        IPrintObject CopyTo(IPrintObject destPrintObject);
        IPrintObject GetCopyOf(IPrintObject sourcePrintObject);
    }
}