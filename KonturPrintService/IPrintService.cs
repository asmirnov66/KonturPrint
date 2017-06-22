using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace KonturPrintService
{
    [Guid("DFFF3A1E-368E-4CF9-83EE-952432D3A6B6")]
    [ComVisible(true)]
    public interface IPrintService
    {
        object PrintScripts { get; set; }

        bool Print(int printDocumentType, object printParams = null);
        object GetPrintDocument(int printDocumentType, object printParams = null);
        void SavePrintDocument(object printParams = null);
    }
}