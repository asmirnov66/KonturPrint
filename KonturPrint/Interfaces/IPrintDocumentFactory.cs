namespace KonturPrint.Interfaces
{
    public enum PrintDocumentType
    {
        ExcelTemplate,
        WordTemplate,
        ExcelWithMacros
    }

    public interface IPrintDocumentFactory
    {
        IPrintDocument GetPrintDocument(PrintDocumentType documentType);
    }
}