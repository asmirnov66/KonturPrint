using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KonturPrint.Interfaces
{
    public interface IWordDocumentStructure
    {
        WordprocessingDocument InnerDoc { get; }
        SectionProperties SectionProperties { get; }
        Settings Settings { get; }
    }
}