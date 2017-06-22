using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.PrintObjects.Tables;

namespace KonturPrint.PrintObjects.HeadersFooters.Headers
{
    public class WordDocumentHeader : WordDocumentHeaderFooter
    {
        public WordDocumentHeader(WordprocessingDocument doc) : base(doc)
        {
            XmlElement = TryFindHeader();
            Tables = new WordDocumentElementTables(doc, XmlElement);
        }

        public WordDocumentHeader(WordprocessingDocument doc, Header header) : base(doc, header)
        {
            XmlElement = header;
        }

        private Header TryFindHeader()
        {
            var hpart = Doc.MainDocumentPart.HeaderParts.FirstOrDefault();
            return hpart?.Header;
        }
    }
}