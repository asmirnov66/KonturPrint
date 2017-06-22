using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.PrintObjects.Tables;

namespace KonturPrint.PrintObjects.HeadersFooters.Footers
{
    public class WordDocumentFooter : WordDocumentHeaderFooter
    {
        public WordDocumentFooter(WordprocessingDocument doc) : base(doc)
        {
            XmlElement = TryFindFooter();
            Tables = new WordDocumentElementTables(doc, XmlElement);
        }

        public WordDocumentFooter(WordprocessingDocument doc, Footer footer) : base(doc, footer)
        {
            XmlElement = footer;
        }

        private Footer TryFindFooter()
        {
            var fpart = Doc.MainDocumentPart.FooterParts.FirstOrDefault();
            return fpart?.Footer;
        }
    }
}