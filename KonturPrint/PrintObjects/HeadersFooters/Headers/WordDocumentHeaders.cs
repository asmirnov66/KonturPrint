using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.HeadersFooters.Headers
{
    public class WordDocumentHeaders : WordDocumentHeadersFooters
    {
        public WordDocumentHeaders(WordprocessingDocument doc) : base(doc)
        {
            FillItems();
        }

        public WordDocumentHeaders(IWordDocumentStructure wordDoc) : base(wordDoc)
        {
            FillItems();
        }

        protected sealed override void FillItems()
        {
            Items = new Dictionary<int, IWordDocumentHeaderFooter>();
            var main = Doc.MainDocumentPart;
            var i = 0;
            foreach (var h in main.HeaderParts)
            {
                Items.Add(i++, new WordDocumentHeader(Doc, h.Header));
            }
        }

        public override void Create(int type)
        {
            var headerType = GetHeadersFootersType(type);
            var mainDocumentPart = Doc.MainDocumentPart;
            DeleteAll<HeaderReference, HeaderPart>(mainDocumentPart.HeaderParts);
            CreateElements(headerType, AddNew);
        }

        protected override HeaderFooterReferenceType AddNew(HeaderFooterValues valueType)
        {
            var mainDocumentPart = Doc.MainDocumentPart;
            var pageHeaderPart = mainDocumentPart.AddNewPart<HeaderPart>();
            var pageHeaderPartId = mainDocumentPart.GetIdOfPart(pageHeaderPart);
            var header = new Header
            {
                MCAttributes = new MarkupCompatibilityAttributes
                {
                    Ignorable = "w14 w15 wp14"
                }
            };
            header.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            pageHeaderPart.Header = header;
            var wordHeader = new WordDocumentHeader(Doc, header);
            Items.Add(Items.Count, wordHeader);
            ItemsByType.Add((int)valueType, wordHeader);
            return new HeaderReference { Id = pageHeaderPartId, Type = valueType };
        }
    }
}