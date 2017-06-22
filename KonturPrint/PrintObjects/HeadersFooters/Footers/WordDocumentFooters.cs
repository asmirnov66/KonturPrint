using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.HeadersFooters.Footers
{
    public class WordDocumentFooters : WordDocumentHeadersFooters
    {
        public WordDocumentFooters(WordprocessingDocument doc) : base(doc)
        {
            FillItems();
        }

        public WordDocumentFooters(IWordDocumentStructure wordDoc) : base(wordDoc)
        {
            FillItems();
        }

        protected sealed override void FillItems()
        {
            Items = new Dictionary<int, IWordDocumentHeaderFooter>();
            var main = Doc.MainDocumentPart;
            var i = 0;
            foreach (var f in main.FooterParts)
            {
                Items.Add(i++, new WordDocumentFooter(Doc, f.Footer));
            }
        }

        public override void Create(int type)
        {
            var footerType = GetHeadersFootersType(type);
            var mainDocumentPart = Doc.MainDocumentPart;
            DeleteAll<FooterReference, FooterPart>(mainDocumentPart.FooterParts);
            CreateElements(footerType, AddNew);
        }

        protected override HeaderFooterReferenceType AddNew(HeaderFooterValues valueType)
        {
            var mainDocumentPart = Doc.MainDocumentPart;
            var pageFooterPart = mainDocumentPart.AddNewPart<FooterPart>();
            var pageFooterPartId = mainDocumentPart.GetIdOfPart(pageFooterPart);
            var footer = new Footer();
            pageFooterPart.Footer = footer;
            var wordFooter = new WordDocumentFooter(Doc, footer);
            Items.Add(Items.Count, wordFooter);
            ItemsByType.Add((int)valueType, wordFooter);
            return new FooterReference { Id = pageFooterPartId, Type = valueType };
        }
    }
}