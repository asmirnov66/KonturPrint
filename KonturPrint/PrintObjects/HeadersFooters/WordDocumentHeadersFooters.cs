using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintDocuments;

namespace KonturPrint.PrintObjects.HeadersFooters
{
    public abstract class WordDocumentHeadersFooters : IWordDocumentHeadersFooters
    {
        protected WordprocessingDocument Doc { get; }
        protected IDictionary<int, IWordDocumentHeaderFooter> Items { get; set; }
        protected IWordDocumentStructure WordDocument { get; set; }
        protected IDictionary<int, IWordDocumentHeaderFooter> ItemsByType { get; set; }

        public int Count
        {
            get
            {
                FillItems();
                return Items.Count;
            }
        }

        protected WordDocumentHeadersFooters(WordprocessingDocument doc)
        {
            Doc = doc;
            WordDocument = new WordTemplateDocument(doc);
            ItemsByType = new Dictionary<int, IWordDocumentHeaderFooter>();
        }

        protected WordDocumentHeadersFooters(IWordDocumentStructure wordDoc)
        {
            Doc = wordDoc.InnerDoc;
            WordDocument = wordDoc;
            ItemsByType = new Dictionary<int, IWordDocumentHeaderFooter>();
        }

        public virtual IWordDocumentHeaderFooter Item(int index)
        {
            if (index <= 0)
            {
                return null;
            }
            var fn = index - 1;
            if (fn < Items.Count)
            {
                return Items.ElementAt(fn).Value;
            }
            return null;
        }

        public virtual IWordDocumentHeaderFooter ItemByType(int type)
        {
            IWordDocumentHeaderFooter item;
            if (ItemsByType.TryGetValue(type, out item))
            {
                return item;
            }
            return null;
        }

        public abstract void Create(int type);

        protected abstract void FillItems();

        protected virtual HeadersFootersType GetHeadersFootersType(int type)
        {
            HeadersFootersType? elType;
            if (TryGetHeadersFootersType(type, out elType))
            {
                if (elType != null)
                {
                    return (HeadersFootersType)elType;
                }
            }
            return HeadersFootersType.Default;
        }

        protected virtual void DeleteAll<T, T1>(IEnumerable<T1> parts) where T : HeaderFooterReferenceType where T1 : OpenXmlPart
        {
            var mainDocumentPart = Doc.MainDocumentPart;
            mainDocumentPart.DeleteParts(parts);
            var sectionProperties = WordDocument.SectionProperties;
            sectionProperties.RemoveAllChildren<T>();
            ItemsByType = new Dictionary<int, IWordDocumentHeaderFooter>();
            Items = new Dictionary<int, IWordDocumentHeaderFooter>();
        }

        protected virtual void CreateElements<T>(HeadersFootersType elementType, Func<HeaderFooterValues, T> addNewFunc) where T : HeaderFooterReferenceType
        {
            var referenceList = new List<T>();
            var sectionProperties = WordDocument.SectionProperties;
            var settings = WordDocument.Settings;

            if (elementType == HeadersFootersType.FirstDefault || elementType == HeadersFootersType.FirstEvenOdd)
            {
                if (sectionProperties.GetFirstChild<TitlePage>() == null)
                {
                    sectionProperties.Append(new TitlePage());
                }
                referenceList.Add(addNewFunc(HeaderFooterValues.First));
            }
            if (elementType == HeadersFootersType.FirstEvenOdd || elementType == HeadersFootersType.EvenOdd)
            {
                if (settings.GetFirstChild<EvenAndOddHeaders>() == null)
                {
                    settings.Append(new EvenAndOddHeaders());
                }
                referenceList.Add(addNewFunc(HeaderFooterValues.Even));
            }
            referenceList.Add(addNewFunc(HeaderFooterValues.Default));
            foreach (var el in referenceList)
            {
                sectionProperties.Append(el);
            }
        }

        protected abstract HeaderFooterReferenceType AddNew(HeaderFooterValues valueType);

        protected virtual bool TryGetHeadersFootersType(int type, out HeadersFootersType? headerFooterType)
        {
            if (Enum.IsDefined(typeof(HeadersFootersType), type))
            {
                headerFooterType = (HeadersFootersType)type;
                return true;
            }
            headerFooterType = null;
            return false;
        }
    }
}