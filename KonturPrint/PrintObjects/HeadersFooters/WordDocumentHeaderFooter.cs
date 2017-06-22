using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.Tables;

namespace KonturPrint.PrintObjects.HeadersFooters
{
    public abstract class WordDocumentHeaderFooter : IWordDocumentHeaderFooter, IPrintObject
    {
        protected WordprocessingDocument Doc { get; set; }
        public IWordDocumentTables Tables { get; protected set; }
        public OpenXmlElement XmlElement { get; protected set; }

        protected WordDocumentHeaderFooter(WordprocessingDocument doc)
        {
            Doc = doc;
        }

        protected WordDocumentHeaderFooter(WordprocessingDocument doc, OpenXmlElement xmlElement)
        {
            Doc = doc;
            XmlElement = xmlElement;
            Tables = new WordDocumentElementTables(doc, XmlElement);
        }

        public virtual IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.CopyTo(destPrintObject);
        }

        public virtual IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            if (TryGetCopyOfTable(sourcePrintObject))
            {
                var lastTable = Tables.Item(Tables.Count);
                return new PrintObject(lastTable.Table);
            }
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.GetCopyOf(sourcePrintObject);
        }

        protected virtual bool TryGetCopyOfTable(IPrintObject sourcePrintObject)
        {
            try
            {
                var tbl = sourcePrintObject as IWordDocumentTable;
                if (tbl != null)
                {
                    GetCopyOfTable(tbl);
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        protected virtual void GetCopyOfTable(IWordDocumentTable sourceTable)
        {
            var newTable = (Table)sourceTable.GetCopy().Table;
            XmlElement.AppendChild(newTable);
            Tables.Add(new WordDocumentTable(Doc, newTable));
        }
    }
}