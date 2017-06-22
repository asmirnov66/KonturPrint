using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.Tables
{
    public class WordDocumentElementTables : IWordDocumentTables
    {
        private WordprocessingDocument Doc { get; }
        private OpenXmlElement Element { get; }
        private IDictionary<string, IWordDocumentTable> Items { get; set; }

        public WordDocumentElementTables(WordprocessingDocument doc, OpenXmlElement element)
        {
            Doc = doc;
            Element = element;
            FillItems();
        }

        public int Count
        {
            get
            {
                FillItems();
                return Items.Count;
            }
        }

        public IWordDocumentTable Item(string index)
        {
            IWordDocumentTable tbl;
            if (Items.TryGetValue(index, out tbl))
            {
                return tbl;
            }
            return null;
        }

        public IWordDocumentTable Item(int index)
        {
            if (index <= 0)
            {
                return null;
            }
            var tn = index - 1;
            if (tn < Items.Count)
            {
                return Items.ElementAt(tn).Value;
            }
            return null;
        }

        public bool Exists(string name)
        {
            IWordDocumentTable tbl;
            return Items.TryGetValue(name, out tbl);
        }

        public IWordDocumentTable Add(IWordDocumentTable table)
        {
            if (string.IsNullOrEmpty(table.Name))
            {
                return null;
            }
            IWordDocumentTable tbl;
            if (Items.TryGetValue(table.Name, out tbl))
            {
                return null;
            }
            Items.Add(table.Name, table);
            return table;
        }

        private void FillItems()
        {
            Items = new Dictionary<string, IWordDocumentTable>(StringComparer.OrdinalIgnoreCase);
            if (Element == null)
            {
                return;
            }
            var tables = Element.Descendants<Table>();
            foreach (var t in tables)
            {
                var el = new WordDocumentTable(Doc, t);
                Items.Add(el.Name, el);
            }
        }
    }
}