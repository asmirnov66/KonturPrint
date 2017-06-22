using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.Tables
{
    public class WordDocumentTables : IWordDocumentTables
    {
        private WordprocessingDocument Doc { get; }
        private IDictionary<string, IWordDocumentTable> Items { get; set; }
        private IWordDocumentStructure WordDocument { get; set; }

        public int Count
        {
            get
            {
                FillItems();
                return Items.Count;
            }
        }

        public WordDocumentTables(WordprocessingDocument doc)
        {
            Doc = doc;
            Items = new Dictionary<string, IWordDocumentTable>(StringComparer.OrdinalIgnoreCase);
        }

        public WordDocumentTables(IWordDocumentStructure wordDoc)
        {
            Doc = wordDoc.InnerDoc;
            WordDocument = wordDoc;
            Items = new Dictionary<string, IWordDocumentTable>(StringComparer.OrdinalIgnoreCase);
        }

        public IWordDocumentTable Item(string index)
        {
            IWordDocumentTable tbl;
            if (TryGetTable(index, out tbl))
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
            if (TryGetTable(name, out tbl))
            {
                return true;
            }
            return false;
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
            var main = Doc.MainDocumentPart;
            var body = main.Document.Body;
            var tableProperties = body.Descendants<TableProperties>()
                                  .Where(tp => tp.TableCaption != null)
                                  .Select(tp => tp);
            foreach (var tp in tableProperties)
            {
                var tlb = (Table)tp.Parent;
                var name = tp.TableCaption.Val;
                Items.Add(name, new WordDocumentTable(Doc, tlb));
            }
        }

        private bool TryGetTable(string name, out IWordDocumentTable table)
        {
            IWordDocumentTable tbl;
            if (Items.TryGetValue(name, out tbl))
            {
                table = tbl;
                return true;
            }
            tbl = GetTable(name);
            if (tbl != null)
            {
                Items.Add(tbl.Name, tbl);
                table = tbl;
                return true;
            }
            table = null;
            return false;
        }

        private IWordDocumentTable GetTable(string name)
        {
            var tbl = new WordDocumentTable(Doc);
            if (tbl.Select(name))
            {
                return tbl;
            }
            return null;
        }
    }
}