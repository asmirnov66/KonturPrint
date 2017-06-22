using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Extensions;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.TableCells;
using KonturPrint.PrintObjects.TableRows;

namespace KonturPrint.PrintObjects.Tables
{
    public class WordDocumentTable : IWordDocumentTable, IPrintObject
    {
        private WordprocessingDocument Doc { get; }
        private Table InternalTable { get; set; }

        public object Table
        {
            get
            {
                return InternalTable;
            }
        }
        public string Name { get; private set; }
        public IWordDocumentTableRows Rows { get; private set; }
        public OpenXmlElement XmlElement { get; private set; }

        public object TableProperties
        {
            get { return GetTableProperties(); }
            set { SetTableProperties((TableProperties)value); }
        }

        public WordDocumentTable(WordprocessingDocument doc)
        {
            Doc = doc;
        }

        public WordDocumentTable(WordprocessingDocument doc, Table table)
        {
            Doc = doc;
            InternalTable = table;
            XmlElement = table;
            Name = GetTableName();
            Rows = new WordDocumentTableRows(this);
        }

        public bool Select(string name)
        {
            Table table;
            if (TryFindByName(name, out table))
            {
                InternalTable = table;
            }
            else
            {
                if (TryFindByLocalName(name, out table))
                {
                    InternalTable = table;
                }
                else
                {
                    return false;
                }
            }
            XmlElement = InternalTable;
            Name = GetTableName();
            Rows = new WordDocumentTableRows(this);
            return true;
        }

        public IWordDocumentTableCell Cell(int rowNum, int colNum)
        {
            IWordDocumentTableCell tableCell;
            if (TryGetCell(rowNum, colNum, out tableCell))
            {
                return tableCell;
            }
            return null;
        }

        public IWordDocumentTable GetCopy(string name = "")
        {
            var newTable = (Table)InternalTable.Clone();
            if (string.IsNullOrEmpty(name))
            {
                name = Name + "_copy";
            }
            newTable.GetFirstDescendant<TableProperties>().TableCaption.Val = name;
            return new WordDocumentTable(Doc, newTable);
        }

        public IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            if (Table == null)
            {
                return null;
            }
            if (TryCopyToTable(destPrintObject))
            {
                return new PrintObject(destPrintObject.XmlElement);
            }
            var sourceObject = (Table)InternalTable.Clone();
            sourceObject.GetFirstDescendant<TableProperties>().TableCaption.Val = Name + "_copy";
            var destObject = destPrintObject.XmlElement.AppendChild((OpenXmlElement)sourceObject);
            return new PrintObject(destObject);
        }

        public IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            if (InternalTable == null)
            {
                return null;
            }
            if (TryAppendCopyOfTable(sourcePrintObject))
            {
                return new PrintObject(InternalTable);
            }
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.GetCopyOf(sourcePrintObject);
        }

        private TableProperties GetTableProperties()
        {
            return InternalTable?.GetFirstChild<TableProperties>();
        }

        private void SetTableProperties(TableProperties tp)
        {
            if (InternalTable == null)
            {
                return;
            }
            if (TableProperties == null)
            {
                InternalTable.AppendChild(tp);
            }
            else
            {
                InternalTable.ReplaceChild(tp, (TableProperties)TableProperties);
            }
        }

        private bool TryGetCell(int rowNum, int colNum, out IWordDocumentTableCell tableCell)
        {
            if (Table != null)
            {
                var cell = new WordDocumentTableCell(this);
                if (cell.Select(rowNum, colNum))
                {
                    tableCell = cell;
                    return true;
                }
            }
            tableCell = null;
            return false;
        }

        private bool TryAppendCopyOfTable(IPrintObject sourcePrintObject)
        {
            try
            {
                var tbl = sourcePrintObject as IWordDocumentTable;
                if (tbl != null)
                {
                    CopyTableToTable((Table)tbl.Table, InternalTable);
                    Rows = new WordDocumentTableRows(this);
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool TryCopyToTable(IPrintObject destPrintObject)
        {
            try
            {
                var tbl = destPrintObject as IWordDocumentTable;
                if (tbl != null)
                {
                    destPrintObject.GetCopyOf(this);
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void CopyTableToTable(Table sourceTable, Table destTable)
        {
            var rows = sourceTable.Descendants<TableRow>();
            foreach (var row in rows)
            {
                AppendCopyOfRow(destTable, row);
            }
        }

        private void AppendCopyOfRow(Table destTable, TableRow row)
        {
            var newRow = (TableRow)row.Clone();
            destTable.Append(newRow);
        }

        private bool TryFindByName(string name, out Table table)
        {
            var main = Doc.MainDocumentPart;
            var body = main.Document.Body;
            var tableProperties = body.Descendants<TableProperties>()
                                  .Where(tp => tp.TableCaption != null)
                                  .FirstOrDefault(tp => string.Equals(tp.TableCaption.Val, name, StringComparison.OrdinalIgnoreCase));
            if (tableProperties == null)
            {
                table = null;
                return false;
            }
            table = (Table)tableProperties.Parent;
            return true;
        }

        private bool TryFindByLocalName(string name, out Table table)
        {
            var main = Doc.MainDocumentPart;
            var body = main.Document.Body;
            var tbl = body.Descendants<Table>()
                                  .FirstOrDefault(t => string.Equals(t.LocalName, name, StringComparison.OrdinalIgnoreCase));
            if (tbl == null)
            {
                table = null;
                return false;
            }
            table = tbl;
            return true;
        }

        private string GetTableName()
        {
            var res = InternalTable.LocalName + "_" + Guid.NewGuid();
            var tableProperties = InternalTable.GetFirstChild<TableProperties>();
            if (tableProperties == null)
            {
                InternalTable.AppendChild(new TableProperties(new TableCaption { Val = res }));
                return res;
            }
            var caption = tableProperties.GetFirstDescendant<TableCaption>();
            if (caption == null)
            {
                tableProperties.Append(new TableCaption { Val = res });
                return res;
            }
            if (string.IsNullOrEmpty(caption.Val))
            {
                caption.Val = res;
                return res;
            }
            return caption.Val;
        }
    }
}