using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.TableRows
{
    public class WordDocumentTableRows : IWordDocumentTableRows
    {
        private Table Table { get; }
        private IDictionary<int, IWordDocumentTableRow> Items { get; set; }

        public int Count
        {
            get
            {
                FillItems();
                return Items.Count;
            }
        }

        public WordDocumentTableRows(IWordDocumentTable table)
        {
            Table = (Table)table.Table;
        }

        public WordDocumentTableRows(Table table)
        {
            Table = table;
        }

        public WordDocumentTableRows(object table)
        {
            Table = (Table)table;
        }

        public IWordDocumentTableRow Item(int index)
        {
            if (index <= 0)
            {
                return null;
            }
            InitItems();
            var rn = index - 1;
            if (rn < Items.Count)
            {
                return Items.ElementAt(rn).Value;
            }
            return null;
        }

        public IWordDocumentTableRow Add()
        {
            return InternalAdd();
        }

        public IWordDocumentTableRow AddEmpty()
        {
            return InternalAdd(false);
        }

        public IWordDocumentTableRow Copy()
        {
            InitItems();
            var newRow = new TableRow();
            CopyCellsByLastRow(newRow);
            return AppendRow(newRow);
        }

        public IWordDocumentTableRow CopyRow(int index)
        {
            if (index <= 0)
            {
                return null;
            }
            var sourceRow = Item(index).Row;
            if (sourceRow == null)
            {
                return null;
            }
            var newRow = new TableRow();
            CopyCells(sourceRow, newRow, true);
            return AppendRow(newRow);
        }

        private void FillItems()
        {
            Items = new Dictionary<int, IWordDocumentTableRow>();
            var rows = Table.Elements<TableRow>();
            var i = 0;
            foreach (var row in rows)
            {
                Items.Add(i++, new WordDocumentTableRow(row));
            }
        }

        private void InitItems()
        {
            if (Items == null)
            {
                FillItems();
            }
        }

        private IWordDocumentTableRow InternalAdd(bool copyProp = true)
        {
            InitItems();
            var newRow = new TableRow();
            AddCellsByLastRow(newRow, false);
            return AppendRow(newRow);
        }

        private void AddCellsByLastRow(TableRow row, bool copyProp = true)
        {
            var lastRow = Table.Elements<TableRow>().LastOrDefault();
            if (lastRow == null)
            {
                var newCell = new TableCell();
                newCell.Append(new Paragraph(new Run(new Text())));
                row.Append(newCell);
                return;
            }
            CopyCells(lastRow, row, false, copyProp);
        }

        private void CopyCellsByLastRow(TableRow row)
        {
            var lastRow = Table.Elements<TableRow>().LastOrDefault();
            if (lastRow == null)
            {
                return;
            }
            CopyCells(lastRow, row, true);
        }

        private void CopyCells(TableRow sourceRow, TableRow destRow, bool copyText = false, bool copyProp = true)
        {
            var cells = sourceRow.Elements<TableCell>();
            foreach (var cell in cells)
            {
                if (copyText)
                {
                    var newCell = new TableCell(cell.OuterXml);
                    destRow.Append(newCell);
                }
                else
                {
                    var newCell = new TableCell();
                    if (copyProp)
                    {
                        var cp = new TableCellProperties(cell.TableCellProperties.OuterXml);
                        newCell.TableCellProperties = cp;
                    }
                    newCell.Append(new Paragraph(new Run(new Text())));
                    destRow.Append(newCell);
                }
            }
        }

        private IWordDocumentTableRow AppendRow(TableRow newRow)
        {
            Table.Append(newRow);
            var newTableRow = new WordDocumentTableRow(newRow);
            Items.Add(Items.Count + 1, newTableRow);
            return newTableRow;
        }
    }
}