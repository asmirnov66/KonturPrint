using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Extensions;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.TableCells
{
    public class WordDocumentTableCell : IWordDocumentTableCell, IPrintObject
    {
        private Table Table { get; }
        private TableCell Cell { get; set; }

        public OpenXmlElement XmlElement { get; private set; }
        public object CellProperties
        {
            get
            {
                return GetCellProperties();
            }
            set
            {
                SetCellProperties((TableCellProperties)value);
            }
        }
        public string Text
        {
            get { return GetText(); }
            set { SetText(value); }
        }

        public WordDocumentTableCell(IWordDocumentTable table)
        {
            Table = (Table)table.Table;
        }

        public WordDocumentTableCell(Table table)
        {
            Table = table;
        }

        public WordDocumentTableCell(object table)
        {
            Table = (Table)table;
        }

        public bool Select(int rowNum, int colNum)
        {
            if ((rowNum <= 0) || (colNum <= 0))
            {
                return false;
            }
            var rn = rowNum - 1;
            var cn = colNum - 1;
            try
            {
                var row = Table?.Elements<TableRow>().ElementAt(rn);
                Cell = row?.Elements<TableCell>().ElementAt(cn);
            }
            catch
            {
                Cell = null;
            }
            XmlElement = Cell;
            return Cell != null;
        }

        public IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.CopyTo(destPrintObject);
        }

        public IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.GetCopyOf(sourcePrintObject);
        }

        private TableCellProperties GetCellProperties()
        {
            return Cell?.TableCellProperties;
        }

        private void SetCellProperties(TableCellProperties cp)
        {
            if (Cell == null)
            {
                return;
            }
            Cell.TableCellProperties = cp;
        }

        private string GetText()
        {
            var cellText = GetCellText();
            if (cellText == null)
            {
                return string.Empty;
            }
            var res = new StringBuilder();
            var list = Cell.Descendants<Text>();
            foreach (var t in list)
            {
                res.AppendLine(t.Text);
            }
            return res.ToString().TrimEnd();
        }

        private void SetText(string text)
        {
            var cellText = GetCellText();
            if (cellText == null)
            {
                AddTextToEmptyCell(text);
                return;
            }
            cellText.Text = text;
        }

        private Text GetCellText()
        {
            return Cell?.GetFirstDescendant<Text>();
        }

        private void AddTextToEmptyCell(string text)
        {
            if (Cell == null)
            {
                return;
            }
            var p = Cell.GetFirstDescendant<Paragraph>();
            if (p == null)
            {
                Cell.Append(new Paragraph(new Run(new Text(text))));
                return;
            }
            var r = p.GetFirstDescendant<Run>();
            if (r == null)
            {
                p.Append(new Run(new Text(text)));
                return;
            }
            r.Append(new Text(text));
        }
    }
}