using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.TableRows
{
    public class WordDocumentTableRow : IWordDocumentTableRow
    {
        public TableRow Row { get; }

        public TableRowProperties RowProperties
        {
            get { return GetRowProperties(); }
            set { SetRowProperties(value); }
        }

        public WordDocumentTableRow()
        {
            Row = new TableRow();
        }

        public WordDocumentTableRow(TableRow row)
        {
            Row = row;
        }

        private TableRowProperties GetRowProperties()
        {
            return Row?.TableRowProperties;
        }

        private void SetRowProperties(TableRowProperties rp)
        {
            if (Row == null)
            {
                return;
            }
            Row.TableRowProperties = rp;
        }
    }
}