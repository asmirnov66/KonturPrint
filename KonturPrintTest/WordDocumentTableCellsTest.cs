using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.PrintObjects.TableCells;
using KonturPrint.PrintObjects.Tables;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentTableCellsTest
    {
        private string TestDirectory { get; set; }
        private string TemplatePath { get; set; }

        [OneTimeSetUp]
        public void SetUp()
        {
            TestDirectory = TestContext.CurrentContext.TestDirectory;
            TemplatePath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.docx");
        }

        [Test]
        public void CanSelectCell()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);

                Assert.IsTrue(cell.Select(1, 1));
            }
        }

        [Test]
        public void GetTableCellProperties()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);
                cell.Select(1, 1);
                var cp = (TableCellProperties)cell.CellProperties;

                Assert.IsNotNull(cp);
                Assert.AreEqual(cp.TableCellWidth.Width.ToString(), "4809");
            }
        }

        [Test]
        public void SetTableCellProperties()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);
                cell.Select(1, 1);

                var ncp = new TableCellProperties(
                    new TableCellWidth
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = "2400"
                    });
                cell.CellProperties = ncp;
                var cp = (TableCellProperties)cell.CellProperties;

                Assert.IsNotNull(cp);
                Assert.AreEqual(cp.TableCellWidth.Width.ToString(), "2400");
            }
        }

        [Test]
        public void GetTableCellText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);
                cell.Select(3, 2);
                var txt = cell.Text;

                Assert.AreEqual(txt, "TextInColumn");
            }
        }

        [Test]
        public void GetTableEmptyCellText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);
                cell.Select(3, 1);
                var txt = cell.Text;

                Assert.IsTrue(string.IsNullOrEmpty(txt));
            }
        }

        [Test]
        public void SetTableCellText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = new WordDocumentTableCell(table.Table);
                cell.Select(3, 2);
                const string val = "New Text";

                cell.Text = val;

                Assert.AreEqual(cell.Text, val);
            }
        }
    }
}