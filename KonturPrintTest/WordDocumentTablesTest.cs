using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.PrintObjects.Bookmarks;
using KonturPrint.PrintObjects.Tables;
using NUnit.Framework;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentTablesTest
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
        public void CanSelectTableByName()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);

                Assert.IsTrue(table.Select("Table1"));
                Assert.IsNotNull(table.Table);
            }
        }

        [Test]
        public void GetTableProperties()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var tp = (TableProperties)table.TableProperties;

                Assert.IsNotNull(tp);
                Assert.AreEqual(tp.TableCaption.Val.ToString(), "Table1");
            }
        }

        [Test]
        public void SetTableProperties()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var ntp = new TableProperties(
                    new TableCaption
                    {
                        Val = "Table2"
                    });
                table.TableProperties = ntp;
                var tp = (TableProperties)table.TableProperties;

                Assert.IsNotNull(tp);
                Assert.AreEqual(tp.TableCaption.Val.ToString(), "Table2");
            }
        }

        [Test]
        public void CanGetCell()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = table.Cell(1, 1);

                Assert.IsNotNull(cell);
            }
        }

        [Test]
        public void GetCellText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = table.Cell(3, 2);
                var txt = cell.Text;

                Assert.AreEqual(txt, "TextInColumn");
            }
        }

        [Test]
        public void SetCellText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var cell = table.Cell(3, 2);
                const string val = "New Text";

                cell.Text = val;

                Assert.AreEqual(cell.Text, val);
            }
        }

        [Test]
        public void GetRowCount()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");

                Assert.AreEqual(table.Rows.Count, 3);
            }
        }

        [Test]
        public void CanAddRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");

                table.Rows.Add();
                var cell = table.Cell(4, 2);

                Assert.IsNotNull(cell);
            }
        }

        [Test]
        public void CopyRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");

                table.Rows.Copy();
                var cell = table.Cell(4, 2);

                Assert.AreEqual(cell.Text, "TextInColumn");
            }
        }

        [Test]
        public void GetTablesCount()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var tables = new WordDocumentTables(doc);

                Assert.AreNotEqual(tables.Count, 0);
            }
        }

        [Test]
        public void TableExists()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var tables = new WordDocumentTables(doc);

                Assert.IsTrue(tables.Exists("Table1"));
            }
        }

        [Test]
        public void CanSelectTable()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var tables = new WordDocumentTables(doc);
                var tlb = tables.Item("Table1");

                Assert.IsNotNull(tlb);
            }
        }

        [Test]
        public void SetCellToEmptyCell()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table2");
                var cell = table.Cell(2, 1);
                Assert.IsTrue(string.IsNullOrEmpty(cell.Text));
                const string val = "New Text";

                cell.Text = val;

                Assert.AreEqual(cell.Text, val);
            }
        }

        [Test]
        public void AddAndCopyRows()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var table = bookmarks.Item("FOOTER_CHIEFS_TABLE").Table;
                var rows = table.Rows;
                const string val = "Директор";

                table.Cell(1, 2).Text = val;
                rows.AddEmpty();
                rows.AddEmpty();
                rows.CopyRow(1);

                Assert.AreEqual(table.Cell(6, 2).Text, val);
            }
        }

        [Test]
        public void FindTableInFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("FOOTER_TABLE");
                var table = bkm.Table;
                const string val = "Составитель";

                table.Cell(1, 1).Text = val;

                Assert.AreEqual(table.Cell(1, 1).Text, val);
            }
        }

        [Test]
        public void FindCellTextInFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("FOOTER_TABLE");
                var table = bkm.Table;

                Assert.AreEqual(table.Cell(1, 1).Text, "составил");
            }
        }
    }
}