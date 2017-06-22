using DocumentFormat.OpenXml.Packaging;
using KonturPrint.PrintObjects.TableRows;
using KonturPrint.PrintObjects.Tables;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentTableRowsTest
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
        public void GetRowsCount()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table.Table);

                Assert.AreEqual(rows.Count, 3);
            }
        }

        [Test]
        public void CanGetRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table.Table);

                Assert.IsNotNull(rows.Item(1));
            }
        }

        [Test]
        public void CanAddRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table);
                var cnt = rows.Count;

                var newRow = rows.Add();

                Assert.IsNotNull(newRow);
                Assert.AreEqual(cnt + 1, rows.Count);
            }
        }

        [Test]
        public void SetTextInNewRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table);
                rows.Add();
                var cell = table.Cell(4, 1);
                const string val = "New Text";

                cell.Text = val;

                Assert.AreEqual(cell.Text, val);
            }
        }

        [Test]
        public void CanCopyRow()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table);
                var cnt = rows.Count;

                var newRow = rows.Copy();

                Assert.IsNotNull(newRow);
                Assert.AreEqual(cnt + 1, rows.Count);
            }
        }

        [Test]
        public void CanCopyRowByIndex()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var rows = new WordDocumentTableRows(table);

                rows.CopyRow(2);
                var cell = table.Cell(4, 3);

                Assert.AreEqual(cell.Text, "Количество");
            }
        }
    }
}