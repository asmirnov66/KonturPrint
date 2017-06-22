using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects;
using KonturPrint.PrintObjects.Bookmarks;
using KonturPrint.PrintObjects.HeadersFooters.Footers;
using KonturPrint.PrintObjects.Tables;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class PrintObjectTest
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
        public void CopyTableToFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var footer = doc.MainDocumentPart.FooterParts
                             .Select(f => f.Footer)
                             .LastOrDefault();

                var source = new PrintObject(table.Table);
                var dest = new PrintObject(footer);

                Assert.AreEqual(footer.Descendants<Table>().Count(), 0);

                source.CopyTo(dest);

                Assert.AreEqual(footer.Descendants<Table>().Count(), 1);
            }
        }

        [Test]
        public void GetCopyOfTableToFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var footer = doc.MainDocumentPart.FooterParts
                             .Select(f => f.Footer)
                             .LastOrDefault();

                var source = new PrintObject(table.Table);
                var dest = new PrintObject(footer);

                Assert.AreEqual(footer.Descendants<Table>().Count(), 0);

                dest.GetCopyOf(source);

                Assert.AreEqual(footer.Descendants<Table>().Count(), 1);
            }
        }

        [Test]
        public void CopyBookmark()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("Table1");
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = new WordDocumentBookmark(doc);
                bkm.Select("order_caption");
                var dest = new PrintObject(table.Table);

                Assert.IsFalse(bookmarks.Exists("order_caption_copy"));

                bkm.CopyTo(dest);

                Assert.IsTrue(bookmarks.Exists("order_caption_copy"));
            }
        }

        [Test]
        public void CopyTableToBookmark()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var table = new WordDocumentTable(doc);
                table.Select("TableForCopy");
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = new WordDocumentBookmark(doc);
                bkm.Select("table_copy");

                Assert.IsNull(bookmarks.Item("table_copy").Table);

                bkm.GetCopyOf(table);

                Assert.IsFalse(string.IsNullOrEmpty(bkm.Text));
                Assert.IsNotNull(bkm.Table);
            }
        }

        [Test]
        public void CopyTableToTable()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var destTable = new WordDocumentTable(doc);
                destTable.Select("TableForCopy");
                var sourceTable = new WordDocumentTable(doc);
                sourceTable.Select("Table2");

                sourceTable.CopyTo(destTable);

                Assert.AreEqual(destTable.Cell(9, 1).Text, "Годен\r\nЗдоров\r\nМолод");
            }
        }

        [Test]
        public void GetCopyOfTableToTable()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var destTable = new WordDocumentTable(doc);
                destTable.Select("TableForCopy");
                var sourceTable = new WordDocumentTable(doc);
                sourceTable.Select("Table2");

                sourceTable.GetCopyOf(destTable);

                Assert.AreEqual(sourceTable.Cell(5, 1).Text, "Вид инструктажа");
            }
        }

        [Test]
        public void FillFooterByFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var footers = new WordDocumentFooters(doc);
                var fistFooter = footers.Item(1);
                var secondFooter = footers.Item(2);
                var table = fistFooter.Tables.Item(1);

                Assert.AreEqual(secondFooter.Tables.Count, 0);

                var destFooter = secondFooter as IPrintObject;
                destFooter.GetCopyOf(table as IPrintObject);
                var destTable = secondFooter.Tables.Item(1);

                Assert.IsNotNull(destTable);
                Assert.AreEqual(destTable.Cell(1, 1).Text, "составил");
            }
        }
    }
}