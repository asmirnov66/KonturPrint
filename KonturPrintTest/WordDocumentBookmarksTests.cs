using DocumentFormat.OpenXml.Packaging;
using KonturPrint.PrintObjects.Bookmarks;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentBookmarksTests
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
        public void GetBookmarksCount()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);

                Assert.AreNotEqual(bookmarks.Count, 0);
            }
        }

        [Test]
        public void BookmarkExists()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);

                Assert.IsTrue(bookmarks.Exists("order_number"));
            }
        }

        [Test]
        public void CanSelectBookmark()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("order_number");

                Assert.IsNotNull(bkm);
            }
        }

        [Test]
        public void GetBookmarkName()
        {
            var templatePath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.docx");
            using (var doc = WordprocessingDocument.Open(templatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("order_number");

                Assert.AreEqual(bkm.Name, "order_number");
            }
        }

        [Test]
        public void GetBookmarkText()
        {
            var templatePath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.docx");
            using (var doc = WordprocessingDocument.Open(templatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("order_caption");

                Assert.AreEqual(bkm.Text, "ПРИКАЗ");
            }
        }

        [Test]
        public void GetBookmarkColumnText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_col");

                Assert.AreEqual(bkm.Text, "TextInColumn");
            }
        }

        [Test]
        public void GetMultiLineText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("multiline_bkm");

                Assert.AreEqual(bkm.Text, "Несколько\r\nСтрок\r\nТекста");
            }
        }

        [Test]
        public void SetBookmarkTextInColumn()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("order_caption");
                const string val = "НОВЫЙ Приказ";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void SetBookmarkText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("text_not_cell");
                const string val = "НОВЫЙ Текст";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void ReplaceBookmarkColumnText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_col");
                const string val = "New Text";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void SetBookmarkColumnText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table");
                const string val = "Text";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void SetBookmarkMultiLineText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("multiline_bkm");
                const string val = "Другие\r\nНесколько\r\nстрок";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void SetBookmarkMultiLineTextInColumn()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("order_caption");
                const string val = "НОВЫЙ\r\nПриказ";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void SetBookmarkTextByReplaceTable()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_bkm");
                const string val = "Вместо таблицы";

                bkm.Text = val;

                Assert.AreEqual(bkm.Text, val);
            }
        }

        [Test]
        public void GetTableByBookmark()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_bkm");

                Assert.IsNotNull(bkm.Table);
            }
        }

        [Test]
        public void GetTextByBookmarkInColumn()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_col");
                var cell = bkm.Table.Cell(3, 2);

                Assert.AreEqual(cell.Text, "TextInColumn");
            }
        }

        [Test]
        public void GetTableByBookmarkInCell()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table");

                Assert.IsNotNull(bkm.Table);
            }
        }

        [Test]
        public void GetBookmarkTableText()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("fondsalary_table_bkm");

                Assert.IsFalse(string.IsNullOrEmpty(bkm.Text));
            }
        }

        [Test]
        public void CanGetFoundBookmark()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("FOOTER_TABLE");
                var bkm1 = bookmarks.Item("FOOTER_TABLE");

                Assert.AreEqual(bkm, bkm1);
            }
        }

        [Test]
        public void ChangePartOfColumn()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var bkm = bookmarks.Item("bkm_part_para");

                bkm.Text = "Болен";

                Assert.AreEqual(bkm.Text, "Годен\r\nБолен\r\nМолод");
            }
        }
    }
}
