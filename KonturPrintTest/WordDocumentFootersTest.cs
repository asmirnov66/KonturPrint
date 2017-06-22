using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.HeadersFooters.Footers;
using KonturPrint.PrintObjects.PageEnumerators;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentFootersTest
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
        public void GetFootersCount()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var footers = new WordDocumentFooters(doc);

                Assert.AreNotEqual(footers.Count, 0);
            }
        }

        [Test]
        public void GetFirstFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var footer = new WordDocumentFooter(doc);

                Assert.IsNotNull(footer);
                Assert.IsNotNull(footer.Tables);
                Assert.AreNotEqual(footer.Tables.Count, 0);
            }
        }

        [Test]
        public void GetFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var footers = new WordDocumentFooters(doc);

                Assert.AreNotEqual(footers.Item(2), 0);
            }
        }

        [Test]
        public void GetTableFromFooter()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, false))
            {
                var footers = new WordDocumentFooters(doc);
                var footer = footers.Item(1);

                Assert.IsNotNull(footer.Tables);
            }
        }

        [Test]
        public void InsertPageEnumerator()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = false }))
            {
                var footers = new WordDocumentFooters(doc);
                var pageEnumerator = new WordDocumentPageEnumerator();

                Assert.AreNotEqual(footers.Count, 0);

                for (var i = 1; i <= footers.Count; i++)
                {
                    var printFooter = (IPrintObject)footers.Item(i);

                    pageEnumerator.CopyTo(printFooter);
                    var block = footers.Item(i).XmlElement.GetFirstChild<SdtBlock>();
                    var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                    Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
                }
            }
        }
    }
}