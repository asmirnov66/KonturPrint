using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.HeadersFooters.Footers;
using KonturPrint.PrintObjects.HeadersFooters.Headers;
using KonturPrint.PrintObjects.PageEnumerators;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class WordDocumentHeadersTest
    {
        private string TestDirectory { get; set; }
        private string TemplatePath { get; set; }
        private string TemplateHFPath { get; set; }
        private string TemplatePFPath { get; set; }

        [OneTimeSetUp]
        public void SetUp()
        {
            TestDirectory = TestContext.CurrentContext.TestDirectory;
            TemplatePath = Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.docx");
            TemplateHFPath = Path.Combine(TestDirectory, @"../../TestTemplate/TestHeadersFooters.docx");
            TemplatePFPath = Path.Combine(TestDirectory, @"../../TestTemplate/TestPF.docx");
        }

        [Test]
        public void InsertPageEnumerator()
        {
            using (var memStream = new MemoryStream())
            {
                var byteArray = File.ReadAllBytes(TemplatePath);
                memStream.Write(byteArray, 0, byteArray.Length);
                using (var doc = WordprocessingDocument.Open(memStream, true, new OpenSettings { AutoSave = false }))
                // using (var doc = WordprocessingDocument.Open(TemplatePath, true, new OpenSettings { AutoSave = true }))
                {
                    var headers = new WordDocumentHeaders(doc);
                    var pageEnumerator = new WordDocumentPageEnumerator();
                    var cnt = headers.Count;

                    Assert.AreEqual(cnt, 0);

                    headers.Create((int)HeadersFootersType.Default);
                    cnt = headers.Count;

                    Assert.AreNotEqual(cnt, 0);

                    for (var i = 1; i <= cnt; i++)
                    {
                        var printHeader = (IPrintObject)headers.Item(i);
                        pageEnumerator.CopyTo(printHeader);
                        var block = headers.Item(i).XmlElement.GetFirstChild<SdtBlock>();
                        var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                        Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
                    }
                }
            }
        }

        [Test]
        public void InsertFistDefaultHeaders()
        {
            using (var memStream = new MemoryStream())
            {
                var byteArray = File.ReadAllBytes(TemplateHFPath);
                memStream.Write(byteArray, 0, byteArray.Length);
                using (var doc = WordprocessingDocument.Open(memStream, true, new OpenSettings { AutoSave = false }))
                //using (var doc = WordprocessingDocument.Open(TemplateHFPath, true, new OpenSettings { AutoSave = true }))

                {
                    var headers = new WordDocumentHeaders(doc);
                    var pageEnumerator = new WordDocumentPageEnumerator();

                    headers.Create((int)HeadersFootersType.FirstDefault);
                    var fh = headers.Item(1);
                    fh.XmlElement.Append(new Paragraph(new Run(new Text { Text = "First" })));
                    var printHeader = (IPrintObject)headers.Item(2);
                    pageEnumerator.CopyTo(printHeader);
                    var block = headers.Item(2).XmlElement.GetFirstChild<SdtBlock>();
                    var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                    Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
                }
            }
        }

        [Test]
        public void InsertEvenOddFooters()
        {
            using (var memStream = new MemoryStream())
            {
                var byteArray = File.ReadAllBytes(TemplateHFPath);
                memStream.Write(byteArray, 0, byteArray.Length);
                using (var doc = WordprocessingDocument.Open(memStream, true, new OpenSettings { AutoSave = false }))
                //using (var doc = WordprocessingDocument.Open(TemplateHFPath, true, new OpenSettings { AutoSave = true }))
                {
                    var footers = new WordDocumentFooters(doc);
                    var pageEnumerator = new WordDocumentPageEnumerator();

                    footers.Create((int)HeadersFootersType.EvenOdd);
                    var fh = footers.Item(1);
                    fh.XmlElement.Append(new Paragraph(new Run(new Text { Text = "Even" })));

                    var printFooter = (IPrintObject)footers.Item(2);
                    pageEnumerator.CopyTo(printFooter);
                    var block = footers.Item(2).XmlElement.GetFirstChild<SdtBlock>();
                    var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                    Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
                }
            }
        }

        [Test]
        public void InsertFirstEvenOddFooters()
        {
            using (var memStream = new MemoryStream())
            {
                var byteArray = File.ReadAllBytes(TemplateHFPath);
                memStream.Write(byteArray, 0, byteArray.Length);
                using (var doc = WordprocessingDocument.Open(memStream, true, new OpenSettings { AutoSave = false }))
                //using (var doc = WordprocessingDocument.Open(TemplateHFPath, true, new OpenSettings { AutoSave = true }))
                {
                    var footers = new WordDocumentFooters(doc);
                    var pageEnumerator = new WordDocumentPageEnumerator();

                    footers.Create((int)HeadersFootersType.FirstEvenOdd);
                    var h = footers.Item(1);
                    h.XmlElement.Append(new Paragraph(new Run(new Text { Text = "First" })));

                    h = footers.Item(2);
                    h.XmlElement.Append(new Paragraph(new Run(new Text { Text = "Even" })));

                    var printFooter = (IPrintObject)footers.Item(3);
                    pageEnumerator.CopyTo(printFooter);
                    var block = footers.Item(3).XmlElement.GetFirstChild<SdtBlock>();
                    var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                    Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
                }
            }
        }

        [Test]
        public void InsertHeadersWhenExistsFooters()
        {
            using (var doc = WordprocessingDocument.Open(TemplatePFPath, true, new OpenSettings { AutoSave = false }))
            {
                var headers = new WordDocumentHeaders(doc);
                var pageEnumerator = new WordDocumentPageEnumerator();
                var footers = new WordDocumentFooters(doc);
                var footerTable = footers.Item(1).Tables.Item(1);
                footerTable.Cell(1, 1).Text = "составил";
                footerTable.Cell(1, 2).Text = "продукт";

                Assert.AreEqual(footerTable.Cell(1, 1).Text, "составил");

                headers.Create((int)HeadersFootersType.FirstDefault);
                var fh = headers.ItemByType((int)HeaderFooterValues.First);
                fh.XmlElement.Append(new Paragraph(new Run(new Text { Text = "First" })));
                var printHeader = (IPrintObject)headers.ItemByType((int)HeaderFooterValues.Default);
                pageEnumerator.CopyTo(printHeader);
                var block = printHeader.XmlElement.GetFirstChild<SdtBlock>();
                var para = block.GetFirstChild<SdtContentBlock>().GetFirstChild<Paragraph>();

                Assert.AreEqual(para.ChildElements[3].FirstChild.InnerText, "PAGE");
            }
        }
    }
}