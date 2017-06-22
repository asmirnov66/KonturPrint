using System.Text;
using ActiveMockLibrary;
using DocumentFormat.OpenXml.Packaging;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.Bookmarks;
using KonturPrint.PrintObjects.HeadersFooters.Footers;
using KonturPrint.PrintObjects.HeadersFooters.Headers;
using KonturPrintService;
using Ninject;
using NUnit.Framework;
using SKBS;
using SKGENERALLib;
using XMachine;

namespace KonturPrintTest
{
    [TestFixture]
    public class PrintServiceTests
    {
        private string TestDirectory { get; set; }
        private IPrintService KonturPrintService { get; set; }

        private static readonly IKernel Instance = new StandardKernel();
        private string tmpPath;
        private string templatePath;

        [OneTimeSetUp]
        public void SetUp()
        {
            Instance.Bind<IXMachine>().To<XMachineMock>();
            Instance.Load(new KonturPrintServiceModule(Instance.Get<IXMachine>()));
            KonturPrintService = Instance.Get<IPrintService>();
            TestDirectory = TestContext.CurrentContext.TestDirectory;
            tmpPath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestResult.docx");
            templatePath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.docx");
        }

        [Test]
        public void XMachineEvaluate()
        {
            var xMachine = Instance.Get<IXMachine>();
            object block = "return;";

            var res = xMachine.Evaluate(block);

            Assert.IsNotNull(res);
            Assert.IsTrue(bool.Parse(res.ToString()));
        }

        [Test]
        public void PrintWordByXMachine()
        {
            const string bookmarkName = "order_number";
            const string partName = "Main";

            IBSDataObject bo = new BsDataObjectMock("TestBo");
            var parts = bo.Parts as PartsMock;
            parts.AddItem(partName, new SKRecordsetMock());
            var ndoc = "4532";
            ((SKRecordsetMock)parts.Item(partName)).AddNew(new object[] { "ndoc" }, new object[] { ndoc });

            var printParams = new Params();
            printParams.SetParams("BookmarkName", bookmarkName);
            printParams.SetParams("Parts", parts);
            printParams.SetParams("PartName", partName);
            printParams.SetParams("FieldName", "ndoc");
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];");
            printScript.Append("var p = new ParamsMock(args[3]);");
            printScript.Append("var parts = (PartsMock) p.GetValue(\"parts\", (PartsMock) null);");
            printScript.Append("var bkmName = (string) p.GetValue(\"BookmarkName\", \"\");");
            printScript.Append("var partName = (string) p.GetValue(\"PartName\", \"\");");
            printScript.Append("var fldName = (string) p.GetValue(\"FieldName\", \"\");");
            printScript.Append("var bkms = (IWordDocumentBookmarks)  wordDoc.Bookmarks;");
            printScript.Append("bkms.Item(bkmName).Text = (string) parts.Item(partName).Fields[fldName].Value;");
            printParams.SetParams("PrintScript", printScript.ToString());

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                Assert.AreEqual(bookmarks.Item(bookmarkName).Text, ndoc);
            }
        }

        [Test]
        public void FillWordTable()
        {
            const string bookmarkName = "fondsalary_table";
            const string partName = "Main";

            IBSDataObject bo = new BsDataObjectMock("TestBo");
            var parts = bo.Parts as PartsMock;
            parts.AddItem(partName, new SKRecordsetMock());
            var part = parts.Item(partName);
            part.AddNew(new object[] { "Field1", "Field2", "Field3" }, new object[] { "оклад", "руб.", "1000" });
            part.AddNew(new object[] { "Field1", "Field2", "Field3" }, new object[] { "премия", "руб.", "2000" });
            part.AddNew(new object[] { "Field1", "Field2", "Field3" }, new object[] { "отпускные", "руб.", "3000" });

            var printParams = new Params();
            printParams.SetParams("BookmarkName", bookmarkName);
            printParams.SetParams("Parts", parts);
            printParams.SetParams("PartName", partName);
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];\r\n");
            printScript.Append("var p = new ParamsMock(args[3]);\r\n");
            printScript.Append("var parts = (PartsMock) p.GetValue(\"parts\", (PartsMock) null);\r\n");
            printScript.Append("var bkmName = (string) p.GetValue(\"BookmarkName\", \"\");\r\n");
            printScript.Append("var partName = (string) p.GetValue(\"PartName\", \"\");\r\n");
            printScript.Append("var bkms = (IWordDocumentBookmarks)  wordDoc.Bookmarks;\r\n");
            printScript.Append("var tbl = (IWordDocumentTable) bkms.Item(bkmName).Table;\r\n");
            printScript.Append("var part = (SKRecordsetMock) parts.Item(partName);\r\n");
            printScript.Append("part.MoveFirst();\r\n");
            printScript.Append("int i = 3;\r\n");
            printScript.Append("while (!part.EOF) {\r\n");
            printScript.Append("tbl.Rows.Add();\r\n");
            printScript.Append("i++;\r\n");
            printScript.Append("for (int j = 0; j < part.Fields.Count; j++) {\r\n");
            printScript.Append("var cell = tbl.Cell(i, j + 1);\r\n");
            printScript.Append("cell.Text = (string) part.Fields[j].Value;}\r\n");
            printScript.Append("part.MoveNext(); }\r\n");
            printParams.SetParams("PrintScript", printScript.ToString());

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                var tbl = bookmarks.Item(bookmarkName).Table;
                part.MoveFirst();

                Assert.AreEqual(tbl.Cell(4, 1).Text, part.Fields[0].Value);
            }
        }

        [Test]
        public void PrintByPrintScripts()
        {
            const string bookmarkName = "order_number";
            const string bookmarkName1 = "fondsalary_table_col";
            const string text = "123";
            const string text1 = "456";

            IBSDataObject bo = new BsDataObjectMock("TestBo");

            var printParams = new Params();
            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);

            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];");
            printScript.Append("var p = new ParamsMock(args[3]);");
            printScript.Append("var bkms = (IWordDocumentBookmarks)  wordDoc.Bookmarks;");
            printScript.Append($"bkms.Item(\"{bookmarkName}\").Text = \"{text}\";");
            printParams.SetParams("PrintScript", printScript.ToString());

            printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];");
            printScript.Append("var p = new ParamsMock(args[3]);");
            printScript.Append("var bkms = (IWordDocumentBookmarks)  wordDoc.Bookmarks;");
            printScript.Append($"bkms.Item(\"{bookmarkName1}\").Text = \"{text1}\";");
            printService.PrintScripts = new[] { printScript.ToString() };

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                Assert.AreEqual(bookmarks.Item(bookmarkName).Text, text);
                Assert.AreEqual(bookmarks.Item(bookmarkName1).Text, text1);
            }
        }

        [Test]
        public void ChangeFooter()
        {
            const string bookmarkName = "footer_table1";
            const string val = "Пользователь";

            IBSDataObject bo = new BsDataObjectMock("TestBo");

            var printParams = new Params();
            printParams.SetParams("BookmarkName", bookmarkName);
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];");
            printScript.Append("var p = new ParamsMock(args[3]);");
            printScript.Append("var bkmName = (string) p.GetValue(\"BookmarkName\", \"\");");
            printScript.Append("var bkms = (IWordDocumentBookmarks)  wordDoc.Bookmarks;");
            printScript.Append($"bkms.Item(bkmName).Text = \"{val}\";");

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);
            printService.PrintScripts = new[] { printScript.ToString() };

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var bookmarks = new WordDocumentBookmarks(doc);
                Assert.AreEqual(bookmarks.Item(bookmarkName).Text, val);
            }
        }

        [Test]
        public void FillFooters()
        {
            const string bookmarkName = "footer_table1";
            const string val = "Пользователь";

            IBSDataObject bo = new BsDataObjectMock("TestBo");

            var printParams = new Params();
            printParams.SetParams("BookmarkName", bookmarkName);
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];\r\n");
            printScript.Append("var p = new ParamsMock(args[3]);\r\n");
            printScript.Append("var footers = (IWordDocumentHeadersFooters)  wordDoc.Footers;\r\n");
            printScript.Append("var firstFooter = footers.Item(1);\r\n");
            printScript.Append("var table = firstFooter.Tables.Item(1);\r\n");
            printScript.Append($"table.Cell(1, 1).Text = \"{val}\";\r\n");
            printScript.Append("var pageEnumerator = wordDoc.PageEnumerator;\r\n");
            printScript.Append("for (int i = 2; i <= footers.Count; i++) {\r\n");
            printScript.Append("var f = (IPrintObject) footers.Item(i);\r\n");
            printScript.Append("var t = f.GetCopyOf(table as IPrintObject);\r\n");
            printScript.Append("var ft = footers.Item(i).Tables.Item(1);\r\n");
            printScript.Append("var cell = (IPrintObject) ft.Cell(1, 2);\r\n");
            printScript.Append("cell.GetCopyOf(pageEnumerator);\r\n");
            printScript.Append("};\r\n");

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);
            printService.PrintScripts = new[] { printScript.ToString() };

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var footers = new WordDocumentFooters(doc);
                var footer = footers.Item(1);
                var table = footer.Tables.Item(1);
                Assert.AreEqual(table.Cell(1, 1).Text, val);

                footer = footers.Item(2);
                table = footer.Tables.Item(1);
                Assert.AreEqual(table.Cell(1, 1).Text, val);
                Assert.IsTrue(table.Cell(1, 2).Text.Contains("стр."));
            }
        }

        [Test]
        public void FillHeadersAndFooters()
        {
            const string bookmarkName = "footer_table1";
            const string val = "Пользователь";

            IBSDataObject bo = new BsDataObjectMock("TestBo");

            var printParams = new Params();
            printParams.SetParams("BookmarkName", bookmarkName);
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];\r\n");
            printScript.Append("var p = new ParamsMock(args[3]);\r\n");
            printScript.Append("var pageEnumerator = wordDoc.PageEnumerator;\r\n");
            printScript.Append("var footers = (IWordDocumentHeadersFooters) wordDoc.Footers;\r\n");
            printScript.Append("var f = footers.Item(1);\r\n");
            printScript.Append("var table = f.Tables.Item(1);\r\n");
            printScript.Append($"table.Cell(1, 1).Text = \"{val}\";\r\n");
            printScript.Append("var headers = (IWordDocumentHeadersFooters)  wordDoc.Headers;\r\n");
            printScript.Append("headers.Create(2);\r\n");
            printScript.Append("var h = (IPrintObject) headers.Item(1);\r\n");
            printScript.Append("pageEnumerator.CopyTo(h);\r\n");

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);
            printService.PrintScripts = new[] { printScript.ToString() };

            Assert.IsTrue(printService.Print((int)PrintDocumentType.WordTemplate, printServiceParams));
            using (var doc = WordprocessingDocument.Open(tmpPath, false))
            {
                var footers = new WordDocumentFooters(doc);
                var footer = footers.Item(1);
                var table = footer.Tables.Item(1);
                Assert.AreEqual(table.Cell(1, 1).Text, val);

                var headers = new WordDocumentHeaders(doc);
                Assert.IsNotNull(headers.Item(1));
                Assert.IsNotNull(headers.Item(2));
            }
        }

        [Test]
        public void GetPrintDocument()
        {
            IBSDataObject bo = new BsDataObjectMock("TestBo");

            var printParams = new Params();
            var printScript = new StringBuilder();
            printScript.Append("var wordDoc = (IWordDocument) args[1];\r\n");

            var printService = KonturPrintService;
            var printServiceParams = new Params();
            printServiceParams.SetParams("Bo", bo);
            printServiceParams.SetParams("TemplatePath", templatePath);
            printServiceParams.SetParams("FileName", tmpPath);
            printServiceParams.SetParams("ErrorStr", "");
            printServiceParams.SetParams("Params", printParams);

            var printDocument = (IWordDocument)printService.GetPrintDocument((int)PrintDocumentType.WordTemplate, printServiceParams);

            Assert.IsNotNull(printDocument);
            Assert.IsNotNull(printDocument.Bookmarks);
        }

    }
}
