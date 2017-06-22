using System;
using System.Linq;
using ActiveMockLibrary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using KonturPrint.Interfaces;
using Ninject;
using NUnit.Framework;

namespace KonturPrintTest
{
    [TestFixture]
    public class ExcelTests
    {
        private IPrintDocumentFactory PrintDocumentFactory { get; set; }
        private string TestDirectory { get; set; }

        [OneTimeSetUp]
        public void SetUp()
        {
            NinjectKernel.Bind();
            PrintDocumentFactory = NinjectKernel.Instance.Get<IPrintDocumentFactory>();
            TestDirectory = TestContext.CurrentContext.TestDirectory;
        }

        public static string GetCellValue(SpreadsheetDocument document, string sheetName, string addressName)
        {
            string value = null;
            var wbPart = document.WorkbookPart;
            var theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            if (theSheet == null)
            {
                throw new ArgumentException(sheetName);
            }
            var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            var theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            if (theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                    }
                }
            }
            return value;
        }

        [Test]
        public void TestBaseTemplate()
        {
            SKBS.IBSDataObject bo = new BsDataObjectMock();

            var parts = bo.Parts as PartsMock;
            parts.AddItem("Main", new SKRecordsetMock());
            parts.AddItem("Rekv", new SKRecordsetMock());
            parts.AddItem("FuelMove", new SKRecordsetMock());

            var tmpPath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestResult.xlsx");
            const string sheet1 = "WorkSheet1";
            const string sheet2 = "Test2";
            var c1 = 3;
            var c2 = 5;
            var okpo = 354;
            var seria = "4124t";
            var ndoc = "4532";
            ((SKRecordsetMock)parts.Item("Main")).AddNew(new object[] { "Seria", "Count", "ndoc" }, new object[] { seria, c1, ndoc });
            ((SKRecordsetMock)parts.Item("Rekv")).AddNew(new object[] { "OKPO", "Count" }, new object[] { okpo, c2 });
            var fuelTankCode = 1;
            var fuelMoveCorrectNormConsumption = 10;
            ((SKRecordsetMock)parts.Item("FuelMove")).AddNew(new object[] { "NormConsumption", "FuelTankCode" }, new object[] { fuelMoveCorrectNormConsumption, fuelTankCode });
            ((SKRecordsetMock)parts.Item("FuelMove")).AddNew(new object[] { "NormConsumption", "FuelTankCode" }, new object[] { -10, -10 });
            var excelTemplateDocument = PrintDocumentFactory.GetPrintDocument(PrintDocumentType.ExcelTemplate);
            excelTemplateDocument.AddBo("TestBo", bo);
            var templatePath = System.IO.Path.Combine(TestDirectory, @"../../TestTemplate/TestTemplate.xlsx");
            excelTemplateDocument.LoadTemplate(templatePath);
            excelTemplateDocument.ProcessDocument();
            excelTemplateDocument.PrintToPath(tmpPath);
            using (var doc = SpreadsheetDocument.Open(tmpPath, true))
            {
                Assert.AreEqual(seria, GetCellValue(doc, sheet1, "A1"));
                Assert.AreEqual(seria, GetCellValue(doc, sheet1, "A2"));
                Assert.AreEqual(c1.ToString(), GetCellValue(doc, sheet1, "B1"));
                Assert.AreEqual(c2.ToString(), GetCellValue(doc, sheet1, "B2"));
                Assert.AreEqual(okpo.ToString(), GetCellValue(doc, sheet1, "D6"));
                Assert.AreEqual(ndoc, GetCellValue(doc, sheet2, "A1"));
            }
        }
    }
}
