using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using KonturPrint.Extensions;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintDocuments
{
    public class ExcelTemplateDocument : BaseDocument
    {
        public override bool IsSameDocumentType(PrintDocumentType type)
        {
            return type == PrintDocumentType.ExcelTemplate;
        }

        public override bool Update()
        {
            return true;
        }

        public override void ProcessDocument()
        {
            CheckForInitialization();
            using (var doc = SpreadsheetDocument.Open(MemStream, true))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                IEnumerable<WorksheetPart> worksheetPart = workbookPart.WorksheetParts;
                workbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                workbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                foreach (WorksheetPart wsp in worksheetPart)
                {
                    ProcessWorkSheetPart(wsp, workbookPart);
                }
            }
        }

        private void ProcessWorkSheetPart(WorksheetPart wsp, WorkbookPart workbookPart)
        {
            var stringTable =
                    workbookPart.GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault();
            var rows = wsp.Worksheet.Descendants<Row>();
            foreach (var row in rows)
            {
                var cells = row.Descendants<Cell>();
                foreach (var cell in cells)
                {
                    var value = cell.InnerText;
                    if (cell.DataType != null)
                    {
                        switch (cell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                if (stringTable != null)
                                {
                                    var element = stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value));
                                    value = element.InnerText;
                                    if (string.IsNullOrEmpty(value)) continue;
                                    if (value.Length > 1)
                                    {
                                        if (value[0] == '#')
                                        {
                                            if (value[1] == 'В')
                                            {

                                            }
                                            else if (value[1] == 'З')
                                            {
                                                var spaceIndex = value.IndexOf(' ');
                                                var expression = value.Substring(spaceIndex + 1);
                                                var values = expression.Split(';');
                                                var boValues = values[0].Split('.');
                                                if (values.Length == 2)
                                                {
                                                    ActiveBo.Parts.Item(boValues[0]).SetFilter(values[1]);
                                                }
                                                if (values.Length == 3)
                                                {
                                                    //TODO: Sort       
                                                }
                                                var field = ActiveBo.Parts.Item(boValues[0]).Fields[boValues[1]];
                                                cell.CellValue = new CellValue(field.DisplayText);
                                                cell.DataType = new EnumValue<CellValues>(field.GetCellValueType());
                                            }
                                            else if (value[1] == 'Б')
                                            {
                                                var boName = value.Split(' ')[1];
                                                SetActiveBo(boName);
                                            }
                                            else
                                            {
                                                var boValues = value.Substring(1).Split('.');
                                                var field = ActiveBo.Parts.Item(boValues[0]).Fields[boValues[1]];
                                                cell.CellValue = new CellValue(field.DisplayText);
                                                cell.DataType = new EnumValue<CellValues>(field.GetCellValueType());
                                            }
                                        }
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }
    }
}
