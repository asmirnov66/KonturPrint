using DocumentFormat.OpenXml.Spreadsheet;
using SKBS;

namespace KonturPrint.Extensions
{
    public static class SKFieldExtensions
    {
        public static CellValues GetCellValueType(this SKField skField)
        {
            switch (skField.Type)
            {
                case 3:
                    return CellValues.Number;
                case 200:
                    return CellValues.String;
            }
            return CellValues.String;
        }
    }
}
