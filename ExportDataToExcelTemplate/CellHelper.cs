using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDataToExcelTemplate
{
    public static class CellHelper
    {
        public static StringValue GetCellReference(Cell cell, UInt32Value rowIndex)
        {
            var cellValue = cell.CellReference.Value;
            return new StringValue(cellValue.Replace(Regex.Replace(cellValue, @"[^\d]+", ""), rowIndex.ToString()));
        }
        public static string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            if (cell == null)
                return null;
            var value = cell.InnerText;
            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    if (stringTable != null)
                    {
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;
            }
            return value;
        }
    }
}
