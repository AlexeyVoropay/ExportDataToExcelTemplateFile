using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace ExportDataToExcelTemplate
{
    public static class SheetHelper
    {
        public static Sheet GetSheet(SpreadsheetDocument document)
        {
            string sheetName = "Лист1";
            Sheet sheet;
            try
            {
                sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == sheetName);
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("Возможно в документе существует два листа с названием \"{0}\"!\n", sheetName), ex);
            }
            if (sheet == null)
            {
                throw new Exception(String.Format("В шаблоне не найден \"{0}\"!\n", sheetName));
            }
            return sheet;
        }
    }
}
