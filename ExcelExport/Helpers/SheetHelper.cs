using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExport.Helpers
{
    public static class SheetHelper
    {
        public static Sheet GetSheet(SpreadsheetDocument document, string sheetName)
        {
            Sheet sheet;
            try
            {
                sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>()
                    .SingleOrDefault(s => s.Name == sheetName);
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
        public static Sheet GetSheetByName(WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            return sheet;
        }

        public static Sheet GetSheetById(WorkbookPart workbookPart, StringValue sheetId)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Id == sheetId).FirstOrDefault();
            return sheet;
        }
    }
}
