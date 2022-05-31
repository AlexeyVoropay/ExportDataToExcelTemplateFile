namespace ExcelExport.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class WorkSheetHelper
    {
        public static Worksheet GetWorksheetByName(SpreadsheetDocument document, string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            if (sheets.Count() == 0)
                return null;
            else
                return worksheetPart.Worksheet;
        }
    }
}
