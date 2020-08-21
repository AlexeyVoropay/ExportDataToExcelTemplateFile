
namespace ExportDataToExcelTemplate.Helpers
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using ExportDataToExcelTemplate.Models;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    public static class Helper
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

        public static void AddEmptyRows(SheetData sheetData)
        {
            var rows = sheetData.Elements<Row>().ToList();
            var maxRowIndex = rows.Max(x => x.RowIndex);
            for (int i = 0; i < maxRowIndex + 1; i++)
            {
                if (rows.FirstOrDefault(x => x.RowIndex == i) == null)
                {
                    var prevRow = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == i - 1);
                    if (prevRow != null)
                    {
                        prevRow.InsertAfterSelf(new Row { RowIndex = (uint)i });
                    }
                }
            }
        }

        public static List<MergeCells> GetMergeCells(WorkbookPart workbookPart, string sheetId)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                List<MergeCells> mergeCellsList = worksheetPart.Worksheet.Elements<MergeCells>().ToList();
                //foreach (MergeCells mergeCells in mergeCellsList)
                //{
                //    foreach (MergeCell mergeCell in mergeCells)
                //    {
                //        // mergeCell.
                //    }
                //}
                return mergeCellsList;
            }
            return new List<MergeCells>();
        }

        public static bool IsMergeCellNotExists(MergeCells mergeCells, MergeCell mergeCell)
        {
            foreach (MergeCell item in mergeCells)
            {
                var mergeCellReference = new MergeCellReference(item.Reference);
                var cellReference = new CellReference(mergeCell.Reference);
                if (mergeCellReference.CellFrom.RowIndex == cellReference.RowIndex &&
                    mergeCellReference.CellFrom.ColumnIndex == cellReference.ColumnIndex)
                {
                    return false;
                }
            }
            return true;
        }
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
        public static uint GetRowIndex(string cellReferenceValue)
        {
            return Convert.ToUInt32(Regex.Replace(cellReferenceValue, @"[^\d]+", ""));
        }

        public static string GetColumnIndex(string cellReferenceValue)
        {
            return new string(cellReferenceValue.ToCharArray().Where(p => !char.IsDigit(p)).ToArray());
        }

    }

}