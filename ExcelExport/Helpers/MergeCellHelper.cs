namespace ExcelExport.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using global::ExcelExport.Models;

    public static class MergeCellHelper
    {
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
                    mergeCellReference.CellFrom.ColumnName == cellReference.ColumnName)
                {
                    return false;
                }
            }
            return true;
        }

        public static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                Worksheet worksheet = WorkSheetHelper.GetWorksheetByName(document, sheetName);
                MergeTwoCells(worksheet, cell1Name, cell2Name);
            }
        }

        public static void MergeTwoCells(Worksheet worksheet, string cell1Name, string cell2Name)
        {
            if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
            {
                return;
            }

            // Добавить код на случай объедения более двух ячеек
            CellHelper.CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
            CellHelper.CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

            MergeCells mergeCells;
            if (worksheet.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = worksheet.Elements<MergeCells>().First();
            }
            else
            {
                mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                }
                else if (worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                }
                else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                }
            }

            // Create the merged cell and append it to the MergeCells collection.
            MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
            mergeCells.Append(mergeCell);

            worksheet.Save();
        }
    }
}
