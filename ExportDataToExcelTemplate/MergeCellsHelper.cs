using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcelTemplate.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDataToExcelTemplate
{
    public static class MergeCellsHelper
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
                    mergeCellReference.CellFrom.ColumnIndex == cellReference.ColumnIndex)
                {
                    return false;
                }
            }
            return true;
        }
    }
}
