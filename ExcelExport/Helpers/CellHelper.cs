using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExport.Models;

namespace ExcelExport.Helpers
{
    public static class CellHelper
    {
        public static string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }

        public static int GetColumnIndex(string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            int sum = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                sum *= 26;
                sum += (columnLetter[i] - 'A' + 1);
            }
            return sum;
        }

        //public static uint? GetColumnIndex(string cellReference)
        //{
        //    if (string.IsNullOrEmpty(cellReference))
        //    {
        //        return null;
        //    }

        //    //remove digits
        //    string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

        //    int columnNumber = -1;
        //    int mulitplier = 1;

        //    //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
        //    //then multiply that number by our multiplier (which starts at 1)
        //    //multiply our multiplier by 26 as there are 26 letters
        //    foreach (char c in columnReference.ToCharArray().Reverse())
        //    {
        //        columnNumber += mulitplier * ((int)c - 64);

        //        mulitplier = mulitplier * 26;
        //    }

        //    //the result is zero based so return columnnumber + 1 for a 1 based answer
        //    //this will match Excel's COLUMN function
        //    return (uint)columnNumber + 1;
        //}

        // Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
        // When two cells are merged, only the content from one cell is preserved:
        // the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
       

        // Given a Worksheet and a cell name, verifies that the specified cell exists.
        // If it does not exist, creates a new cell. 
        public static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            var rows = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).ToArray();

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                sheetData.Append(row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }
        
        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        public static void InsertText(string docName, string sheetName, string columnName, uint rowIndex, string text)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                InsertText(spreadSheet, sheetName, columnName, rowIndex, text);
            }
        }

        public static void InsertText(SpreadsheetDocument spreadSheet, string sheetName, string columnName, uint rowIndex, string text)
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = SharedStringTablePartHelper.InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            //WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
            Sheet sheet = SheetHelper.GetSheet(spreadSheet, sheetName);
            var relationshipId = sheet.Id.Value;
            var worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(relationshipId);

            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        public static void CopyCellStyle(Worksheet worksheet, string column1, int row1, string column2, int row2)
        {
            var cell1 = CellHelper.GetCell(worksheet, column1, row1);
            var cell2 = CellHelper.GetCell(worksheet, column2, row2);
            cell2.StyleIndex = cell1.StyleIndex;
        }

        public static void CopyCellStyle(Worksheet worksheet, string column1, int row1, Cell cell)
        {
            var cellFrom = CellHelper.GetCell(worksheet, column1, row1);
            cell.StyleIndex = cellFrom.StyleIndex;
        }

        public static Cell GetCell(Worksheet worksheet, string columnName, int rowIndex)
        {
            var row = RowHelper.GetRow(worksheet, (uint)rowIndex);
            if (row == null)
                return null;
            var cells = row.Elements<Cell>();
            if (cells == null || !cells.Any())
                return null;
            return cells.Where(c => c.CellReference.Value == $"{columnName}{rowIndex}")
                .FirstOrDefault();
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
                    var stringTable = wbPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null && int.TryParse(value, out var intValue))
                    {
                        value = stringTable.SharedStringTable.ElementAt(intValue).InnerText;
                    }
                    break;
            }
            return value;
        }

        public static string GetCellValue2(Cell cell, WorkbookPart workbookPart)
        {
            string cellValue = string.Empty;

            if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.SharedString)
                {
                    int id = -1;

                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        SharedStringItem item = SharedStringItemHelper.GetSharedStringItemById(workbookPart, id);

                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }                
            }
            return cellValue;
        }

        public static StringValue GetCellReference(Cell cell, UInt32Value rowIndex)
        {
            var cellValue = cell.CellReference.Value;
            return new StringValue(cellValue.Replace(Regex.Replace(cellValue, @"[^\d]+", ""), rowIndex.ToString()));
        }
    } 
}