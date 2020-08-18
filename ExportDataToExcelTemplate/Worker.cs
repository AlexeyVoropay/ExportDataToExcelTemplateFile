using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcelTemplate.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDataToExcelTemplate
{
    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public partial class Worker
    {
        public void Export(List<System.Data.DataTable> dataTables, List<KeyValuePair<string, string>> fields, String templateName)
        {
            var filePath = FileHelper.CreateFile(templateName);
            OpenForRewriteFile(filePath, dataTables, fields);

            //OpenFile(filePath);
        }

        private void OpenForRewriteFile(String filePath, List<System.Data.DataTable> dataTables, List<KeyValuePair<string, string>> fieldsTable)
        {
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                Sheet sheet = SheetHelper.GetSheet(document);
                var workbookPart = document.WorkbookPart;
                FillFields(workbookPart, sheet.Id.Value, fieldsTable);
                FillTables(workbookPart, sheet.Id.Value, dataTables);

                //var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
                //if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                //{ 
                //    var worksheet = worksheetPart.Worksheet;
                //    var mergeCells = worksheet.Elements<MergeCells>().First();
                //    mergeCells.Append(new MergeCell() { Reference = new StringValue("E8:F8") });
                //    worksheetPart.Worksheet.Save();
                //}
            }
        }

        private void FillFields(WorkbookPart workbookPart, string sheetId, List<KeyValuePair<string, string>> fieldsTable)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Descendants<Cell>())
                {
                    if (cell == null)
                        continue;
                    var cellValue = CellHelper.GetCellValue(cell, workbookPart);
                    if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                        continue;
                    cellValue = cellValue.Substring(2, cellValue.Length - 4);
                    if (fieldsTable.FirstOrDefault(x => cellValue == x.Key) is KeyValuePair<string, string> fieldTable)
                    {
                        if (!String.IsNullOrWhiteSpace(fieldTable.Key))
                        {
                            cell.CellValue = new CellValue(fieldTable.Value);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                    }
                    else
                    {
                        //throw new Exception(String.Format("Нет такого лэйбла \"{0}\"", value));
                    }
                }
            }
        }

        private void FillTables(WorkbookPart workbookPart, string sheetId, List<System.Data.DataTable> dataTables)
        {
            var processedTablesRows = dataTables.ToDictionary(x => x.TableName, y => 0);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            var rows = sheetData.Elements<Row>().ToArray();
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                var row = rows[rowIndex];
                if (!IsRowContainsCellsForFill(row, workbookPart, dataTables.Select(x => x.TableName).ToArray()))
                    continue;
                var fields = GetRowFieldsForFill(row, workbookPart, dataTables.Select(x => x.TableName).ToArray());
                var generatedRowIndex = row.RowIndex;
                if (fields.Any())
                {
                    var tableNamesForAddOneRow = new List<string>();
                    int rowsForProcess = 0;
                    if (!fields.Any(x => x._Field.Contains(":1")))
                    {
                        rowsForProcess = dataTables.Max(x => x.Rows.Count - processedTablesRows[x.TableName]);
                    }
                    else
                    {
                        rowsForProcess = 1;
                        for (int i = 0; i < fields.Count; i++)
                        {
                            fields[i] = new Models.Field(fields[i].Row, fields[i].Column, fields[i]._Field.Replace(":1", ""));
                            tableNamesForAddOneRow.Add(fields[i]._Field.Split('.')[0]);
                        }
                    }
                    for (int i = 0; i < rowsForProcess; i++)
                    {
                        var generatedRow = CreateRow(row, generatedRowIndex, dataTables, i, fields, processedTablesRows);
                        if (i == 0)
                        {
                            row.InsertBeforeSelf(generatedRow);
                        }
                        else
                        {
                            Helper.InsertRow(generatedRowIndex, worksheetPart, generatedRow);
                        }

                        if (row.RowIndex != generatedRowIndex)
                        {
                            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                            {
                                var worksheet = worksheetPart.Worksheet;
                                var mergeCells = worksheet.Elements<MergeCells>().First();
                                var rowMergeCellsList = new List<MergeCell>();
                                var mergeCellChildElements = mergeCells.ChildElements;
                                foreach (MergeCell item in mergeCellChildElements)
                                {
                                    var mergeCellReference = new MergeCellReference(item.Reference);
                                    if (mergeCellReference.CellFrom.RowIndex == row.RowIndex)
                                    {
                                        var cellFrom = mergeCellReference.CellFrom;
                                        var cellTo = mergeCellReference.CellTo;
                                        cellFrom.RowIndex = (int)generatedRowIndex.Value;
                                        cellTo.RowIndex = (int)generatedRowIndex.Value;
                                        var mergeCell = new MergeCell { Reference = $"{cellFrom.Reference}:{cellTo.Reference}" };
                                        mergeCells.Append(mergeCell);
                                    }
                                }
                            }
                        }

                        generatedRowIndex++;
                    }
                    foreach (var tableNameForAddOneRow in tableNamesForAddOneRow.Distinct())
                    {
                        processedTablesRows[tableNameForAddOneRow]++;
                    }
                    row.Remove();
                }
            }           

            #region old
            //foreach (var newRow in footer.Select(item => CreateLabel(item, (UInt32)dataTable.Rows.Count)))
            //{
            //    sheetData.InsertBefore(newRow, rowTemplate);
            //}

            //foreach (var row in sheetData.Elements<Row>())
            //{
            //    if (!IsRowContainsCellsForFill(row, workbookPart, dataTable.TableName))
            //        continue;
            //    var fields = GetRowFieldsForFill(row, workbookPart, dataTable.TableName);
            //    var generatedRowIndex = row.RowIndex;
            //    if (fields.Any())
            //    {                    
            //        var dataTableRowsCount = dataTable.Rows.Count;
            //        for (int i = 0; i < 2; i++)
            //        {
            //            var item = dataTable.Rows[i];
            //            var generatedRow = CreateRow(row, generatedRowIndex, item, fields);
            //            row.InsertBeforeSelf(generatedRow);
            //            generatedRowIndex++;
            //        }
            //        row.Remove();
            //    }
            //}
            //var t1Count = sheetData.Elements<Row>().Count();
            #endregion old
        }

        

        private bool IsRowContainsCellsForFill(Row row, WorkbookPart workbookPart, string[] tableNames)
        {
            foreach (var cell in row.Descendants<Cell>())
            {
                var cellValue = CellHelper.GetCellValue(cell, workbookPart);
                if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                    continue;
                if (!cellValue.StartsWith("{{") || !cellValue.EndsWith("}}"))
                    continue;
                cellValue = cellValue.Substring(2, cellValue.Length - 4);
                foreach (var tableName in tableNames)
                {
                    if (cellValue.IndexOf($"{tableName}.", StringComparison.Ordinal) != -1)
                        return true;
                }
            }
            return false;
        }

        private List<Models.Field> GetRowFieldsForFill(Row rowTemplate, WorkbookPart workbookPart, string[] tableNames)
        {
            var fields = new List<Models.Field>();
            foreach (var cell in rowTemplate.Descendants<Cell>())
            {
                var cellValue = CellHelper.GetCellValue(cell, workbookPart);
                if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                    continue;
                if (!cellValue.StartsWith("{{") || !cellValue.EndsWith("}}"))
                    continue;
                cellValue = cellValue.Substring(2, cellValue.Length - 4);

                foreach (var tableName in tableNames)
                {
                    if (cellValue.IndexOf($"{tableName}.", StringComparison.Ordinal) != -1)
                    {
                        var rowIndex = RowHelper.GetRowIndex(cell.CellReference.Value);
                        var columnIndex = ColumnHelper.GetColumnIndex(cell.CellReference.Value);
                        fields.Add(new Models.Field(rowIndex, columnIndex, cellValue));
                    }
                }                
            }
            return fields;
        }

        
 
        private Row CreateRow(Row rowTemplate, uint rowIndex, List<System.Data.DataTable> tables, int tableRowIndex, List<Models.Field> fields, Dictionary<string, int> processedTablesRows)
        {
            //var newRow = (Row)rowTemplate.Clone();
            var newRow = (Row)rowTemplate.CloneNode(true);
            newRow.RowIndex = rowIndex;
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = CellHelper.GetCellReference(cell, rowIndex);
                foreach (var field in fields.Where(fil => cell.CellReference == fil.Column + rowIndex))
                {
                    var tableName = field._Field.Split('.')[0];
                    var table = tables.FirstOrDefault(x => x.TableName == tableName);
                    var index = tableRowIndex + processedTablesRows[tableName];
                    cell.CellValue = table.Rows.Count > index
                        ? new CellValue(table.Rows[index][field._Field].ToString())
                        : new CellValue(string.Empty);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }
                
    }
}
