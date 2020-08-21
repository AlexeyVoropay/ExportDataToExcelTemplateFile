using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcelTemplate.Helpers;
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
        private Row CreateLabel(GeneratingRow item, uint count)
        {
            var row = item.Row;
            row.RowIndex = new UInt32Value(item.Row.RowIndex + (count - 1));
            foreach (var cell in item.Cells)
            {
                cell.Cell.CellReference = Helper.GetCellReference(cell.Cell, row.RowIndex);
                cell.Cell.CellValue = new CellValue(cell.Value);
                cell.Cell.DataType = new EnumValue<CellValues>(CellValues.String);
                row.Append(cell.Cell);
            }
            return row;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataRow item, List<ExportDataToExcelTemplate.Models.Field> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);

            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = Helper.GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.Column + rowIndex))
                {
                    cell.CellValue = new CellValue(item[field._Field].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }


        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataTable table, int tableRowIndex, List<ExportDataToExcelTemplate.Models.Field> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = Helper.GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.Column + rowIndex))
                {
                    cell.CellValue = new CellValue(table.Rows[tableRowIndex][field._Field].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }

        private void OpenFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new Exception(String.Format("Не удалось найти файл \"{0}\"!", filePath));
            }

            var process = Process.Start(filePath);
            if (process != null)
            {
                process.WaitForExit();
            }
        }

        /// <summary>
        /// Подавать только файлы в формате .xlsx
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public System.Data.DataTable ReadFile(String path)
        {
            FileHelper.CheckFile(path);
            return OpenDocumentForRead(path);
        }

        private System.Data.DataTable OpenDocumentForRead(string path)
        {
            System.Data.DataTable data = null;
            using (var document = SpreadsheetDocument.Open(path, false))
            {
                Sheet sheet = Helper.GetSheet(document);
                var relationshipId = sheet.Id.Value;
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                var firstRow = true;
                var columsNames = new List<ColumnName>();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (firstRow)
                    {
                        columsNames.AddRange(GetNames(row, document.WorkbookPart));
                        data = GetTable(columsNames);
                        firstRow = false;
                        continue;
                    }

                    var item = data.NewRow();
                    foreach (var line in columsNames)
                    {
                        var coordinates = String.Format("{0}{1}", line.Liter, row.RowIndex);
                        var cc = row.Elements<Cell>().SingleOrDefault(p => p.CellReference == coordinates);
                        if (cc == null)
                        {
                            throw new Exception(String.Format("Не удалось найти ячейку \"{0}\"!", coordinates));
                        }
                        item[line.Name.Trim()] = GetVal(cc, document.WorkbookPart);

                    }
                    data.Rows.Add(item);
                }
            }
            return data;
        }



        private System.Data.DataTable GetTable(IEnumerable<ColumnName> columsNames)
        {
            var teb = new System.Data.DataTable("ExelTable");

            foreach (var col in columsNames.Select(columnName => new System.Data.DataColumn { DataType = typeof(String), ColumnName = columnName.Name.Trim() }))
            {
                teb.Columns.Add(col);
            }

            return teb;
        }

        private IEnumerable<ColumnName> GetNames(Row row, WorkbookPart wbPart)
        {
            return (from cell in row.Elements<Cell>()
                    where cell != null
                    let
                        text = GetVal(cell, wbPart)
                    where !String.IsNullOrWhiteSpace(text)
                    select
                    new ColumnName(text, Regex.Replace(cell.CellReference.Value, @"[\0-9]", ""))).ToList();
        }


        private string GetVal(Cell cell, WorkbookPart wbPart)
        {
            string value = cell.InnerText;

            if (cell.DataType == null)
            {
                return value;
            }
            switch (cell.DataType.Value)
            {
                case CellValues.SharedString:

                    var stringTable =
                        wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                    }
                    break;
            }

            return value;
        }
    }
}
