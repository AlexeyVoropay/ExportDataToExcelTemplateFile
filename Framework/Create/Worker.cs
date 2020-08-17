using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Framework.Create
{
    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public class Worker
    {
        /// <summary>
        /// путь к папке с шаблонами 
        /// </summary>
        private const String TemplateFolder = @"..\..\Templates\";

        /// <summary>
        /// имя листа шаблона (с которым мы будем работать) 
        /// </summary>
        private const String SheetName = "Лист1";

        /// <summary>
        /// тип документа
        /// </summary>
        private const String FileType = ".xlsx";

        /// <summary>
        /// Папка, для хранения выгруженных файлов
        /// </summary>
        public static String Directory
        {
            get
            {
                const string excelFilesPath = @"C:\xlsx_repository\";
                if (System.IO.Directory.Exists(excelFilesPath) == false)
                {
                    System.IO.Directory.CreateDirectory(excelFilesPath);
                }
                return excelFilesPath;
            }
        }

        public void Export(List<System.Data.DataTable> dataTables, List<KeyValuePair<string, string>> fields, String templateName)
        {
            var filePath = CreateFile(templateName);

            OpenForRewriteFile(filePath, dataTables, fields);

            OpenFile(filePath);
        }

        private String CreateFile(String templateName)
        {
            var templateFelePath = String.Format("{0}{1}{2}", TemplateFolder, templateName, FileType);
            var templateFolderPath = String.Format("{0}{1}", Directory, templateName);
            if (!File.Exists(String.Format("{0}{1}{2}", TemplateFolder, templateName, FileType)))
            {
                throw new Exception(String.Format("Не удалось найти шаблон документа \n\"{0}{1}{2}\"!", TemplateFolder, templateName, FileType));
            }

            //Если в пути шаблона (в templateName) присутствуют папки, то при выгрузке, тоже создаём папки
            var index = (templateFolderPath).LastIndexOf("\\", System.StringComparison.Ordinal);
            if (index > 0)
            {
                var directoryTest = (templateFolderPath).Remove(index, (templateFolderPath).Length - index);
                if (System.IO.Directory.Exists(directoryTest) == false)
                {
                    System.IO.Directory.CreateDirectory(directoryTest);
                }
            }

            var newFilePath = String.Format("{0}_{1}{2}", templateFolderPath, Regex.Replace((DateTime.Now.ToString(CultureInfo.InvariantCulture)), @"[^a-z0-9]+", ""), FileType);
            File.Copy(templateFelePath, newFilePath, true);
            return newFilePath;
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
                    var cellValue = GetCellValue(cell, workbookPart);
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

        private UInt32 GetRowIndex(string cellReferenceValue)
        {
            return Convert.ToUInt32(Regex.Replace(cellReferenceValue, @"[^\d]+", ""));
        }
        private string GetColumnIndex(string cellReferenceValue)
        {
            return new string(cellReferenceValue.ToCharArray().Where(p => !char.IsDigit(p)).ToArray());
        }

        private Sheet GetSheet(SpreadsheetDocument document)
        {
            Sheet sheet;
            try
            {
                sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == SheetName);
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("Возможно в документе существует два листа с названием \"{0}\"!\n", SheetName), ex);
            }
            if (sheet == null)
            {
                throw new Exception(String.Format("В шаблоне не найден \"{0}\"!\n", SheetName));
            }
            return sheet;
        }

        private bool IsRowContainsCellsForFill(Row row, WorkbookPart workbookPart, string[] tableNames)
        {
            foreach (var cell in row.Descendants<Cell>())
            {
                var cellValue = GetCellValue(cell, workbookPart);
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

        private List<Field> GetRowFieldsForFill(Row rowTemplate, WorkbookPart workbookPart, string[] tableNames)
        {
            var fields = new List<Field>();
            foreach (var cell in rowTemplate.Descendants<Cell>())
            {
                var cellValue = GetCellValue(cell, workbookPart);
                if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                    continue;
                if (!cellValue.StartsWith("{{") || !cellValue.EndsWith("}}"))
                    continue;
                cellValue = cellValue.Substring(2, cellValue.Length - 4);

                foreach (var tableName in tableNames)
                {
                    if (cellValue.IndexOf($"{tableName}.", StringComparison.Ordinal) != -1)
                    {
                        var rowId = GetRowIndex(cell.CellReference.Value);
                        var columnId = GetColumnIndex(cell.CellReference.Value);
                        fields.Add(new Field(rowId, columnId, cellValue));
                    }
                }                
            }
            return fields;
        }

        private void FillTables(WorkbookPart workbookPart, string sheetId, List<System.Data.DataTable> dataTables)
        {
            var processedTablesRows = dataTables.ToDictionary(x=>x.TableName, y => 0);
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
                        rowsForProcess = dataTables.Max(x => x.Rows.Count);
                    }
                    else
                    {
                        rowsForProcess = 1;                        
                        for (int i = 0; i < fields.Count; i++)
                        {
                            fields[i] = new Field(fields[i].Row, fields[i].Column, fields[i]._Field.Replace(":1", ""));
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
                            Helper1.InsertRow(generatedRowIndex, worksheetPart, generatedRow);
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
        }

        private void OpenForRewriteFile(String filePath, List<System.Data.DataTable> dataTables, List<KeyValuePair<string, string>> fieldsTable)
        {            
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                Sheet sheet = GetSheet(document);
                var workbookPart = document.WorkbookPart;
                FillFields(workbookPart, sheet.Id.Value, fieldsTable);
                //FillTable(workbookPart, sheet.Id.Value, dataTables.Skip(1).Take(1).FirstOrDefault());
                FillTables(workbookPart, sheet.Id.Value, dataTables);
            }
        }

        private StringValue GetCellReference(Cell cell, UInt32Value rowIndex)
        {
            var cellValue = cell.CellReference.Value;
            return new StringValue(cellValue.Replace(Regex.Replace(cellValue, @"[^\d]+", ""), rowIndex.ToString()));
        }

        private Row CreateLabel(GeneratingRow item, uint count)
        {
            var row = item.Row;
            row.RowIndex = new UInt32Value(item.Row.RowIndex + (count - 1));
            foreach (var cell in item.Cells)
            {
                cell.Cell.CellReference = GetCellReference(cell.Cell, row.RowIndex);
                cell.Cell.CellValue = new CellValue(cell.Value);
                cell.Cell.DataType = new EnumValue<CellValues>(CellValues.String);
                row.Append(cell.Cell);
            }
            return row;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataRow item, List<Field> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);

            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.Column + rowIndex))
                {
                    cell.CellValue = new CellValue(item[field._Field].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataTable table, int tableRowIndex, List<Field> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.Column + rowIndex))
                {
                    cell.CellValue = new CellValue(table.Rows[tableRowIndex][field._Field].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, List<System.Data.DataTable> tables, int tableRowIndex, List<Field> fields, Dictionary<string, int> processedTablesRows)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = GetCellReference(cell, new UInt32Value(rowIndex));
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

        private string GetCellValue(Cell cell, WorkbookPart wbPart)
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
    }
}
