using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExport.Helpers;
using ExcelExport.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelExport
{
    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public class ExcelExportNotUsed
    {
        private Row CreateLabel(GeneratingRow item, uint count)
        {
            var row = item.Row;
            row.RowIndex = new UInt32Value(item.Row.RowIndex + (count - 1));
            foreach (var cell in item.Cells)
            {
                cell.Cell.CellReference = CellHelper.GetCellReference(cell.Cell, row.RowIndex);
                cell.Cell.CellValue = new CellValue(cell.Value);
                cell.Cell.DataType = new EnumValue<CellValues>(CellValues.String);
                row.Append(cell.Cell);
            }
            return row;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataRow item, List<LocationWithValue> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);

            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = CellHelper.GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.ColumnIndex + rowIndex))
                {
                    cell.CellValue = new CellValue(item[field.ValueName].ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            return newRow;
        }

        private Row CreateRow(Row rowTemplate, uint rowIndex, System.Data.DataTable table, int tableRowIndex, List<LocationWithValue> fields)
        {
            var newRow = (Row)rowTemplate.Clone();
            newRow.RowIndex = new UInt32Value(rowIndex);
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = CellHelper.GetCellReference(cell, new UInt32Value(rowIndex));
                foreach (var field in fields.Where(fil => cell.CellReference == fil.ColumnIndex + rowIndex))
                {
                    cell.CellValue = new CellValue(table.Rows[tableRowIndex][field.ValueName].ToString());
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
                Sheet sheet = SheetHelper.GetSheet(document, "Лист1");
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

        //private static IEnumerable<Type> GetExcelChartTypes()
        //{
        //    IEnumerable<Type> items = new List<Type>();
        //    try
        //    {
        //        DocumentFormat.OpenXml.Drawing.Charts.LineChart linechart = new DocumentFormat.OpenXml.Drawing.Charts.LineChart();
        //        items = Assembly.GetAssembly(linechart.GetType()).GetTypes().Where(S => S.Name.EndsWith("Chart"));
        //    }
        //    catch
        //    {

        //    }
        //    return items;
        //}

        //public static void ReplaceChartValuesLinearChart(string placeholderCaption, System.Data.DataTable chartData, string newCaption)
        //{
        //    Document wordDocument = null;
        //    MainDocumentPart mainDocumentPart = null;
        //    if (wordDocument != null)
        //    {
        //        try
        //        {
        //            // Get the exact Chart Part where Caption matches the place holder value.
        //            ChartPart target = mainDocumentPart
        //                .ChartParts
        //                .Where(r => r
        //                    .ChartSpace
        //                    .GetFirstChild<Chart>()
        //                    .Title
        //                    .InnerText
        //                    .StartsWith(placeholderCaption)
        //                )
        //                .FirstOrDefault();

        //            if (target != null)
        //            {
        //                // Set the new caption.
        //                target
        //                    .ChartSpace
        //                    .GetFirstChild<Chart>()
        //                    .Title
        //                    .ChartText
        //                    .RichText
        //                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.Paragraph>()
        //                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.Run>()
        //                    .Text
        //                    .Text = newCaption;

        //                // Update all NumberingCache values to reflect total number of records.
        //                foreach (NumberingCache currentNumberingCache in target.ChartSpace.Descendants<NumberingCache>())
        //                {
        //                    currentNumberingCache.PointCount = new PointCount() { Val = (DocumentFormat.OpenXml.UInt32Value)(UInt32)chartData.Rows.Count };
        //                    currentNumberingCache.RemoveAllChildren<NumericPoint>();
        //                }

        //                // Set the Numeric Point values with formats and add to the appropriate NumberingCache.
        //                for (int ctr = 0; ctr < chartData.Rows.Count; ctr++)
        //                {
        //                    // First Range - contains date.
        //                    NumericPoint newNumericPoint = new NumericPoint();
        //                    newNumericPoint.Index = new DocumentFormat.OpenXml.UInt32Value((uint)ctr);
        //                    newNumericPoint.FormatCode = "[$-409]ddmmmyyyy";
        //                    newNumericPoint.NumericValue = new NumericValue(chartData.Rows[ctr][0].ToString());
        //                    target
        //                        .ChartSpace
        //                        .Descendants<NumberingCache>()
        //                        .ToArray()[0]
        //                        .AppendChild(newNumericPoint);

        //                    // Third Range - contains date.
        //                    newNumericPoint = new NumericPoint();
        //                    newNumericPoint.Index = new DocumentFormat.OpenXml.UInt32Value((uint)ctr);
        //                    newNumericPoint.FormatCode = "[$-409]ddmmmyyyy";
        //                    newNumericPoint.NumericValue = new NumericValue(chartData.Rows[ctr][0].ToString());
        //                    target
        //                        .ChartSpace
        //                        .Descendants<NumberingCache>()
        //                        .ToArray()[2]
        //                        .AppendChild(newNumericPoint);

        //                    // Second Range - contains reference data.
        //                    if (chartData.Rows[ctr][2] != DBNull.Value)
        //                    {
        //                        newNumericPoint = new NumericPoint();
        //                        newNumericPoint.Index = new DocumentFormat.OpenXml.UInt32Value((uint)ctr);
        //                        newNumericPoint.FormatCode = "0.00%";
        //                        newNumericPoint.NumericValue = new NumericValue(chartData.Rows[ctr][2].ToString());
        //                        target
        //                            .ChartSpace
        //                            .Descendants<NumberingCache>()
        //                            .ToArray()[1]
        //                            .AppendChild(newNumericPoint);
        //                    }

        //                    // Second Range - contains current data.
        //                    if (chartData.Rows[ctr][3] != DBNull.Value)
        //                    {
        //                        newNumericPoint = new NumericPoint();
        //                        newNumericPoint.Index = new DocumentFormat.OpenXml.UInt32Value((uint)ctr);
        //                        newNumericPoint.FormatCode = "0.00%";
        //                        newNumericPoint.NumericValue = new NumericValue(chartData.Rows[ctr][3].ToString());
        //                        target
        //                            .ChartSpace
        //                            .Descendants<NumberingCache>()
        //                            .ToArray()[3]
        //                            .AppendChild(newNumericPoint);
        //                    }
        //                }

        //                // Update all variable length formula to point to updated number of rows.
        //                foreach (var currentFormula in target.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>())
        //                {
        //                    if (currentFormula.Text.Contains(":"))
        //                    {
        //                        currentFormula.Text =
        //                            currentFormula.Text.Substring(0, currentFormula.Text.LastIndexOf("$") + 1)
        //                            + (chartData.Rows.Count + 1).ToString();
        //                    }
        //                }

        //                // Get handle to ExternalData for accessing embedded Excel document.
        //                ExternalData externalData =
        //                    target
        //                    .ChartSpace
        //                    .Elements<ExternalData>()
        //                    .FirstOrDefault();

        //                if (externalData != null)
        //                {
        //                    // Get handle to Package Part containing excel document.
        //                    EmbeddedPackagePart embeddedPackagePart =
        //                        (EmbeddedPackagePart)
        //                        target
        //                        .Parts
        //                        .Where(r => r.RelationshipId == externalData.Id)
        //                        .FirstOrDefault()
        //                        .OpenXmlPart;

        //                    if (embeddedPackagePart != null)
        //                    {
        //                        // Get handle to Stream for modifying data.
        //                        using (Stream stream = embeddedPackagePart.GetStream())
        //                        {
        //                            // Open Excel for manipulation.
        //                            using (SpreadsheetDocument spreadsheetDocument =
        //                                SpreadsheetDocument.Open(stream, true))
        //                            {
        //                                // Get handle to first sheet.
        //                                DocumentFormat
        //                                    .OpenXml
        //                                    .Spreadsheet
        //                                    .Sheet worksheet = (DocumentFormat.OpenXml.Spreadsheet.Sheet)
        //                                        spreadsheetDocument
        //                                        .WorkbookPart
        //                                        .Workbook
        //                                        .Sheets
        //                                        .FirstOrDefault();

        //                                // Get handle to first worksheet.
        //                                WorksheetPart worksheetPart = (WorksheetPart)
        //                                    spreadsheetDocument
        //                                    .WorkbookPart
        //                                    .Parts
        //                                    .Where(r => r.RelationshipId == worksheet.Id)
        //                                    .FirstOrDefault()
        //                                    .OpenXmlPart;

        //                                // Set Table range on the first worksheet.
        //                                worksheetPart
        //                                    .TableDefinitionParts
        //                                    .FirstOrDefault()
        //                                    .Table
        //                                    .Reference
        //                                    .Value = "A1:D" + (chartData.Rows.Count + 1).ToString();

        //                                // Get handle to access entire sheet data.
        //                                DocumentFormat
        //                                    .OpenXml
        //                                    .Spreadsheet
        //                                    .SheetData sheetData =
        //                                        worksheetPart
        //                                        .Worksheet
        //                                        .Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
        //                                        .FirstOrDefault();

        //                                // Select all data rows.
        //                                var existingRows = sheetData
        //                                    .Elements<DocumentFormat.OpenXml.Spreadsheet.Row>()
        //                                    .Skip(1)
        //                                    .ToArray();

        //                                // Remove all existing data rows.
        //                                for (int ctr = 0; ctr < existingRows.Length; ctr++)
        //                                {
        //                                    sheetData
        //                                        .RemoveChild<DocumentFormat.OpenXml.Spreadsheet.Row>(existingRows[ctr]);
        //                                }

        //                                // Create new rows.
        //                                for (int ctr1 = 0; ctr1 < chartData.Rows.Count; ctr1++)
        //                                {
        //                                    DocumentFormat
        //                                        .OpenXml
        //                                        .Spreadsheet
        //                                        .Row newRecord = new
        //                                            DocumentFormat
        //                                            .OpenXml
        //                                            .Spreadsheet.Row();

        //                                    // Set values and formats for each cell for new row.
        //                                    for (int ctr2 = 0; ctr2 < chartData.Columns.Count; ctr2++)
        //                                    {
        //                                        // Create a new cell.
        //                                        DocumentFormat
        //                                            .OpenXml
        //                                            .Spreadsheet
        //                                            .Cell newCell = new
        //                                                DocumentFormat
        //                                                .OpenXml
        //                                                .Spreadsheet
        //                                                .Cell();

        //                                        // Create a new cell value for holding actual value of the cell.
        //                                        DocumentFormat
        //                                            .OpenXml
        //                                            .Spreadsheet
        //                                            .CellValue newCellValue = new
        //                                                DocumentFormat
        //                                                .OpenXml
        //                                                .Spreadsheet
        //                                                .CellValue();

        //                                        // Set appropriate Style, Data Type and value for the cell.
        //                                        switch (ctr2)
        //                                        {
        //                                            case 0:
        //                                                newCell.StyleIndex = new
        //                                                        DocumentFormat
        //                                                        .OpenXml
        //                                                        .UInt32Value((uint)2);
        //                                                newCellValue.Text =
        //                                                    chartData.Rows[ctr1][ctr2].ToString();
        //                                                break;
        //                                            case 1:

        //                                                newCellValue.Text =
        //                                                    GetSharedStringIndex(
        //                                                            spreadsheetDocument
        //                                                            .WorkbookPart
        //                                                            .SharedStringTablePart,
        //                                                            chartData.Rows[ctr1][ctr2].ToString());
        //                                                newCell.StyleIndex = new
        //                                                    DocumentFormat
        //                                                    .OpenXml
        //                                                    .UInt32Value((uint)2);
        //                                                newCell.DataType =
        //                                                        DocumentFormat
        //                                                        .OpenXml
        //                                                        .Spreadsheet
        //                                                        .CellValues
        //                                                        .SharedString;
        //                                                break;
        //                                            case 2:
        //                                            case 3:
        //                                                newCellValue.Text =
        //                                                    chartData.Rows[ctr1][ctr2].ToString();

        //                                                if (chartData.Rows[ctr1][ctr2] != DBNull.Value &&
        //                                                    chartData.Rows[ctr1][ctr2].ToString().Trim().Length > 0)
        //                                                {
        //                                                    newCell.StyleIndex = new
        //                                                        DocumentFormat
        //                                                        .OpenXml
        //                                                        .UInt32Value((uint)3);
        //                                                }
        //                                                else
        //                                                {
        //                                                    newCell.StyleIndex = new
        //                                                        DocumentFormat
        //                                                        .OpenXml
        //                                                        .UInt32Value((uint)1);
        //                                                }
        //                                                break;
        //                                        }

        //                                        // Append newly created cell value to the cell.
        //                                        newCell.AppendChild(newCellValue);

        //                                        // Append newly created cell to the Row.
        //                                        newRecord.AppendChild(newCell);
        //                                    }

        //                                    // Append newly created row to the Excel sheet.
        //                                    sheetData.AppendChild(newRecord);
        //                                }

        //                                spreadsheetDocument.Save();
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception)
        //        {
        //        }
        //        finally
        //        {
        //            // Save the document.
        //            mainDocumentPart.Document.Save();
        //        }
        //    }
        //}

        //private static void FillChartsOld(WorkbookPart workbookPart, string sheetId, List<ValueToInsert> values)
        //{
        //    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
        //    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        //    //перед этим получите все части диаграммы с помощью
        //    IEnumerable<ChartPart> chartParts = worksheetPart.DrawingsPart.ChartParts;
        //    ChartPart target = chartParts
        //            .Where(r => r
        //                .ChartSpace
        //                .GetFirstChild<Chart>()
        //                .Title
        //                .InnerText
        //                .StartsWith("Распределение НПВ по типам"))
        //            .FirstOrDefault();
        //    // Update all NumberingCache values to reflect total number of records.
        //    var chartDataRowsCount = 3;

        //    var formulas = target.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().ToArray();
        //    for (int i = 0; i < formulas.Count(); i++)
        //    {
        //        formulas[i].Text = "НПВ!$H$7:$H$13";
        //    }
        //    foreach (NumberingCache currentNumberingCache in target.ChartSpace.Descendants<NumberingCache>())
        //    {
        //        var t = currentNumberingCache.Parent.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Formula>();
        //        //currentNumberingCache.Parent.GetFirstChild<Formula>() = "НПВ!$H$7:$H$14";
        //        //t = new DocumentFormat.OpenXml.Drawing.Charts.Formula("НПВ!$H$7:$H$13");
        //        //var y = currentNumberingCache.Parent.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Formula>();
        //        currentNumberingCache.PointCount = new PointCount() { Val = (UInt32Value)(UInt32)chartDataRowsCount };
        //        currentNumberingCache.RemoveAllChildren<NumericPoint>();
        //    }

        //    // First Range - contains date.
        //    NumericPoint newNumericPoint = new NumericPoint();
        //    newNumericPoint.Index = new UInt32Value((uint)0);
        //    //newNumericPoint.FormatCode = "[$-409]ddmmmyyyy";
        //    //newNumericPoint.NumericValue = new NumericValue(DateTime.Parse("10-10-2020").ToOADate().ToString());
        //    newNumericPoint.NumericValue = new NumericValue(13.ToString());
        //    ////////////////////foreach (var item in target.ChartSpace.Descendants<NumberingCache>())
        //    ////////////////////{
        //    ////////////////////    var itemq = item.InnerXml;
        //    ////////////////////}
        //    target
        //        .ChartSpace
        //        .Descendants<NumberingCache>()
        //        .ToArray()[0]
        //        .AppendChild(newNumericPoint);

        //    // Second Range - contains date.
        //    newNumericPoint = new NumericPoint();
        //    newNumericPoint.Index = new UInt32Value((uint)1);
        //    //newNumericPoint.FormatCode = "[$-409]ddmmmyyyy";
        //    //newNumericPoint.NumericValue = new NumericValue(DateTime.Parse("10-10-2020").ToOADate().ToString());
        //    newNumericPoint.NumericValue = new NumericValue(4.ToString());
        //    ////////////////////foreach (var item in target.ChartSpace.Descendants<NumberingCache>())
        //    ////////////////////{
        //    ////////////////////    var itemq = item.InnerXml;
        //    ////////////////////}
        //    target
        //        .ChartSpace
        //        .Descendants<NumberingCache>()
        //        .ToArray()[0]
        //        .AppendChild(newNumericPoint);

        //    // Third Range - contains date.
        //    newNumericPoint = new NumericPoint();
        //    newNumericPoint.Index = new UInt32Value((uint)2);
        //    //newNumericPoint.FormatCode = "[$-409]ddmmmyyyy";
        //    //newNumericPoint.NumericValue = new NumericValue(DateTime.Parse("10-10-2020").ToOADate().ToString());
        //    newNumericPoint.NumericValue = new NumericValue(1.ToString());
        //    ////////////////////foreach (var item in target.ChartSpace.Descendants<NumberingCache>())
        //    ////////////////////{
        //    ////////////////////    var itemq = item.InnerXml;
        //    ////////////////////}
        //    target
        //        .ChartSpace
        //        .Descendants<NumberingCache>()
        //        .ToArray()[0]
        //        .AppendChild(newNumericPoint);


        //    ////////foreach (var chartpart in chartParts)
        //    ////////{
        //    ////////    OpenXmlElementList diagrams = chartpart.ChartSpace.ChildElements;
        //    ////////    foreach (var diagram in diagrams)
        //    ////////    {
        //    ////////        var chart = diagram.GetFirstChild<Chart>();
        //    ////////    }
        //    ////////    //затем получите объект диаграммы с нижеприведенной линией
        //    ////////    //IEnumerable diagrams = chartpart.ChartSpace.ChildElements.OfType();
        //    ////////    //Каждый объект диаграммы будет иметь объект области построения получить это и получить тип объекта с GetType().
        //    ////////    //Тогда, конечно, вы получите имя диаграммы.
        //    ////////}
        //    ////var t = sheetData.Elements<Table>();
        //    //foreach (var row in sheetData.Elements<Row>())
        //    //{
        //    //    foreach (var cell in row.Descendants<Cell>())
        //    //    {
        //    //        if (cell == null)
        //    //            continue;
        //    //        var cellValue = ExcelHelper.GetCellValue(cell, workbookPart);
        //    //        if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
        //    //            continue;
        //    //        if (!cellValue.StartsWith("{{") || !cellValue.EndsWith("}}"))
        //    //            continue;
        //    //        cellValue = cellValue.Substring(2, cellValue.Length - 4);

        //    //        var valueToInsert = values.FirstOrDefault(x => x.FieldName == cellValue);
        //    //        SetCellValues(cell, valueToInsert);
        //    //    }
        //    //}
        //}


        //private static string GetSharedStringIndex(SharedStringTablePart sharedStringTablePart, string valueToSearch)
        //{
        //    int counter = 0;

        //    // Return index if item already exists.
        //    foreach (DocumentFormat.OpenXml.Spreadsheet.SharedStringItem currentItem in sharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>())
        //    {
        //        if (currentItem.InnerText == valueToSearch)
        //        {
        //            return counter.ToString();
        //        }
        //        counter++;
        //    }

        //    // The text does not exist in the part. Create the SharedStringItem and return its index.
        //    sharedStringTablePart.SharedStringTable.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(valueToSearch)));
        //    sharedStringTablePart.SharedStringTable.Save();

        //    return counter.ToString();
        //}

    }
}
