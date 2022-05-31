namespace ExcelExport
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using DocumentFormat.OpenXml.Drawing.Spreadsheet;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;
    using global::ExcelExport.Helpers;
    using global::ExcelExport.Models;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Создание Excel файла
    /// </summary>
    public static class ExcelExport
    {
        public static void CreateFilledFile(string templatePath, List<SheetExportData> sheetsExportData
            , ILogger logger)
        {
            var filePath = FileHelper.CreateResultFile(templatePath);
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                FillDocument(document, sheetsExportData, logger);
            }
        }

        public static MemoryStream GetFilledMemoryStreamFile(string filePath, List<SheetExportData> sheetsExportData
            , ILogger logger)
        {
            // Copy file content to MemeoryStream via byte array
            MemoryStream stream = new MemoryStream();
            byte[] fileBytesArray = File.ReadAllBytes(filePath);
            stream.Write(fileBytesArray, 0, fileBytesArray.Length); // copy file content to MemoryStream
            stream.Position = 0;
            using (var document = SpreadsheetDocument.Open(stream, true))
            {
                FillDocument(document, sheetsExportData, logger);
            }
            return stream;
        }
        
        public static MemoryStream GetFilledMemoryStreamFile(Stream templateStream, List<SheetExportData> sheetsExportData
            , ILogger logger)
        {
            MemoryStream stream = new MemoryStream();
            byte[] fileBytesArray = new byte[templateStream.Length];
            templateStream.Read(fileBytesArray, 0, fileBytesArray.Length);
            stream.Write(fileBytesArray, 0, fileBytesArray.Length); // copy file content to MemoryStream
            stream.Position = 0;
            using (var document = SpreadsheetDocument.Open(stream, true))
            {
                FillDocument(document, sheetsExportData, logger);
            }
            return stream;
        }

        public static void FillDocument(SpreadsheetDocument document, List<SheetExportData> sheetsData, ILogger logger)
        {
            foreach (var sheetData in sheetsData)
            {
                var sw = new Stopwatch();
                sw.Start();

                FillDocumentSheet(document, sheetData);

                sw.Stop();
                logger.LogDebug($"FillDocument Sheet: {sheetData.SheetName}, Time: {sw.Elapsed.TotalSeconds}");
            }
        }

        public static void AddColumns(Worksheet workSheet2, int columnId, double width)
        {
            // Check if the column collection exists
            Columns cs = workSheet2.Elements<Columns>().FirstOrDefault();

            if ((cs == null))
            {
                // If Columns appended to worksheet after sheetdata Excel will throw an error.
                SheetData sd = workSheet2.Elements<SheetData>().FirstOrDefault();
                if ((sd != null))
                {
                    //cs = workSheet2.InsertBefore(new Columns(), sd);
                    cs = workSheet2.InsertBefore(new Columns(), sd);
                }
                else
                {
                    cs = new Columns();
                    workSheet2.Append(cs);
                }
            }

            //create a column object to define the width of columns 1 to 3  
            Column c = new Column
            {
                Min = (uint)columnId,
                Max = (uint)columnId,
                Width = width,
                CustomWidth = true
            };
            cs.Append(c);
        }

        private static void GenerateTemplate(WorkbookPart workbookPart, string sheetId, List<ColumnBlockToInsert> columnsBlockToInsert)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var lastColumnName = "A";
            foreach (var columnBlockToInsert in columnsBlockToInsert)
            {
                var startColumnIndex = CellHelper.GetColumnIndex(columnBlockToInsert.FirstColumnName);
                for (int i = 0; i < columnBlockToInsert.ColumnsWidths.Length; i++)
                {
                    AddColumns(worksheetPart.Worksheet, startColumnIndex + i, columnBlockToInsert.ColumnsWidths[i]);
                }
                foreach (var blockToInsert in columnBlockToInsert.RowBlocksToInsert)
                {
                    var fromIndex = 0;
                    foreach (var cellToInsert in blockToInsert.CellsToInsert)
                    {
                        var column1Name = CellHelper.ColumnIndexToColumnLetter(startColumnIndex + fromIndex);
                        var row1Index = blockToInsert.RowId;
                        var cell1Reference = new CellReference($"{column1Name}{row1Index}");

                        var column2Name = CellHelper.ColumnIndexToColumnLetter(startColumnIndex + fromIndex + cellToInsert.RowSize - 1);
                        var row2Index = blockToInsert.RowId;
                        var cell2Reference = new CellReference($"{column2Name}{row2Index}");
                        lastColumnName = column2Name;

                        if (cellToInsert.RowSize > 1)
                        {
                            MergeCellHelper.MergeTwoCells(worksheetPart.Worksheet, cell1Reference.Reference, cell2Reference.Reference);
                        }
                        else
                        {
                            CellHelper.CreateSpreadsheetCellIfNotExist(worksheetPart.Worksheet, cell1Reference.Reference);
                        }

                        if (!string.IsNullOrEmpty(cellToInsert.FieldName))
                        {
                            var valueToInsert = new ValueToInsert
                            {
                                FieldName = null,
                                IsFormula = false,
                                Type = typeof(string),
                                Value = cellToInsert.FieldName,
                            };
                            SetCellValues(worksheetPart.Worksheet, column1Name, blockToInsert.RowId, valueToInsert);
                        }

                        CellHelper.CopyCellStyle(worksheetPart.Worksheet,
                            new CellReference(cellToInsert.StyleCellReference).ColumnName,
                            new CellReference(cellToInsert.StyleCellReference).RowIndex,
                            column1Name, row1Index);
                        fromIndex++;
                    }
                }
            }
            MergeCellHelper.MergeTwoCells(worksheetPart.Worksheet, "B3", $"{lastColumnName}3");
        }

        private static void MoveAllCharts(WorkbookPart workbookPart, string sheetId, List<ChartToInsert> chartsToInsert)
        {
            if (chartsToInsert == null || !chartsToInsert.Any())
                return;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var firstChart = chartsToInsert.FirstOrDefault();
            if (firstChart.MoveOnRows > 0 && worksheetPart.DrawingsPart != null)
            {
                foreach (ChartPart chart in worksheetPart.DrawingsPart.ChartParts)
                {
                    var tcas = worksheetPart.DrawingsPart.WorksheetDrawing.Descendants<TwoCellAnchor>();
                    foreach (var tca in tcas)
                    {
                        tca.FromMarker.RowId = new RowId($"{int.Parse(tca.FromMarker.RowId.Text) + firstChart.MoveOnRows}");
                        tca.ToMarker.RowId = new RowId($"{int.Parse(tca.ToMarker.RowId.Text) + firstChart.MoveOnRows}");
                    }
                }
            }
        }

        public static void FillDocumentSheet(SpreadsheetDocument document, SheetExportData sheetExportData)
        {
            Sheet sheet = SheetHelper.GetSheet(document, sheetExportData.SheetName);
            var workbookPart = document.WorkbookPart;
            if (sheetExportData.SheetName == "Крепление")
            {
                GenerateTemplate(workbookPart, sheet.Id.Value, sheetExportData.ColumnsBlockToInsert);
            }
            MoveAllCharts(workbookPart, sheet.Id.Value, sheetExportData.ChartsToInserts);
            if (sheetExportData.SheetName == "КСБ-6.00")
            {
                // Общее количество строк в шаблоне
                var templateRowsCount = sheetExportData.CopyRowIndexTo - sheetExportData.CopyRowIndexFrom;
                for (int i = 0; i < sheetExportData.RowBlocksForCopyAndInsert - 1; i++)
                {
                    var toRowIndex = sheetExportData.CopyRowIndexTo + (templateRowsCount * i);
                    CopyAndInsertRows(sheetExportData.CopyRowIndexFrom, toRowIndex, templateRowsCount, i + 2, 
                            workbookPart, sheet.Id.Value, sheetExportData.ArraysToInserts);
                }
            }
            FillFields(workbookPart, sheet.Id.Value, sheetExportData.FieldsToInserts);
            FillTables(workbookPart, sheet.Id.Value, sheetExportData.ArraysToInserts);
            FillCharts(workbookPart, sheet.Id.Value, sheetExportData.ChartsToInserts);
            HideCharts(workbookPart, sheet.Id.Value, sheetExportData.ChartsToInserts);
            FillChartsXY(workbookPart, sheet.Id.Value, sheetExportData.ChartsXYToInserts);
        }

        private static void HideCharts(WorkbookPart workbookPart, string sheetId, List<ChartToInsert> values)
        {
            if (values == null || !values.Any())
                return;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            foreach (var row in sheetData.Elements<Row>())
            {
                if (values.Any(x => x.IsHide && x.StartRowIndex <= row.RowIndex && row.RowIndex <= x.EndRowIndex))
                    row.Height = 0;
            }
        }

        private static void FillFields(WorkbookPart workbookPart, string sheetId, List<ValueToInsert> values)
        {
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            //var t = sheetData.Elements<Table>();
            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Descendants<Cell>())
                {
                    if (cell == null)
                        continue;
                    var cellValue = CellHelper.GetCellValue(cell, workbookPart);
                    if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                        continue;
                    if (!cellValue.StartsWith("{{") || !cellValue.EndsWith("}}"))
                        continue;
                    cellValue = cellValue.Substring(2, cellValue.Length - 4);

                    var valueToInsert = values.FirstOrDefault(x => x.FieldName == cellValue);
                    if (valueToInsert != null && !string.IsNullOrEmpty(valueToInsert.CellReferenceStyle))
                    {
                        CellHelper.CopyCellStyle(worksheetPart.Worksheet,
                            new CellReference(valueToInsert.CellReferenceStyle).ColumnName,
                            new CellReference(valueToInsert.CellReferenceStyle).RowIndex,
                            cell);
                    }
                    SetCellValues(cell, valueToInsert);
                }
            }
        }

        private static void FillTables(WorkbookPart workbookPart, string sheetId, List<TableToInsert> dataTables)
        {
            var processedTablesRows = dataTables.ToDictionary(x => x.TableName, y => 0);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            ExcelHelper.AddEmptyRows(sheetData);
            var rows = sheetData.Elements<Row>().ToList();
            foreach (var row in rows)
            {
                if (!IsRowContainsCellsForFill(row, workbookPart, dataTables.Select(x => x.TableName).ToArray()))
                    continue;
                var fields = GetRowFieldsForFill(row, workbookPart, dataTables.Select(x => x.TableName).ToArray());
                var generatedRowIndex = row.RowIndex;
                if (fields.Any())
                {
                    var tableNamesForAddOneRow = new List<string>();
                    // Кол-во строк которое надо обработать
                    int rowsForProcess = 0;
                    // Если нет полей, которые надо оставить без изменений
                    if (!fields.Any(x => x.ValueName.Contains(":1")))
                    {
                        // Кол-во строк из задействованных таблиц, которое осталось обработать
                        var maxTableRowsForAdd = 0;
                        foreach (var dataTable in dataTables.Where(x => fields.Any(y => y.ValueName.Split('.')[0] == x.TableName)))
                        {
                            // Если таблица пустая, то указываем, единицу для пустой строки
                            var rowsCount = dataTable.Rows.Count > 0 ? dataTable.Rows.Count : 1;
                            // Оставшиеся количество строк для таблицы, которое осталось обработать
                            var remainingTableRowsForAdd = rowsCount - processedTablesRows[dataTable.TableName];
                            if (maxTableRowsForAdd == 0 &&
                                processedTablesRows[dataTable.TableName] > 0 &&
                                remainingTableRowsForAdd <= 0)
                            {
                                // Добавление пустой строки
                                maxTableRowsForAdd = 1;
                            }
                            else
                            {
                                maxTableRowsForAdd = Math.Max(maxTableRowsForAdd, remainingTableRowsForAdd);
                            }
                        }
                        rowsForProcess = maxTableRowsForAdd;
                    }
                    else
                    {
                        // Т.к. есть поля которые надо оставить без изменений,
                        // то кол-во строк для обработки равно единице
                        rowsForProcess = 1;
                        for (int i = 0; i < fields.Count; i++)
                        {
                            fields[i] = new LocationWithValue(fields[i].RowIndex, fields[i].ColumnIndex, fields[i].ValueName.Replace(":1", ""));
                            // Имена таблиц (списков) для которых надо указать что еще одна строка обработана
                            tableNamesForAddOneRow.Add(fields[i].ValueName.Split('.')[0]);
                        }
                    }
                    // Обработка строк
                    //var merges = new List<MergeCell>();
                    for (int i = 0; i < rowsForProcess; i++)
                    {
                        var generatedRow = CreateRow(worksheetPart.Worksheet, row, generatedRowIndex, dataTables, i, fields, processedTablesRows);
                        if (i == 0)
                        {
                            row.InsertBeforeSelf(generatedRow);
                        }
                        else
                        {
                            InsertRowHelper.InsertRow(generatedRowIndex, worksheetPart, generatedRow);
                        }
                        if (row.RowIndex != generatedRowIndex)
                        {
                            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
                            {
                                // if (!merges.Any())
                                // {
                                var worksheet = worksheetPart.Worksheet;
                                var mergeCells = worksheet.Elements<MergeCells>().First();
                                var rowMergeCellsList = new List<MergeCell>();
                                var mergeCellChildElements = mergeCells.ChildElements;
                                foreach (MergeCell item in mergeCellChildElements)
                                {
                                    var mergeCellReference = new MergeCellReference(item.Reference);
                                    if (mergeCellReference.CellFrom.RowIndex == row.RowIndex)
                                    {
                                        // merges.Add(item);
                                        var cellFrom = mergeCellReference.CellFrom;
                                        var cellTo = mergeCellReference.CellTo;
                                        cellFrom.RowIndex = (int)generatedRowIndex.Value;
                                        cellTo.RowIndex = (int)generatedRowIndex.Value;
                                        var mergeCell = new MergeCell { Reference = $"{cellFrom.Reference}:{cellTo.Reference}" };
                                        mergeCells.Append(mergeCell);
                                    }
                                }
                                // }
                                // else
                                // {
                                //     var worksheet = worksheetPart.Worksheet;
                                //     var mergeCells = worksheet.Elements<MergeCells>().First();
                                //     foreach (var merge in merges)
                                //     {
                                //         var mergeCellReference = new MergeCellReference(merge.Reference);
                                //                                      
                                //         var cellFrom = mergeCellReference.CellFrom;
                                //         var cellTo = mergeCellReference.CellTo;
                                //         cellFrom.RowIndex = (int)generatedRowIndex.Value;
                                //         cellTo.RowIndex = (int)generatedRowIndex.Value;
                                //         var mergeCell = new MergeCell { Reference = $"{cellFrom.Reference}:{cellTo.Reference}" };
                                //         mergeCells.Append(mergeCell);
                                //     }
                                // }
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
        }

        public static void CopyAndInsertRows(int fromRowIndex, int toRowIndex, int rowsCount, int wellIndex,
                WorkbookPart workbookPart, string sheetId, List<TableToInsert> dataTables)
        {
            var processedTablesRows = dataTables.ToDictionary(x => x.TableName, y => 0);
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var mergeCells = MergeCellHelper.GetMergeCells(workbookPart, sheetId).FirstOrDefault()
                    .ChildElements.Select(x => x as MergeCell);

            Row oldRow = null;
            Row generatedRow = null;
            var mergeCellsForCopy = new List<MergeCell>();
            for (int i = 0; i < rowsCount; i++)
            {
                var oldRowIndex = fromRowIndex + i;
                oldRow = RowHelper.GetRow(worksheetPart.Worksheet, (uint)oldRowIndex);
                generatedRow = CreateRow(worksheetPart.Worksheet, oldRow, (uint)(toRowIndex + i), dataTables, 0, new List<LocationWithValue>(), processedTablesRows);
                foreach (var cell in generatedRow.Descendants<Cell>())
                {
                    if (cell == null)
                        continue;
                    var cellValue = CellHelper.GetCellValue(cell, workbookPart);
                    if (String.IsNullOrWhiteSpace(cellValue) || cellValue.Length <= 4)
                        continue;
                    if (!cellValue.StartsWith("{{Well1"))
                        continue;
                    cellValue = cellValue.Replace("{{Well1", $"{{{{Well{wellIndex}");
                    SetCellValues(cell, new ValueToInsert(null, typeof(string), cellValue));
                }
                InsertRowHelper.InsertRow((uint)(toRowIndex + i), worksheetPart, generatedRow);

                var rowMergeCellsForCopy = mergeCells.Where(x => new MergeCellReference(x.Reference).CellFrom.RowIndex == oldRowIndex);
                mergeCellsForCopy.AddRange(rowMergeCellsForCopy);
            }
            foreach (var mergeCellForCopy in mergeCellsForCopy)
            {
                var newMergeCell = MergeCellReferenceMoveByRows(new MergeCellReference(mergeCellForCopy.Reference), toRowIndex - fromRowIndex);
                MergeCellHelper.MergeTwoCells(worksheetPart.Worksheet, newMergeCell.CellFrom.Reference, newMergeCell.CellTo.Reference);
            }
        }

        public static MergeCellReference MergeCellReferenceMoveByRows(MergeCellReference mergeCell, int rowsCount)
        {
            var newFromRowIndex = mergeCell.CellFrom.RowIndex + rowsCount;
            var newToRowIndex = mergeCell.CellTo.RowIndex + rowsCount;
            var newFromCellReference = $"{mergeCell.CellFrom.ColumnName}{newFromRowIndex}";
            var newToCellReference = $"{mergeCell.CellTo.ColumnName}{newToRowIndex}";
            return new MergeCellReference($"{newFromCellReference}:{newToCellReference}");
        }

        private static void FillCharts(WorkbookPart workbookPart, string sheetId, List<ChartToInsert> chartsToInsert)
        {
            if (chartsToInsert == null || !chartsToInsert.Any())
                return;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var chartParts = worksheetPart.DrawingsPart.ChartParts;
            foreach (var chartToInsert in chartsToInsert)
            {
                FillChart(chartParts, chartToInsert);
            }
        }

        private static void FillChart(IEnumerable<ChartPart> chartParts, ChartToInsert chartToInsert)
        {
            //var chartPart = GetChartPartByTitle(chartParts, chartToInsert.ChartTitle);
            var chartPart = chartParts
                .Where(r => r.ChartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>().Title.InnerText.StartsWith(chartToInsert.ChartTitle))
                .FirstOrDefault();

            var formulas = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().ToArray();
            for (int i = 0; i < formulas.Count(); i++)
            {
                var pointCount = chartToInsert.PointCount;
                formulas[i].Text = ChartHelper.ChangeFormula(formulas[i].Text, 0, pointCount);
            }

            if (chartToInsert.DataPointsColors?.Count() > 0)
            {
                var chart = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Chart>()
                    .FirstOrDefault();
                var chartBar = chart.PlotArea.Descendants<BarChart>().FirstOrDefault();
                var chartBarSeries = chartBar.Descendants<BarChartSeries>().FirstOrDefault();
                var dataPoints = chartBarSeries.Descendants<DataPoint>().ToArray();
                for (int i = 0; i < chartToInsert.DataPointsColors.Length; i++)
                {
                    var dataPoint = dataPoints[i];
                    var solidFill = dataPoint.ChartShapeProperties.Descendants<SolidFill>()
                        .FirstOrDefault();
                    solidFill.RgbColorModelHex.Val = chartToInsert.DataPointsColors[i];
                }
            }
        }

        private static void FillChartsXY(WorkbookPart workbookPart, string sheetId, List<ChartXYToInsert> chartsXYToInsert)
        {
            if (chartsXYToInsert == null || !chartsXYToInsert.Any())
                return;
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
            var chartParts = worksheetPart.DrawingsPart.ChartParts;
            foreach (var chartToInsert in chartsXYToInsert)
            {
                FillChartXY(chartParts, chartToInsert);
            }
        }

        private static void FillChartXY(IEnumerable<ChartPart> chartParts, ChartXYToInsert chartXYToInsert)
        {
            var chartPart = ChartHelper.GetChartPartByTitle(chartParts, chartXYToInsert.ChartTitle);

            var scatterChartSeries = chartPart.ChartSpace.Descendants<ScatterChartSeries>();
            foreach (var item in chartXYToInsert.SeriesListToInsert)
            {
                var scatterChartSeriesOne = ChartHelper.GetScatterChartSeriesBySeriesText(scatterChartSeries, item.SeriesText);
                if (scatterChartSeriesOne != null)
                {
                    var xValues = scatterChartSeriesOne.GetFirstChild<XValues>();
                    var yValues = scatterChartSeriesOne.GetFirstChild<YValues>();
                    var xFormula = xValues.GetFirstChild<StringReference>().Formula;
                    var yFormula = yValues.GetFirstChild<NumberReference>().Formula;
                    var pointCount = item.PointCount;
                    xFormula.Text = ChartHelper.ChangeFormula(xFormula.Text, item.PointSkip, pointCount);
                    yFormula.Text = ChartHelper.ChangeFormula(yFormula.Text, item.PointSkip, pointCount);
                }
            }
            if (chartXYToInsert.AxisListToInsert?.Count() > 0)
            {
                var axises = chartPart.ChartSpace.Descendants<ValueAxis>();
                foreach (var item in chartXYToInsert.AxisListToInsert)
                {
                    var axis = axises.FirstOrDefault(x => x.InnerText == item.AxisText);
                    var scaling = axis.GetFirstChild<Scaling>();
                    var minAxisValue = scaling.GetFirstChild<MinAxisValue>();
                    minAxisValue.Val = item.ScalingMinAxisValue;
                    var maxAxisValue = scaling.GetFirstChild<MaxAxisValue>();
                    maxAxisValue.Val = item.ScalingMaxAxisValue;
                    if (item.MajorUnitValue.HasValue)
                    {
                        var majorUnit = axis.GetFirstChild<MajorUnit>();
                        majorUnit.Val = item.MajorUnitValue;
                    }
                }
            }
        }

        private static bool IsRowContainsCellsForFill(Row row, WorkbookPart workbookPart, string[] tableNames)
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

        private static List<LocationWithValue> GetRowFieldsForFill(Row rowTemplate, WorkbookPart workbookPart, string[] tableNames)
        {
            var fields = new List<LocationWithValue>();
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
                        var rowIndex = CellReferenceHelper.GetRowIndex(cell.CellReference.Value);
                        var columnIndex = CellReferenceHelper.GetColumnIndex(cell.CellReference.Value);
                        fields.Add(new LocationWithValue(rowIndex, columnIndex, cellValue));
                    }
                }
            }
            return fields;
        }

        private static Row CreateRow(Worksheet worksheet, Row rowTemplate, uint rowIndex, List<TableToInsert> tables, int tableRowIndex, List<LocationWithValue> fields, Dictionary<string, int> processedTablesRows)
        {
            var newRow = (Row)rowTemplate.CloneNode(true);
            newRow.RowIndex = rowIndex;
            foreach (var cell in newRow.Elements<Cell>())
            {
                cell.CellReference = CellHelper.GetCellReference(cell, rowIndex);
                foreach (var field in fields.Where(fil => cell.CellReference == fil.ColumnIndex + rowIndex))
                {
                    var tableName = field.ValueName.Split('.')[0];
                    var fieldName = field.ValueName.Substring(tableName.Length + 1);
                    var table = tables.FirstOrDefault(x => x.TableName == tableName);
                    var index = tableRowIndex + processedTablesRows[tableName];
                    if (table.Rows.Count > index)
                    {
                        var row = table.Rows.FirstOrDefault(x => x.RowId == index);
                        var valueToInsert = row.Values.FirstOrDefault(x => x.FieldName == fieldName);
                        if (!string.IsNullOrEmpty(valueToInsert.CellReferenceStyle))
                        {
                            CellHelper.CopyCellStyle(worksheet,
                                new CellReference(valueToInsert.CellReferenceStyle).ColumnName,
                                new CellReference(valueToInsert.CellReferenceStyle).RowIndex,
                                cell);
                        }
                        SetCellValues(cell, valueToInsert);
                    }
                    else
                    {
                        cell.CellValue = new CellValue(String.Empty);
                        cell.DataType = CellValues.String;
                    }
                }
            }
            return newRow;
        }

        private static void SetCellValues(Worksheet worksheet, string columnName, int rowId, ValueToInsert value)
        {
            var cell = CellHelper.GetCell(worksheet, columnName, rowId);
            SetCellValues(cell, value);
        }

        private static void SetCellValues(Cell cell, ValueToInsert value)
        {
            if (value != null)
            {
                if (value.IsFormula)
                {
                    cell.CellFormula = new CellFormula();
                    cell.CellFormula.Text = value.Value.ToString();
                    cell.CellFormula.CalculateCell = new BooleanValue(true);
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(string.Empty);
                }
                else if (value?.Type == typeof(int) || value?.Type == typeof(long))
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(value.Value?.ToString());
                }
                else if (value?.Type == typeof(int?) || value?.Type == typeof(long?))
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(value.Value != null
                            ? value.Value.ToString()
                            : String.Empty);
                }
                else if (value?.Type == typeof(decimal))
                {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(new DecimalValue((decimal)value.Value));
                }
                else if (value?.Type == typeof(decimal?))
                {
                    cell.DataType = CellValues.Number;
                    var nullableDecimalValue = (decimal?)value.Value;
                    cell.CellValue = new CellValue(nullableDecimalValue != null
                        ? new DecimalValue(nullableDecimalValue.Value)
                        : new DecimalValue());
                }
                else if (value?.Type == typeof(DateTime))
                {
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    var dateTimeValue = (DateTime)value.Value;
                    cell.CellValue = new CellValue(dateTimeValue.ToOADate().ToString().Replace(',', '.'));
                }
                else if (value?.Type == typeof(DateTime?))
                {
                    // этот вариант тоже рабочий
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    var dateTimeValue = (DateTime?)value.Value;
                    cell.CellValue = new CellValue(dateTimeValue.HasValue
                        ? dateTimeValue.Value.ToOADate().ToString().Replace(',', '.')
                        : string.Empty);
                }
                else
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(value.Value?.ToString());
                }
            }
        }
    }
}