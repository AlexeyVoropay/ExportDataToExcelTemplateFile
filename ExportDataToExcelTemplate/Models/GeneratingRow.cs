using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcelTemplate.Models;

namespace ExportDataToExcelTemplate.Models
{
    public class GeneratingRow
    {
        /// <summary>
        /// строка
        /// </summary>
        public Row Row { get; private set; }
        /// <summary>
        /// ячейки данной строки
        /// </summary>
        public List<CellForFooter> Cells { get; private set; }

        public GeneratingRow(Row row, Cell cell, String cellValue)
        {
            Row = (Row)row.Clone();
            Row.RowIndex = row.RowIndex;
            var cellClone = (Cell)cell.Clone();
            //cellClone.CellReference = cell.CellReference;
            Cells = new List<CellForFooter> { new CellForFooter((Cell)cell.Clone(), cellValue) };
        }

        public void AddMoreCell(Cell cell, String cellValue)
        {
            var _Cell = (Cell)cell.Clone();
            _Cell.CellReference = cell.CellReference;
            Cells.Add(new CellForFooter(_Cell, cellValue));
        }
    }
}
