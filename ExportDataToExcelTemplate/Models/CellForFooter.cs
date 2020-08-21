namespace ExportDataToExcelTemplate.Models
{
    using System;
    using DocumentFormat.OpenXml.Spreadsheet;

    public class CellForFooter
    {
        /// <summary>
        /// ячейка
        /// </summary>
        public Cell Cell { get; private set; }
        /// <summary>
        /// значение
        /// </summary>
        public String Value { get; private set; }

        public CellForFooter(Cell cell, String value)
        {
            Cell = cell;
            Value = value;
        }
    }
}
