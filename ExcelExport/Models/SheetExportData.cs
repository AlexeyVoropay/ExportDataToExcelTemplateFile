namespace ExcelExport.Models
{
    using System.Collections.Generic;

    public class SheetExportData
    {
        /// <summary>
        /// Наименование листа
        /// </summary>
        public string SheetName { get; set; }
        /// <summary>
        /// Количество блоков, которые будут копироваться
        /// </summary>
        public int RowBlocksForCopyAndInsert { get; set; }
        /// <summary>
        /// Индекс первой строки, шаблоного блока
        /// </summary>
        public int CopyRowIndexFrom { get; set; }
        /// <summary>
        /// Индекс последней строки, шаблоного блока 
        /// </summary>
        public int CopyRowIndexTo { get; set; }

        public List<ColumnBlockToInsert> ColumnsBlockToInsert { get; set; }
        public List<ValueToInsert> FieldsToInserts { get; set; }
        public List<TableToInsert> ArraysToInserts { get; set; }
        public List<ChartToInsert> ChartsToInserts { get; set; }
        public List<ChartXYToInsert> ChartsXYToInserts { get; set; }
    }
}