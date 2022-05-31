
namespace ExcelExport.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Spreadsheet;
    using global::ExcelExport.Interfaces;
    using global::ExcelExport.Models;

    public static class ExcelHelper
    {
        public static void AddEmptyRows(SheetData sheetData, uint? maxRowIndex = null)
        {
            var rows = sheetData.Elements<Row>().ToList();
            if (!maxRowIndex.HasValue)
            {
                maxRowIndex = rows.Max(x => x.RowIndex);
            }
            for (int i = 0; i < maxRowIndex + 1; i++)
            {
                if (rows.FirstOrDefault(x => x.RowIndex == i) == null)
                {
                    var prevRow = sheetData.Elements<Row>().FirstOrDefault(x => x.RowIndex == i - 1);
                    if (prevRow != null)
                    {
                        prevRow.InsertAfterSelf(new Row { RowIndex = (uint)i });
                    }
                }
            }
        }
                
        public static TableToInsert GetTable(string tableName, List<IExcelItem> items)
        {
            var arrays = new List<ArrayToInsert>();
            for (int i = 0; i < items.Count(); i++)
            {                
                arrays.Add(new ArrayToInsert(i, items[i].GetFields(null)));
            }
            return new TableToInsert
            {
                TableName = tableName,
                Rows = arrays,
            };
        }

        public static List<ValueToInsert> GetValuesToInsert(string modelName, List<string> values, string itemName, string countPrefix)
        {
            modelName = modelName ?? string.Empty;
            var separator = !string.IsNullOrEmpty(modelName) ? "." : null;
            var result = new List<ValueToInsert>();
            if (values == null || !values.Any())
                values = new List<string> { string.Empty };            
            for (int i = 0; i < values.Count; i++)
            {
                result.Add(new ValueToInsert($"{modelName}{separator}{itemName}{separator}{countPrefix}{i + 1}",
                    typeof(string), values[i]));
            }
            return result;
        }
    }
}