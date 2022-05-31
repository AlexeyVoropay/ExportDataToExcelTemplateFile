namespace ExcelExport.Models
{
    using System.Collections.Generic;

    public class TableToInsert
    {
        public string TableName { get; set; }
        public List<ArrayToInsert> Rows { get; set; }
    }
}