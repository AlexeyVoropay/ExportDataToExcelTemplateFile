using System;
using System.Collections.Generic;

namespace ExcelExport.Models
{
    public class ColumnBlockToInsert
    {
        public string FirstColumnName { get; set; }
        public int[] ColumnsWidths { get; set; }
        public List<RowBlockToInsert> RowBlocksToInsert { get; set; }

        public ColumnBlockToInsert()
        { }        
    }
}