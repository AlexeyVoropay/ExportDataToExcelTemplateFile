using System;
using System.Collections.Generic;

namespace ExcelExport.Models
{
    public class RowBlockToInsert
    {
        public int RowId { get; set; }
        //public string ColumnName { get; set; }
        public List<CellToInsert> CellsToInsert { get; set; }              

        public RowBlockToInsert()
        { }        
    }
}