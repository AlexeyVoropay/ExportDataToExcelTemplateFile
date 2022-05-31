using System;

namespace ExcelExport.Models
{
    public class CellToInsert
    {        
        public int RowSize { get; set; }
        public int ColumnSize { get; set; }
        public string FieldName { get; set; }

        public Type Type { get; set; }        
        public string StyleCellReference { get; set; }

        public CellToInsert()
        { }
    }
}