namespace ExcelExport.Models
{
    using System.Collections.Generic;

    public class ArrayToInsert
    {
        public int RowId { get; set; }
        public List<ValueToInsert> Values { get; set; }

        public ArrayToInsert()
        { }

        public ArrayToInsert(int rowId, List<ValueToInsert> values)
        {
            RowId = rowId;
            Values = values;
        }
    }
}