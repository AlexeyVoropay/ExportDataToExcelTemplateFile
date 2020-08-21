namespace ExportDataToExcelTemplate.Models
{
    using DocumentFormat.OpenXml;
    using System.Linq;

    public class CellReference
    {
        public StringValue Reference { get; set; } 
        
        public int RowIndex
        {
            get
            {
                return int.Parse(new string(Reference.Value.ToCharArray().Where(p => char.IsDigit(p)).ToArray()));
            }
            set
            {
                Reference = $"{ColumnIndex}{value}";
            }
        }

        public string ColumnIndex
        {
            get
            {
                return new string(Reference.Value.ToCharArray().Where(p => !char.IsDigit(p)).ToArray());
            }
            set
            {
                Reference =  $"{value}{RowIndex}";
            }
        }

        public CellReference(StringValue cellReference)
        {
            Reference = cellReference;
        }
    }
}