using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ExportDataToExcelTemplate.Models
{
    public class MergeCellReference
    {
        public StringValue Reference { get; set; } 
        
        public CellReference CellFrom 
        {
            get
            {
                return new CellReference(Reference.Value.Split(':')[0]);
            }
            set
            {
                Reference = $"{value}:{CellTo.Reference}";
            }
        }

        public CellReference CellTo
        {
            get
            {
                return new CellReference(Reference.Value.Split(':')[1]);
            }
            set
            {
                Reference = $"{CellFrom.Reference}:{value}";
            }
        }

        public MergeCellReference(StringValue mergeCellReference)
        {
            Reference = mergeCellReference;
        }
    }
}