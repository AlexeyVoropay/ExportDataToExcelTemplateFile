using DocumentFormat.OpenXml;

namespace ExcelExport.Models
{
    public class ChartAxisToInsert
    {
        public string AxisText { get; set; }
        public double ScalingMinAxisValue { get; set; }        
        public double ScalingMaxAxisValue { get; set; }
        public double? MajorUnitValue { get; set; }
    }
}