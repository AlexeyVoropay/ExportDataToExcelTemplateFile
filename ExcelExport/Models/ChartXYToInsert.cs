namespace ExcelExport.Models
{
    using System.Collections.Generic;

    public class ChartXYToInsert
    {
        public int Number { get; set; }
        public string ChartTitle { get; set; }
        public List<ChartSeriesToInsert> SeriesListToInsert { get; set; }
        public List<ChartAxisToInsert> AxisListToInsert { get; set; }
    }
}