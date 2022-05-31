namespace GoExcelReport.Models.ExcelExport
{
    using System.Collections.Generic;
    using global::ExcelExport.Interfaces;
    using global::ExcelExport.Models;

    public class ExcelDrillingReportTrajectoryUoms : IExcelItem
    {        
        public string Md { get; set; }
        public string Incl { get; set; }
        public string Azi { get; set; }
        public string Tvd { get; set; }
        public string Closure { get; set; }
        public string Dls { get; set; }

        public List<ValueToInsert> GetFields(string modelName)
        {
            modelName = modelName ?? string.Empty;
            var separator = !string.IsNullOrEmpty(modelName) ? "." : null;
            return new List<ValueToInsert>
            {
                new ValueToInsert($"{modelName}{separator}{nameof(Md)}", typeof(string), Md),
                new ValueToInsert($"{modelName}{separator}{nameof(Incl)}", typeof(string), Incl),
                new ValueToInsert($"{modelName}{separator}{nameof(Azi)}", typeof(string), Azi),
                new ValueToInsert($"{modelName}{separator}{nameof(Tvd)}", typeof(string), Tvd),
                new ValueToInsert($"{modelName}{separator}{nameof(Closure)}", typeof(string), Closure),
                new ValueToInsert($"{modelName}{separator}{nameof(Dls)}", typeof(string), Dls),
            };
        }
    }
}