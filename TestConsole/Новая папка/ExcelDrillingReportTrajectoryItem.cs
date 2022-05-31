namespace GoExcelReport.Models.ExcelExport
{
    using System.Collections.Generic;
    using GoExcelExport.Interfaces;
    using GoExcelExport.Models;

    public class ExcelDrillingReportTrajectoryItem : IExcelItem
    {
        public int? Index { get; set; }
        public decimal? Md { get; set; }
        public decimal? Incl { get; set; }
        public decimal? Azi { get; set; }
        public decimal? Tvd { get; set; }
        public decimal? Closure { get; set; }
        public decimal? Dls { get; set; }
        public string Compare { get; set; }

        public List<ValueToInsert> GetFields(string modelName)
        {
            modelName = modelName ?? string.Empty;
            var separator = !string.IsNullOrEmpty(modelName) ? "." : null;
            return new List<ValueToInsert>
            {
                new ValueToInsert($"{modelName}{separator}{nameof(Index)}", typeof(int?), Index),
                new ValueToInsert($"{modelName}{separator}{nameof(Md)}", typeof(decimal?), Md),
                new ValueToInsert($"{modelName}{separator}{nameof(Incl)}", typeof(decimal?), Incl),
                new ValueToInsert($"{modelName}{separator}{nameof(Azi)}", typeof(decimal?), Azi),
                new ValueToInsert($"{modelName}{separator}{nameof(Tvd)}", typeof(decimal?), Tvd),
                new ValueToInsert($"{modelName}{separator}{nameof(Closure)}", typeof(decimal?), Closure),
                new ValueToInsert($"{modelName}{separator}{nameof(Dls)}", typeof(decimal?), Dls),
                new ValueToInsert($"{modelName}{separator}{nameof(Compare)}", typeof(string), Compare),
            };
        }
    }
}