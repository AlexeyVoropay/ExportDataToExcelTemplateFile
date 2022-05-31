namespace ExcelTemplates.TemplatesModels
{
    using System.Collections.Generic;
    using System.Data;

    public class WalletReport
    {
        public string ReportDate { get; set; }

        public string ReportNumber { get; set; }

        public List<WalletReportItem> Trajectory { get; set; }
        

        public List<KeyValuePair<string, string>> GetFields()
        {
            return new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("ReportDate", ReportDate),
                    new KeyValuePair<string, string>("ReportNumber", ReportNumber),
                };
        }

        public List<DataTable> GetTables()
        {
            var trajectoryTable = new DataTable("WalletReportItem");
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Md" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Incl" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Azi" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Tvd" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Closure" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Dls" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "WalletReportItem.Compare" });
            foreach (var trajectoryItem in Trajectory)
            {
                var row = trajectoryTable.NewRow();
                row["WalletReportItem.Md"] = trajectoryItem.Md;
                row["WalletReportItem.Incl"] = trajectoryItem.Incl;
                row["WalletReportItem.Azi"] = trajectoryItem.Azi;
                row["WalletReportItem.Tvd"] = trajectoryItem.Tvd;
                row["WalletReportItem.Closure"] = trajectoryItem.Closure;
                row["WalletReportItem.Dls"] = trajectoryItem.Dls;
                row["WalletReportItem.Compare"] = trajectoryItem.Compare;
                trajectoryTable.Rows.Add(row);
            }
            
            return new List<DataTable>
                {
                    trajectoryTable,
                };
        }
    }
}