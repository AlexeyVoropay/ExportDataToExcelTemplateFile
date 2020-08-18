using System.Collections.Generic;
using System.Data;

namespace ExcelTemplates.Models
{
    public class DrillingReport
    {
        public string ReportDate { get; set; }

        public string ReportNumber { get; set; }

        public List<KeyValuePair<string, string>> WellInfo { get; set; }

        public List<KeyValuePair<string, string>> SvInfo { get; set; }

        public List<KeyValuePair<string, string>> Сonstruction { get; internal set; }

        public List<KeyValuePair<string, string>> GetFields()
        {
            return new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("ReportDate", ReportDate),
                    new KeyValuePair<string, string>("ReportNumber", ReportNumber),
                    new KeyValuePair<string, string>("Hse.NumStopCards", Hse.NumStopCards.ToString()),
                };
        }

        public List<DataTable> GetTables()
        {
            var wellInfoTable = new DataTable("WellInfo");
            var wellInfoTableCol1 = new DataColumn { DataType = typeof(string), ColumnName = "WellInfo.Name" };
            var wellInfoTableCol2 = new DataColumn { DataType = typeof(string), ColumnName = "WellInfo.Value" };
            wellInfoTable.Columns.Add(wellInfoTableCol1);
            wellInfoTable.Columns.Add(wellInfoTableCol2);
            foreach (var wellInfoItem in WellInfo)
            {
                var row = wellInfoTable.NewRow();
                row["WellInfo.Name"] = wellInfoItem.Key;
                row["WellInfo.Value"] = wellInfoItem.Value;
                wellInfoTable.Rows.Add(row);
            }

            var svInfoTable = new DataTable("SvInfo");
            var svInfoTableCol1 = new DataColumn { DataType = typeof(string), ColumnName = "SvInfo.Name" };
            var svInfoTableCol2 = new DataColumn { DataType = typeof(string), ColumnName = "SvInfo.Value" };
            svInfoTable.Columns.Add(svInfoTableCol1);
            svInfoTable.Columns.Add(svInfoTableCol2);
            foreach (var svInfoItem in SvInfo)
            {
                var row = svInfoTable.NewRow();
                row["SvInfo.Name"] = svInfoItem.Key;
                row["SvInfo.Value"] = svInfoItem.Value;
                svInfoTable.Rows.Add(row);
            }

            var сonstructionTable = new DataTable("Сonstruction");
            var сonstructionTableCol1 = new DataColumn { DataType = typeof(string), ColumnName = "Сonstruction.Name" };
            var сonstructionTableCol2 = new DataColumn { DataType = typeof(string), ColumnName = "Сonstruction.Value" };
            сonstructionTable.Columns.Add(сonstructionTableCol1);
            сonstructionTable.Columns.Add(сonstructionTableCol2);
            foreach (var сonstructionItem in Сonstruction)
            {
                var row = сonstructionTable.NewRow();
                row["Сonstruction.Name"] = сonstructionItem.Key;
                row["Сonstruction.Value"] = сonstructionItem.Value;
                сonstructionTable.Rows.Add(row);
            }

            return new List<DataTable>
                {
                    wellInfoTable,
                    svInfoTable,
                    сonstructionTable,
                };
        }

        public Hse Hse { get; set; }
    }

    public class Hse
    {
        public long Id { get; set; }
        public long ParentId { get; set; }
        public string ReportDate { get; set; }
        public string LastIncidentInjury { get; set; }
        public short NumStopCards { get; set; }
        public short NumWorkPermits { get; set; }
        public short NumAlarmsDone { get; set; }
        public string LastSafetyMeeting { get; set; }
    }
}