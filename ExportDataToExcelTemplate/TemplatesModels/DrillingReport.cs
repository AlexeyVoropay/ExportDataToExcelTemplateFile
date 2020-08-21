namespace ExcelTemplates.TemplatesModels
{
    using System.Collections.Generic;
    using System.Data;

    public class DrillingReport
    {
        public string ReportDate { get; set; }

        public string ReportNumber { get; set; }

        public List<KeyValuePair<string, string>> WellInfo { get; set; }

        public List<KeyValuePair<string, string>> SvInfo { get; set; }

        public List<KeyValuePair<string, string>> Сonstruction { get; set; }

        public List<KnbkItem> Knbk { get; set; }

        public List<TrajectoryItem> Trajectory { get; set; }

        public string GtiSummaryDuration { get; set; }

        public List<GtiItem> Gti { get; set; }

        public Hse Hse { get; set; }
        

        public List<KeyValuePair<string, string>> GetFields()
        {
            return new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("ReportDate", ReportDate),
                    new KeyValuePair<string, string>("ReportNumber", ReportNumber),
                    new KeyValuePair<string, string>("Hse.NumStopCards", Hse.NumStopCards.ToString()),
                    new KeyValuePair<string, string>("Hse.NumAlarmsDone", Hse.NumAlarmsDone.ToString()),
                    new KeyValuePair<string, string>("Hse.LastSafetyMeeting", Hse.LastSafetyMeeting),
                    new KeyValuePair<string, string>("Hse.Incident", Hse.Incident.ToString()),
                    new KeyValuePair<string, string>("GtiSummaryDuration", GtiSummaryDuration),
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

            var knbkTable = new DataTable("Knbk");
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.Name" });
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.In" });
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.Od" });
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.Connection" });
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.Len" });
            knbkTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Knbk.TotalLen" });
            foreach (var knbkItem in Knbk)
            {
                var row = knbkTable.NewRow();
                row["Knbk.Name"] = knbkItem.Name;
                row["Knbk.In"] = knbkItem.In;
                row["Knbk.Od"] = knbkItem.Od;
                row["Knbk.Connection"] = knbkItem.Connection;
                row["Knbk.Len"] = knbkItem.Len;
                row["Knbk.TotalLen"] = knbkItem.TotalLen;
                knbkTable.Rows.Add(row);
            }

            var trajectoryTable = new DataTable("Trajectory");
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Md" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Incl" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Azi" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Tvd" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Closure" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Dls" });
            trajectoryTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Trajectory.Compare" });
            foreach (var trajectoryItem in Trajectory)
            {
                var row = trajectoryTable.NewRow();
                row["Trajectory.Md"] = trajectoryItem.Md;
                row["Trajectory.Incl"] = trajectoryItem.Incl;
                row["Trajectory.Azi"] = trajectoryItem.Azi;
                row["Trajectory.Tvd"] = trajectoryItem.Tvd;
                row["Trajectory.Closure"] = trajectoryItem.Closure;
                row["Trajectory.Dls"] = trajectoryItem.Dls;
                row["Trajectory.Compare"] = trajectoryItem.Compare;
                trajectoryTable.Rows.Add(row);
            }
            
            var gtiTable = new DataTable("Gti");
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.StartTime" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.EndTime" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.Duration" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.Duration2" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.StartDepth" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.EndDepth" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.Operation" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.NptCategory" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.NptDuration" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.NptResponsible" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.Modes" });
            gtiTable.Columns.Add(new DataColumn { DataType = typeof(string), ColumnName = "Gti.Comment" });
            foreach (var gtiItem in Gti)
            {
                var row = gtiTable.NewRow();
                row["Gti.StartTime"] = gtiItem.StartTime;
                row["Gti.EndTime"] = gtiItem.EndTime;
                row["Gti.Duration"] = gtiItem.Duration;
                row["Gti.Duration2"] = gtiItem.Duration2;
                row["Gti.StartDepth"] = gtiItem.StartDepth;
                row["Gti.EndDepth"] = gtiItem.EndDepth;
                row["Gti.Operation"] = gtiItem.Operation;
                row["Gti.NptCategory"] = gtiItem.NptCategory;
                row["Gti.NptDuration"] = gtiItem.NptDuration;
                row["Gti.NptResponsible"] = gtiItem.NptResponsible;
                row["Gti.Modes"] = gtiItem.Modes;
                row["Gti.Comment"] = gtiItem.Comment;
                gtiTable.Rows.Add(row);
            }

            return new List<DataTable>
                {
                    wellInfoTable,
                    svInfoTable,
                    сonstructionTable,
                    knbkTable,
                    trajectoryTable,
                    gtiTable,
                };
        }
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
        public string Incident { get; internal set; }
    }
}