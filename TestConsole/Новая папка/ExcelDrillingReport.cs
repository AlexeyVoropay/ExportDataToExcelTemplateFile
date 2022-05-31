namespace GoExcelReport.Models.ExcelExport
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using global::ExcelExport.Helpers;
    using global::ExcelExport.Interfaces;
    using global::ExcelExport.Models;

    public class ExcelDrillingReport
    {
        public ExcelDrillingReport()
        {
            //Well = new ExcelDrillingReportWell();
            //Personnel = new ExcelDrillingReportPersonnel();
            //Сonstruction = new List<ExcelDrillingReportСonstructionItem>();
            //Hse = new ExcelDrillingReportHse();
            //Knbk = new List<ExcelDrillingReportKnbkItem>();
            //KnbkUom = new ExcelDrillingReportKnbkUoms();
            Trajectory = new List<ExcelDrillingReportTrajectoryItem>();
            TrajectoryUom = new ExcelDrillingReportTrajectoryUoms();
            //Gti = new List<ExcelDrillingReportGtiItem>();
            //NightGti = new List<ExcelDrillingReportGtiItem>();
            //Bit = new List<ExcelDrillingReportBitItem>();
            //MudPlan = new List<ExcelDrillingReportMudPlanItem>();
            //Mud = new List<ExcelDrillingReportMudItem>();
            //MudUom = new ExcelDrillingReportMudUoms();
        }

        /// <summary>
        /// Дата
        /// </summary>
        public DateTime? ReportDate { get; set; }
        /// <summary>
        /// Рапорт №
        /// </summary>
        public int? ReportNumber { get; set; }
        /// <summary>
        /// Забой на 24:00, м
        /// </summary>
        public decimal? DepthAt24 { get; set; }
        /// <summary>
        /// Выполняемые работы на 24:00
        /// </summary>
        public string ActivityAt24 { get; set; }
        /// <summary>
        /// НПВ за сутки, час
        /// </summary>
        public decimal? NptTime { get; set; }
        /// <summary>
        /// Накопительное НПВ, час
        /// </summary>
        public decimal? TotalNptTime { get; set; }
        /// <summary>
        /// Забой на 6:00, м
        /// </summary>
        public decimal? DepthAt6 { get; set; }
        /// <summary>
        /// Выполняемые работы на 06:00 (Работы на 6:00)
        /// </summary>
        public string ActivityAt6 { get; set; }
        /// <summary>
        /// Планируемые работы на сутки
        /// </summary>
        public string PlanActivity { get; set; }
        /// <summary>
        /// Суточные операции, Продолжительность, Итого часов   
        /// </summary>
        public decimal GtiSummaryDuration { get; set; }
        /// <summary>
        /// Операции с 00:00 до 06:00, Продолжительность, Итого часов   
        /// </summary>
        public decimal NightGtiSummaryDuration { get; set; }

        //public ExcelDrillingReportWell Well { get; set; }
        //public ExcelDrillingReportPersonnel Personnel { get; set; }
        //public List<ExcelDrillingReportСonstructionItem> Сonstruction { get; set; }
        //public ExcelDrillingReportHse Hse { get; set; }
        //public List<ExcelDrillingReportKnbkItem> Knbk { get; set; }
        //public ExcelDrillingReportKnbkUoms KnbkUom { get; set; }
        public List<ExcelDrillingReportTrajectoryItem> Trajectory { get; set; }
        public ExcelDrillingReportTrajectoryUoms TrajectoryUom { get; set; }
        //public List<ExcelDrillingReportGtiItem> Gti { get; set; }
        //public List<ExcelDrillingReportGtiItem> NightGti { get; set; }
        //public List<ExcelDrillingReportBitItem> Bit { get; set; }
        //public List<ExcelDrillingReportMudPlanItem> MudPlan { get; set; }
        //public List<ExcelDrillingReportMudItem> Mud { get; set; }
        //public ExcelDrillingReportMudUoms MudUom { get; set; }

        public SheetExportData GetSheetExportData()
        {
            return new SheetExportData
            {
                SheetName = "Суточный отчет",
                FieldsToInserts = GetFieldsToInserts(),
                ArraysToInserts = GetArraysToInserts(),
            };
        }

        private List<ValueToInsert> GetFieldsToInserts()
        {
            var fields = new List<ValueToInsert>
            {
                new ValueToInsert(nameof(ReportDate), typeof(DateTime?), ReportDate),
                new ValueToInsert(nameof(ReportNumber), typeof(int?), ReportNumber),

                new ValueToInsert(nameof(DepthAt24), typeof(decimal?), DepthAt24),
                new ValueToInsert(nameof(ActivityAt24), typeof(string), ActivityAt24),
                new ValueToInsert(nameof(NptTime), typeof(decimal?), NptTime),
                new ValueToInsert(nameof(TotalNptTime), typeof(decimal?), TotalNptTime),
                new ValueToInsert(nameof(DepthAt6), typeof(decimal?), DepthAt6),
                new ValueToInsert(nameof(ActivityAt6), typeof(string), ActivityAt6),
                new ValueToInsert(nameof(PlanActivity), typeof(string), PlanActivity),

                new ValueToInsert(nameof(GtiSummaryDuration), typeof(decimal), GtiSummaryDuration),
                new ValueToInsert(nameof(NightGtiSummaryDuration), typeof(decimal), NightGtiSummaryDuration),       
            };
            //fields.AddRange(Well.GetFields(nameof(Well)));
            //fields.AddRange(Personnel.GetFields(nameof(Personnel)));
            //fields.AddRange(Hse.GetFields(nameof(Hse)));
            //fields.AddRange(KnbkUom.GetFields(nameof(KnbkUom)));
            fields.AddRange(TrajectoryUom.GetFields(nameof(TrajectoryUom)));
            //fields.AddRange(MudUom.GetFields(nameof(MudUom)));
            return fields;
        }

        private List<TableToInsert> GetArraysToInserts()
        {
            //var mudList = MudPlan.Select(x => (IExcelItem) x).ToList();
            //mudList.AddRange(Mud.Select(x => (IExcelItem)x).ToList());
            //if (!mudList.Any())
            //    mudList.Add(new ExcelDrillingReportMudPlanItem());
            return new List<TableToInsert>
            {
                //ExcelHelper.GetTable("Сonstruction", Сonstruction.Select(x => (IExcelItem)x).ToList()),
                //ExcelHelper.GetTable("Knbk", Knbk.Select(x => (IExcelItem)x).ToList()),
                ExcelHelper.GetTable("Trajectory", Trajectory.Select(x => (IExcelItem)x).ToList()),
                //ExcelHelper.GetTable("Bit", Bit.Select(x => (IExcelItem)x).ToList()),
                //ExcelHelper.GetTable("Gti", Gti.Select(x => (IExcelItem)x).ToList()),
                //ExcelHelper.GetTable("NightGti", NightGti.Select(x => (IExcelItem)x).ToList()),
                
                //ExcelHelper.GetTable("Mud", mudList),
            };
        }
    }
}