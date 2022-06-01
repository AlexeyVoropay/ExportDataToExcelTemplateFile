namespace ExcelTemplates.TemplatesModels
{
    using System.Collections.Generic;
    using System.Data;
    using System.Linq;
    using ExcelExport.Helpers;
    using ExcelExport.Interfaces;
    using ExcelExport.Models;

    public class ExcelWalletReport
    {
        //public string ReportDate { get; set; }

        //public string ReportNumber { get; set; }

        public List<WalletReportItem> Items { get; set; }

        public ExcelWalletReport()
        {
            Items = new List<WalletReportItem>();
        }

        public SheetExportData GetSheetExportData()
        {
            return new SheetExportData
            {
                SheetName = "WalletReport",
                FieldsToInserts = GetFieldsToInserts(),
                ArraysToInserts = GetArraysToInserts(),
            };
        }

        private List<ValueToInsert> GetFieldsToInserts()
        {
            var fields = new List<ValueToInsert>
            {
                //new ValueToInsert(nameof(ReportDate), typeof(DateTime?), ReportDate),
                //new ValueToInsert(nameof(ReportNumber), typeof(int?), ReportNumber),
            };
            //fields.AddRange(Personnel.GetFields(nameof(Personnel)));
            return fields;
        }

        private List<TableToInsert> GetArraysToInserts()
        {
            return new List<TableToInsert>
            {
                ExcelHelper.GetTable("Item", Items.Select(x => (IExcelItem)x).ToList()),
            };
        }
    }
}