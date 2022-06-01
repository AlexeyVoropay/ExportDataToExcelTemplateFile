using System.Collections.Generic;
using ExcelTemplates.TemplatesModels;
using GoExcelReport.Models.ExcelExport;

namespace ExcelTemplates
{
    public static class NewTestData
    {
        public static GoExcelReport.Models.ExcelExport.ExcelDrillingReport GetTestData()
        {
            return new GoExcelReport.Models.ExcelExport.ExcelDrillingReport
            {
                ReportDate = System.DateTime.Today,
                ReportNumber = 87,
                
                
                Trajectory = new List<ExcelDrillingReportTrajectoryItem>
                {
                    new ExcelDrillingReportTrajectoryItem{ Md = 2500, Incl = 76.5m, Azi = 213.5m,Tvd =2000,Closure =500, Dls = 1.5m, Compare = "0,5м выше / 0,5м правее"},
                },
                GtiSummaryDuration = 24,                
            };
        }

        public static ExcelWalletReport GetWalletTestData()
        {
            return new ExcelWalletReport
            {
                Items = new List<WalletReportItem> 
                { 
                    new WalletReportItem
                    {
                        Date = System.DateTime.Today,
                        ClientId = 1,
                        InTransactions = 1,
                        OutTransactions = 1,
                    },
                    new WalletReportItem
                    {
                        Date = System.DateTime.Today.AddDays(2),
                        ClientId = 1,
                        InTransactions = 2,
                        OutTransactions = 3,
                    },
                    new WalletReportItem
                    {
                        Date = System.DateTime.Today.AddDays(3),
                        ClientId = 2,
                        InTransactions = 6,
                        OutTransactions = 7,
                    },
                }
            };
        }
    }
}