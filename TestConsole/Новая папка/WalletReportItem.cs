using System;
using System.Collections.Generic;
using ExcelExport.Interfaces;
using ExcelExport.Models;

namespace ExcelTemplates.TemplatesModels
{
    public class WalletReportItem : IExcelItem
    {
        public DateTime? Date { get; set; }
        public int? ClientId { get; set; }
        public int? OutTransactions { get; set; }
        public int? InTransactions { get; set; }

        public List<ValueToInsert> GetFields(string modelName)
        {
            modelName ??= string.Empty;
            var separator = !string.IsNullOrEmpty(modelName) ? "." : null;
            return new List<ValueToInsert>
            {
                new ValueToInsert($"{modelName}{separator}{nameof(Date)}", typeof(DateTime?), Date),
                new ValueToInsert($"{modelName}{separator}{nameof(ClientId)}", typeof(int?), ClientId),
                new ValueToInsert($"{modelName}{separator}{nameof(OutTransactions)}", typeof(int?), OutTransactions),
                new ValueToInsert($"{modelName}{separator}{nameof(InTransactions)}", typeof(int?), InTransactions),
            };
        }
    }
}