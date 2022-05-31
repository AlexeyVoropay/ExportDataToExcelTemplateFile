using System;
using System.Collections.Generic;
using System.Text;
using ExcelExport.Models;

namespace ExcelExport.Interfaces
{
    public interface IExcelItem
    {
        List<ValueToInsert> GetFields(string modelName);
    }
}
