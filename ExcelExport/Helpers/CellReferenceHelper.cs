namespace ExcelExport.Helpers
{
    using System;
    using System.Linq;
    using System.Text.RegularExpressions;

    public static class CellReferenceHelper
    {
        public static uint GetRowIndex(string cellReferenceValue)
        {
            return Convert.ToUInt32(Regex.Replace(cellReferenceValue, @"[^\d]+", ""));
        }

        public static string GetColumnIndex(string cellReferenceValue)
        {
            return new string(cellReferenceValue.ToCharArray().Where(p => !char.IsDigit(p)).ToArray());
        }
    }
}
