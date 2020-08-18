using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExportDataToExcelTemplate.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDataToExcelTemplate
{
    public static class ColumnHelper
    {
        public static string GetColumnIndex(string cellReferenceValue)
        {
            return new string(cellReferenceValue.ToCharArray().Where(p => !char.IsDigit(p)).ToArray());
        }
    }
}
