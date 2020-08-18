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
    public static class RowHelper
    {
        public static uint GetRowIndex(string cellReferenceValue)
        {
            return Convert.ToUInt32(Regex.Replace(cellReferenceValue, @"[^\d]+", ""));
        }
    }
}
