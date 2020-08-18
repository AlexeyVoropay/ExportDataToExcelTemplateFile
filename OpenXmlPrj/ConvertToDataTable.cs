using OpenXmlPrj.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace OpenXmlPrj
{
    public class ConvertToDataTable
    {
        public DataTable ExcelTableLines(IEnumerable<IDataForTest> lines)
        {
            var dt = CreateTable();
            dt.TableName = "DataField";
            foreach (var line in lines)
            {
                var row = dt.NewRow();
                row["AAA"] = line.A;
                row["BBB"] = line.B;
                row["CCC"] = line.C;
                dt.Rows.Add(row);
            }
            return dt;
        }

        public DataTable ExcelTableLines2(IEnumerable<IDataForTest> lines)
        {
            var dt = CreateTable2();
            dt.TableName = "DataField2";
            foreach (var line in lines)
            {
                var row = dt.NewRow();
                row["AAA"] = line.A;
                row["BBB"] = line.B;
                row["CCC"] = line.C;
                dt.Rows.Add(row);
            }
            return dt;
        }

        //public Hashtable ExcelTableHeader(Int32 count)
        //{
        //    var head = new Dictionary<String, String> { { "Date", DateTime.Today.Date.ToShortDateString() }, { "Count", count.ToString() } };
        //    return new Hashtable(head);
        //}

        public List<KeyValuePair<String, String>> Fields(Int32 count)
        {
            return new List<KeyValuePair<String, String>> {
                new KeyValuePair<String, String>("Label.Date", DateTime.Today.Date.ToShortDateString()),
                new KeyValuePair<String, String>("Label.Count", count.ToString()) 
            };
        }

        private DataTable CreateTable()
        {
            var dt = new DataTable("ExelTable");
            var col = new DataColumn { DataType = typeof(String), ColumnName = "AAA" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "BBB" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "CCC" };
            dt.Columns.Add(col);
            return dt;
        }

        private DataTable CreateTable2()
        {
            var dt = new DataTable("ExelTable");
            var col = new DataColumn { DataType = typeof(String), ColumnName = "AAA" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "BBB" };
            dt.Columns.Add(col);
            col = new DataColumn { DataType = typeof(String), ColumnName = "CCC" };
            dt.Columns.Add(col);
            return dt;
        }
    }
}
