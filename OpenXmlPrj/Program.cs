using ExcelTemplates;
using ExportDataToExcelTemplate;
using System;

namespace OpenXmlPrj
{
    class Program
    {        
        static void Main(string[] args)
        {
            var data = TestData.GetTestData();
            new Worker().Export(data.GetTables(), data.GetFields(), "template");
            Console.WriteLine("Done. Press any key, for exit!");
            Console.ReadKey();
        }
    }
}
