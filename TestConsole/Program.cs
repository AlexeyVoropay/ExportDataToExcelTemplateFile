using ExcelTemplates;
using ExportDataToExcelTemplate;
using System;

namespace TestConsole
{
    class Program
    {        
        static void Main(string[] args)
        {
            //var data = TestData.GetTestData();
            //new Worker().Export(data.GetTables(), data.GetFields(), "template");

            var data2 = TestData.GetTestData2();
            new Worker().Export(data2.GetTables(), data2.GetFields(), "template2");

            var data = TestData.GetTestData();
            var path = @"C:\Users\Zver\Desktop\_Projects\ExportDataToExcelTemplateFile\TestConsole\Новая папка\DrillingReport111.xlsx";
            ExcelExport.ExcelExport.CreateFilledFile(path, )

            Console.WriteLine("Done. Press any key, for exit!");
            Console.ReadKey();
        }
    }
}
