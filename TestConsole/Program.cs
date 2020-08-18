using ExcelTemplates;
using ExportDataToExcelTemplate;
using System;

namespace TestConsole
{
    class Program
    {        
        static void Main(string[] args)
        {
            var data = TestData.GetTestData();
            new Worker().Export(data.GetTables(), data.GetFields(), "template");

            //Test01.RunTest();
            Console.WriteLine("Done. Press any key, for exit!");
            Console.ReadKey();
        }
    }
}
