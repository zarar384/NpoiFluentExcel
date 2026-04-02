using NPOI.XSSF.UserModel;
using NpoiFluentExcel.ConsoleDemo.Infrastructure;
using NpoiFluentExcel.ConsoleDemo.Seed;
using NpoiFluentExcel.Extensions;

namespace NpoiFluentExcel.ConsoleDemo.Scenarios
{
    public static class MappingDemo
    {
        public static void Run()
        {
            // Data
            var items = PortfolioSeed.Generate();

            // Create workbook and write data
            var workbook = new XSSFWorkbook();

            workbook.WriteSheet(items);
            //  .AddTableStyle()
            //  .AutoSizeColumns();


            string filePath = @"C:\temp\mapping.xlsx";

            ExcelSaver.Save(workbook, filePath);

            Console.WriteLine($"File saved to {filePath}");
            Console.WriteLine("Mapping demo done");
        }
    }
}
