using NPOI.XSSF.UserModel;
using NpoiFluentExcel.ConsoleDemo.Domain;
using NpoiFluentExcel.ConsoleDemo.Infrastructure;
using NpoiFluentExcel.ConsoleDemo.Seed;
using NpoiFluentExcel.Extensions;

namespace NpoiFluentExcel.ConsoleDemo.Scenarios
{
    public static class BasicDemo
    {
        public static void Run()
        {
            // Data
            var items = PortfolioSeed.Generate();

            // Create workbook and write data
            var workbook = new XSSFWorkbook();

            workbook.AddSheet<PortfolioRecord>("Report")
                .AddColumn("Asset", x => x.AssetName)
                .AddColumn("ISIN Code", x => x.IsinCode)
                .AddColumn("Category", x => x.Category)
                .AddColumn("Counterparty", x => x.Counterparty)
                .AddColumn("Settlement Date", x => x.SettlementDate)
                .AddColumn("Quantity", x => x.Quantity)
                .AddColumn("Price", x => x.Price)
                .AddColumn("Total Value", x => x.TotalValue)
                .AddColumn("Commission", x => x.Commission)
                .AddColumn("Service Fee", x => x.ServiceFee)
                .AddColumn("Custody Fee", x => x.CustodyFee)
                .AddColumn("Cash Flow", x => x.CashFlow)
                .WriteRows(items);


            string filePath = @"C:\temp\basic.xlsx";

            ExcelSaver.Save(workbook, filePath);

            Console.WriteLine($"File saved to {filePath}");
            Console.WriteLine("Basic demo done");
        }
    }
}
