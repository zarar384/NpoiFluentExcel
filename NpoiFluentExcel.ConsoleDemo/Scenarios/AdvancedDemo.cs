using NPOI.XSSF.UserModel;
using NpoiFluentExcel.ConsoleDemo.Domain;
using NpoiFluentExcel.ConsoleDemo.Infrastructure;
using NpoiFluentExcel.ConsoleDemo.Seed;
using NpoiFluentExcel.Extensions;

namespace NpoiFluentExcel.ConsoleDemo.Scenarios
{
    public static class AdvancedDemo
    {
        public static void Run()
        {
            // Data
            var items = PortfolioSeed.Generate();
            var first = items.First();

            // Prepare header info and logo
            var logoFile = AssetLocator.ResolveLogo();
            var owner = new AccountOwner
            {
                CompanyName = first.OwnerName,
                AddressLine = first.Address,
                Identifier = first.IdValue
            };

            // Create workbook and write data
            var workbook = new XSSFWorkbook();

            workbook.AddSheet<PortfolioRecord>("Report")
                .StartingCell(1)
                .AddCustomBlock(block =>
                {
                    var from = items.Min(x => x.SettlementDate);
                    var to = items.Max(x => x.SettlementDate);

                    block.WriteCell(0, "Date range:", block.Style.Bold);
                    block.WriteCell(1, $"{from:dd.MM.yyyy} - {to:dd.MM.yyyy}", block.Style.Bold);
                    block.NextRow();

                    block.WriteCell(0, "Generated:", block.Style.Bold);
                    block.WriteCell(1, first.GeneratedOn.ToString("dd.MM.yyyy"), block.Style.Bold);
                    block.NextRow();
                    block.NextRow();

                    block.WriteCell(0, $"Portfolio ({first.AccountType})", block.Style.Header);
                    block.NextRow();
                    block.NextRow();

                    block.WriteCell(0, "Account ID:", block.Style.Normal);
                    block.WriteCell(1, first.AccountNumber, block.Style.Bold);
                    block.NextRow();

                    block.WriteCell(0, "Owner:", block.Style.Normal);
                    block.WriteCell(1, first.OwnerName, block.Style.Bold);
                    block.NextRow();

                    block.WriteCell(0, "ID Type:", block.Style.Normal);
                    block.WriteCell(1, first.IdType, block.Style.Bold);
                    block.NextRow();

                    block.WriteCell(0, "ID Value:", block.Style.Normal);
                    block.WriteCell(1, first.IdValue, block.Style.Bold);
                    block.NextRow();

                    block.WriteCell(0, "Address:", block.Style.Normal);
                    block.WriteCell(1, first.Address, block.Style.Bold);
                    block.NextRow();
                })
                .AddClientInfo(owner.FormatHeader(), logoFile)
                .AddColumn("Asset", x => x.AssetName)
                .AddColumn("ISIN Code", x => x.IsinCode)
                .AddColumn("Category", x => x.Category)
                .AddColumn("Counterparty", x => x.Counterparty)
                .AddColumn("Settlement Date", x => x.SettlementDate)
                .AddColumn("Quantity", x => x.Quantity)
                .AddColumn("Price", x => x.Price)
                .AddColumn("Total Value", x => x.TotalValue)
                .AddColumn("Commission", x => x.Commission).Sum()
                .AddColumn("Service Fee", x => x.ServiceFee).Sum()
                .AddColumn("Custody Fee", x => x.CustodyFee).Sum()
                .AddColumn("Cash Flow", x => x.CashFlow).Sum()
                .WriteRows(items)
                .AddTableStyle()
                .AddCustomBlock(x =>
                {
                    x.NextRow();
                    x.WriteCell(0, "Signature:", x.Style.Normal);
                    x.WriteCell(1, "________________________", x.Style.Normal);
                })
                .AutoSizeColumns();


            string filePath = @"C:\temp\advanced.xlsx";

            ExcelSaver.Save(workbook, filePath);

            Console.WriteLine($"File saved to {filePath}");
            Console.WriteLine("Advanced demo done");
        }
    }
}
