using NpoiFluentExcel.ConsoleDemo.Domain;

namespace NpoiFluentExcel.ConsoleDemo.Seed
{
    public static class PortfolioSeed
    {
        public static List<PortfolioRecord> Generate()
        {
            return new List<PortfolioRecord>
            {
                new PortfolioRecord
                {
                    AssetName = "Apple Inc.",
                    IsinCode = "US0378331005",
                    Category = "Equity",
                    Counterparty = "Goldman Sachs",
                    SettlementDate = new DateTime(2024, 01, 10),

                    Quantity = 10,
                    Price = 180.50m,
                    TotalValue = 1805.00m,

                    Commission = 5.20m,
                    ServiceFee = 1.10m,
                    CustodyFee = 0.50m,
                    CashFlow = -1811.80m,

                    AccountType = "Investment",
                    AccountNumber = "ACC-001",
                    OwnerName = "John Doe",
                    IdType = "Passport",
                    IdValue = "X1234567",
                    Address = "Main Street 1, New York",

                    GeneratedOn = DateTime.Today
                },

                new PortfolioRecord
                {
                    AssetName = "Microsoft Corp.",
                    IsinCode = "US5949181045",
                    Category = "Equity",
                    Counterparty = "JP Morgan",
                    SettlementDate = new DateTime(2024, 01, 12),

                    Quantity = 5,
                    Price = 320.75m,
                    TotalValue = 1603.75m,

                    Commission = 4.10m,
                    ServiceFee = 0.90m,
                    CustodyFee = 0.40m,
                    CashFlow = -1609.15m,

                    AccountType = "Investment",
                    AccountNumber = "ACC-001",
                    OwnerName = "John Doe",
                    IdType = "Passport",
                    IdValue = "X1234567",
                    Address = "Main Street 1, New York",

                    GeneratedOn = DateTime.Today
                },

                new PortfolioRecord
                {
                    AssetName = "Tesla Inc.",
                    IsinCode = "US88160R1014",
                    Category = "Equity",
                    Counterparty = "Morgan Stanley",
                    SettlementDate = new DateTime(2024, 01, 15),

                    Quantity = 3,
                    Price = 250.00m,
                    TotalValue = 750.00m,

                    Commission = 3.50m,
                    ServiceFee = 0.80m,
                    CustodyFee = 0.30m,
                    CashFlow = -754.60m,

                    AccountType = "Investment",
                    AccountNumber = "ACC-001",
                    OwnerName = "John Doe",
                    IdType = "Passport",
                    IdValue = "X1234567",
                    Address = "Main Street 1, New York",

                    GeneratedOn = DateTime.Today
                },

                new PortfolioRecord
                {
                    AssetName = "Amazon.com Inc.",
                    IsinCode = "US0231351067",
                    Category = "Equity",
                    Counterparty = "Citibank",
                    SettlementDate = new DateTime(2024, 01, 18),

                    Quantity = 2,
                    Price = 140.30m,
                    TotalValue = 280.60m,

                    Commission = 2.10m,
                    ServiceFee = 0.50m,
                    CustodyFee = 0.20m,
                    CashFlow = -283.40m,

                    AccountType = "Investment",
                    AccountNumber = "ACC-001",
                    OwnerName = "John Doe",
                    IdType = "Passport",
                    IdValue = "X1234567",
                    Address = "Main Street 1, New York",

                    GeneratedOn = DateTime.Today
                }
            };
        }
    }
}
