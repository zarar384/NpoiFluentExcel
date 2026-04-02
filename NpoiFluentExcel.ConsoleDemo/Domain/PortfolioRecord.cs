using NpoiFluentExcel.Mapping.Attributes;

namespace NpoiFluentExcel.ConsoleDemo.Domain
{

    [Sheet("Report Mapping")]
    public class PortfolioRecord
    {
        [Column("Asset")]
        public string AssetName { get; set; }

        [Column("ISIN Code")]
        public string IsinCode { get; set; }

        [Column("Category")]
        public string Category { get; set; }

        [Column("Counterparty")]
        public string Counterparty { get; set; }

        [Column("Settlement Date")]
        public DateTime SettlementDate { get; set; }

        [Column("Quantity")]
        public decimal Quantity { get; set; }

        [Column("Price")]
        public decimal Price { get; set; }

        [Column("Total Value")]
        public decimal TotalValue { get; set; }

        [Column("Commission")]
        public decimal Commission { get; set; }

        [Column("Service Fee")]
        public decimal ServiceFee { get; set; }

        [Column("Custody Fee")]
        public decimal CustodyFee { get; set; }

        [Column("Cash Flow")]
        public decimal CashFlow { get; set; }

        // details for custom block (not exported as columns)
        public string AccountType { get; set; }
        public string AccountNumber { get; set; }
        public string OwnerName { get; set; }
        public string IdType { get; set; }
        public string IdValue { get; set; }
        public string Address { get; set; }
        public DateTime GeneratedOn { get; set; }
    }
}
