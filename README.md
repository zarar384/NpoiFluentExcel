# NpoiFluentExcel

A fluent Excel generation library built on top of NPOI with strong typing, attribute-based mapping, and a clean builder API.

---

## Description

NpoiFluentExcel is a .NET library that simplifies Excel file generation by providing a structured abstraction over NPOI.

It replaces low-level Excel manipulation with a fluent and strongly typed API, improving readability, maintainability, and developer productivity.

The library supports both manual column definition and automatic mapping using attributes.

---

## Features

- Fluent builder API
- Strongly typed column definitions
- Attribute-based mapping (DTO to Excel)
- Built-in styling support
- Excel table creation with autofilter
- Custom header and footer blocks
- Logo insertion support
- Separation of concerns (generation vs IO)

---

## Usage

### Builder API

```csharp
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

ExcelSaver.Save(workbook, "report.xlsx");
```

![test](https://i.imgur.com/ngIIjEz.png)

---

### Attribute-Based Mapping

```csharp
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

    ....
}
```

```csharp
var workbook = new XSSFWorkbook();

workbook.WriteSheet(items);
```

![test](https://i.imgur.com/D2tASy5.png)

---

## Advanced Usage

```csharp
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
```

![test](https://i.imgur.com/NZ4NRxB.png)
