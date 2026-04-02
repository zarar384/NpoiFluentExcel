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
    .WriteRows(items)
    .AddTableStyle()
    .AutoSizeColumns();

ExcelSaver.Save(workbook, "report.xlsx");
```

---

### Attribute-Based Mapping

```csharp
[Sheet("Report")]
public class PortfolioRecord
{
    [Column("Asset")]
    public string AssetName { get; set; }

    [Column("ISIN Code")]
    public string IsinCode { get; set; }

    [Column("Quantity")]
    public decimal Quantity { get; set; }

    [Column("Price")]
    public decimal Price { get; set; }
}
```

```csharp
var workbook = new XSSFWorkbook();

workbook.WriteSheet(items)
        .AddTableStyle()
        .AutoSizeColumns();
```

---

## Advanced Usage

```csharp
var workbook = new XSSFWorkbook();

workbook.AddSheet<PortfolioRecord>("Report")
    .StartingCell(1)
    .AddCustomBlock(x =>
    {
        x.WriteCell(0, "Report Date:", x.Helper.Bold);
        x.WriteCell(1, DateTime.Today.ToString("dd.MM.yyyy"), x.Helper.Bold);
        x.NextRow();
    })
    .AddClientInfo(
        "Big Patient Gorilla a.s.\nAddress line\nID: 12345678",
        "resources/company-logo.png")
    .AddColumn("Asset", x => x.AssetName)
    .AddColumn("Total Value", x => x.TotalValue)
    .WriteRows(items)
    .AddTableStyle()
    .AutoSizeColumns();
```

---

## Styling

Available styles:

- Normal
- Bold
- Header
- Date
- Number
- Integer
- BoldNumber
- BoldInteger

---

## Custom Blocks

```csharp
.AddCustomBlock(x =>
{
    x.WriteCell(0, "Title", x.Helper.Bold);
    x.NextRow();
})
```

---

## Excel Tables

```csharp
.AddTableStyle()
```

---

## Demo

samples/NpoiFluentExcel.Demo/

---

## Project Structure

```
src/
 ├── Builder/
 ├── Core/
 ├── Extensions/
 ├── Interfaces/
 ├── Mapping/
 └── Styling/

samples/
 └── NpoiFluentExcel.Demo/

tests/
 └── NpoiFluentExcel.Tests/
```