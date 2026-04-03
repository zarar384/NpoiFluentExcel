using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NpoiFluentExcel.Builders;
using NpoiFluentExcel.Interfaces;

namespace NpoiFluentExcel.Tests
{
    public class ExcelTableBuilderTests
    {
        private readonly XSSFWorkbook _workbook;
        private readonly ISheet _sheet;

        public ExcelTableBuilderTests()
        {
            // in-memory workbook with one sheet for testing
            _workbook = new XSSFWorkbook();
            _sheet = _workbook.CreateSheet("Test");
        }

        private ExcelTableBuilder<TestDto> CreateBuilder()
            => new ExcelTableBuilder<TestDto>(_sheet, _workbook);

        private class TestDto
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public double Value { get; set; }
        }

        // BASIC FLOW

        [Fact]
        public void Should_WriteSingleRow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            builder
                .AddColumn("Id", x => x.Id)
                .AddColumn("Name", x => x.Name)
                .WriteRow(new TestDto { Id = 1, Name = "A" });

            var row = _sheet.GetRow(1);

            Assert.Equal(1, row.GetCell(0).NumericCellValue);
            Assert.Equal("A", row.GetCell(1).StringCellValue);
        }

        [Fact]
        public void Should_WriteMultipleRows()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            builder
                .AddColumn("Id", x => x.Id)
                .WriteRows(new[]
                {
                    new TestDto { Id = 1 },
                    new TestDto { Id = 2 }
                });

            Assert.Equal(1, _sheet.GetRow(1).GetCell(0).NumericCellValue);
            Assert.Equal(2, _sheet.GetRow(2).GetCell(0).NumericCellValue);
        }

        // STAGES AND FLUENT FLOW

        [Fact]
        public void Should_FollowValidFluentFlow()
        {
            IExcelPrepareStage<TestDto> prepare = CreateBuilder();

            var rowMode = prepare
                .StartingCell(1)
                .AddClientInfo("Client", null);

            var writeStage = rowMode
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 10 });

            writeStage.AutoSizeColumns()
                      .AddTableStyle();
        }

        // SUM

        [Fact]
        public void Sum_ShouldCreateFormulaRow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            builder
                .AddColumn("Value", x => x.Value)
                .Sum()
                .WriteRows(new[]
                {
                    new TestDto { Value = 10 },
                    new TestDto { Value = 20 }
                });

            var totalRow = _sheet.GetRow(3);
            var cell = totalRow.GetCell(0);

            Assert.Equal(CellType.Formula, cell.CellType);
            Assert.Contains("SUBTOTAL", cell.CellFormula);
        }

        [Fact]
        public void Sum_WithoutColumn_ShouldThrow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            Assert.Throws<InvalidOperationException>(() => builder.Sum());
        }

        // STARTING CELL

        [Fact]
        public void StartingCell_ShouldShiftColumns()
        {
            IExcelPrepareStage<TestDto> prepare = CreateBuilder();

            var rowMode = prepare
                .StartingCell(2)
                .AddClientInfo("Client", null);

            rowMode
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 5 });

            var row = _sheet.GetRow(3);

            Assert.Equal(5, row.GetCell(2).NumericCellValue);
        }

        [Fact]
        public void StartingCell_Negative_ShouldThrow()
        {
            IExcelPrepareStage<TestDto> prepare = CreateBuilder();

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                prepare.StartingCell(-1));
        }

        // CONDITIONS

        [Fact]
        public void Column_WithFalseCondition_ShouldNotBeWritten()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            builder
                .AddColumn("Id", x => x.Id, () => false)
                .WriteRow(new TestDto { Id = 1 });

            var row = _sheet.GetRow(1);

            Assert.Null(row.GetCell(0));
        }

        // CLIENT INFO

        [Fact]
        public void AddClientInfo_ShouldWriteHeaderBlock()
        {
            IExcelPrepareStage<TestDto> prepare = CreateBuilder();

            var rowMode = prepare.AddClientInfo("Client A", null);

            rowMode
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 1 });

            var row = _sheet.GetRow(0);

            Assert.Equal("Client A", row.GetCell(0).StringCellValue);
        }

        // CUSTOM BLOCKS (TOP)

        [Fact]
        public void TopCustomBlock_ShouldBeRendered()
        {
            IExcelPrepareStage<TestDto> prepare = CreateBuilder();

            var rowMode = prepare
                .AddCustomBlock(w => w.WriteCell(0, "Top"))
                .AddClientInfo("Client", null);

            rowMode
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 1 });

            var hasTopBlock = Enumerable.Range(0, 5)
                .Select(i => _sheet.GetRow(i))
                .Where(r => r != null)
                .Select(r => r.GetCell(0)?.ToString())
                .Any(v => v == "Top");

            Assert.True(hasTopBlock);
        }

        // CUSTOM BLOCKS (BOTTOM)

        [Fact]
        public void BottomCustomBlock_ShouldBeRendered()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            var final = builder
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 1 });

            final.AddCustomBlock(w => w.WriteCell(0, "Bottom"));

            var lastRow = _sheet.LastRowNum;
            var cell = _sheet.GetRow(lastRow).GetCell(0);

            Assert.Equal("Bottom", cell.StringCellValue);
        }

        // TABLE STYLE

        [Fact]
        public void AddTableStyle_DuplicateHeaders_ShouldThrow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            var final = builder
                .AddColumn("Same", x => x.Id)
                .AddColumn("Same", x => x.Name)
                .WriteRow(new TestDto());

            Assert.Throws<Exception>(() => final.AddTableStyle());
        }

        [Fact]
        public void AddTableStyle_ShouldNotThrow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            var final = builder
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 1 });

            final.AddTableStyle();
        }

        // AUTOSIZE

        [Fact]
        public void AutoSizeColumns_ShouldNotThrow()
        {
            IExcelRowMode<TestDto> builder = CreateBuilder();

            var final = builder
                .AddColumn("Id", x => x.Id)
                .WriteRow(new TestDto { Id = 1 });

            final.AutoSizeColumns();
        }
    }
}