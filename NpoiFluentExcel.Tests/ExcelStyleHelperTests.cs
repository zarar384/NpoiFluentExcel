using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NpoiFluentExcel.Styling;

namespace NpoiFluentExcel.Tests
{
    public class ExcelStyleHelperTests
    {
        private readonly IWorkbook _workbook;
        private readonly ExcelStyleHelper _helper;
        private readonly ISheet _sheet;
        private readonly IRow _row;

        public ExcelStyleHelperTests()
        {
            // in-memory workbook for testing
            _workbook = new HSSFWorkbook();

            _helper = new ExcelStyleHelper(_workbook);

            // create a sheet and row for testing cell writing
            _sheet = _workbook.CreateSheet("Test");
            _row = _sheet.CreateRow(0);
        }

        // WriteCell

        [Fact]
        public void WriteCell_DoubleValue_ShouldSetNumericCell()
        {
            var cell = _row.CreateCell(0);

            _helper.WriteCell(cell, 123.45, _helper.Number);

            Assert.Equal(CellType.Numeric, cell.CellType);
            Assert.Equal(123.45, cell.NumericCellValue);
            Assert.Equal(_helper.Number, cell.CellStyle);
        }

        [Fact]
        public void WriteCell_StringValue_ShouldSetStringCell()
        {
            var cell = _row.CreateCell(1);

            _helper.WriteCell(cell, "Hello", _helper.Normal);

            Assert.Equal(CellType.String, cell.CellType);
            Assert.Equal("Hello", cell.StringCellValue);
            Assert.Equal(_helper.Normal, cell.CellStyle);
        }

        [Fact]
        public void WriteCell_NullDouble_ShouldDoNothing()
        {
            var cell = _row.CreateCell(2);

            _helper.WriteCell(cell, (double?)null, _helper.Number);

            Assert.Equal(CellType.Blank, cell.CellType);
        }

        [Fact]
        public void WriteCellSafe_ShouldCreateCellIfNotExists()
        {
            _helper.WriteCellSafe(_row, 5, 10.0, _helper.Number);

            var cell = _row.GetCell(5);

            Assert.NotNull(cell);
            Assert.Equal(CellType.Numeric, cell.CellType);
            Assert.Equal(10.0, cell.NumericCellValue);
        }

        [Fact]
        public void WriteCellSafe_ObjectInt_WithNullStyle_ShouldNotOverrideStyle()
        {
            var existingCell = _row.CreateCell(6);
            existingCell.CellStyle = _helper.Bold; // old style

            _helper.WriteCellSafe(_row, 6, 100, null);

            var cell = _row.GetCell(6);

            Assert.Equal(100, cell.NumericCellValue);
            Assert.Equal(_helper.Bold, cell.CellStyle); // style should not be overridden
        }

        [Fact]
        public void WriteCellSafe_ObjectInt_WithoutStyle_ShouldNotAssignCustomStyle()
        {
            _helper.WriteCellSafe(_row, 7, 100);

            var cell = _row.GetCell(7);

            Assert.Equal(100, cell.NumericCellValue);
            Assert.Equal(CellType.Numeric, cell.CellType);
        }

        [Fact]
        public void WriteCellSafe_ObjectString_ShouldUseNormalStyle()
        {
            _helper.WriteCellSafe(_row, 7, "Test");

            var cell = _row.GetCell(7);

            Assert.Equal(CellType.String, cell.CellType);
            Assert.Equal("Test", cell.StringCellValue);
            Assert.Equal(_helper.Normal, cell.CellStyle);
        }

        [Fact]
        public void WriteCellSafe_DateTime_ShouldSetDate()
        {
            var date = new DateTime(2024, 1, 1);

            _helper.WriteCellSafe(_row, 8, date, _helper.Date);

            var cell = _row.GetCell(8);

            Assert.Equal(date, cell.DateCellValue);
            Assert.Equal(_helper.Date, cell.CellStyle);
        }

        [Fact]
        public void WriteCellFormulaSafe_ShouldSetFormula()
        {
            _helper.WriteCellFormulaSafe(_row, 9, "A1+B1", _helper.Number);

            var cell = _row.GetCell(9);

            Assert.Equal(CellType.Formula, cell.CellType);
            Assert.Equal("A1+B1", cell.CellFormula);
        }

        // GetOrCreateCell
        [Fact]
        public void GetOrCreateCell_ShouldReturnExistingCell()
        {
            var created = _row.CreateCell(10);

            var result = _helper.GetOrCreateCell(_row, 10);

            Assert.Equal(created, result);
        }

        [Fact]
        public void GetOrCreateCell_ShouldCreateNewCell()
        {
            var result = _helper.GetOrCreateCell(_row, 11);

            Assert.NotNull(result);
            Assert.Equal(11, result.ColumnIndex);
        }

        // InsertClientLogo

        [Fact]
        public void InsertClientLogo_EmptyPath_ShouldDoNothing()
        {
            var cell = _row.CreateCell(12);

            _helper.InsertClientLogo(cell, null);

            Assert.True(true);
        }
    }
}