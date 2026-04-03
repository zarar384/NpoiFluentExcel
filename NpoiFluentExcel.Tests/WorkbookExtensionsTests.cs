using NPOI.XSSF.UserModel;
using NpoiFluentExcel.Extensions;
using NpoiFluentExcel.Mapping.Attributes;

namespace NpoiFluentExcel.Tests
{
    public class WorkbookExtensionsTests
    {
        private readonly XSSFWorkbook _workbook;

        public WorkbookExtensionsTests()
        {
            // in-memory workbook for testing
            _workbook = new XSSFWorkbook();
        }

        // Test DTOs

        [Sheet("TestSheet")]
        private class ValidDto
        {
            [Column("Id")]
            public int Id { get; set; }

            [Column("Name")]
            public string Name { get; set; }
        }

        private class NoSheetDto
        {
            [Column("Id")]
            public int Id { get; set; }
        }

        [Sheet("")]
        private class EmptySheetNameDto
        {
            [Column("Id")]
            public int Id { get; set; }
        }

        [Sheet("NoColumns")]
        private class NoColumnsDto
        {
            public int Id { get; set; }
        }

        [Sheet("BadColumn")]
        private class BadColumnDto
        {
            [Column("")]
            public int Id { get; set; }
        }

        [Sheet("NoGetter")]
        private class NoGetterDto
        {
            [Column("Id")]
            public int Id { private get; set; }
        }

        // AddSheet

        [Fact]
        public void AddSheet_ShouldCreateSheet()
        {
            var table = _workbook.AddSheet<ValidDto>("MySheet");

            var sheet = _workbook.GetSheet("MySheet");

            Assert.NotNull(sheet);
            Assert.NotNull(table);
        }

        [Fact]
        public void AddSheet_ShouldReturnExistingSheet()
        {
            _workbook.CreateSheet("MySheet");

            var table = _workbook.AddSheet<ValidDto>("MySheet");

            Assert.Equal(1, _workbook.NumberOfSheets);
            Assert.NotNull(table);
        }

        // WriteSheetColumns

        [Fact]
        public void WriteSheetColumns_ShouldCreateSheetAndColumns()
        {
            var table = _workbook.WriteSheetColumns<ValidDto>();

            var sheet = _workbook.GetSheet("TestSheet");

            Assert.NotNull(sheet);
            Assert.NotNull(table);
        }

        [Fact]
        public void WriteSheetColumns_NoSheetAttribute_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheetColumns<NoSheetDto>());

            Assert.Contains("does not define a SheetAttribute", ex.Message);
        }

        [Fact]
        public void WriteSheetColumns_EmptySheetName_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheetColumns<EmptySheetNameDto>());

            Assert.Contains("empty sheet name", ex.Message);
        }

        [Fact]
        public void WriteSheetColumns_NoColumns_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheetColumns<NoColumnsDto>());

            Assert.Contains("does not contain any properties", ex.Message);
        }

        [Fact]
        public void WriteSheetColumns_EmptyColumnName_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheetColumns<BadColumnDto>());

            Assert.Contains("empty column name", ex.Message);
        }

        [Fact]
        public void WriteSheetColumns_PropertyWithoutGetter_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheetColumns<NoGetterDto>());

            Assert.Contains("does not have a public getter", ex.Message);
        }

        // WriteSheet

        [Fact]
        public void WriteSheet_ShouldWriteData()
        {
            var data = new List<ValidDto>
            {
                new ValidDto { Id = 1, Name = "A" },
                new ValidDto { Id = 2, Name = "B" }
            };

            var result = _workbook.WriteSheet(data);

            Assert.NotNull(result);

            var sheet = _workbook.GetSheet("TestSheet");
            Assert.NotNull(sheet);
        }

        [Fact]
        public void WriteSheet_NullItems_ShouldThrow()
        {
            var ex = Assert.Throws<InvalidOperationException>(() =>
                _workbook.WriteSheet<ValidDto>(null));

            Assert.Contains("cannot be null", ex.Message);
        }

        // DeleteEmptySheets

        [Fact]
        public void DeleteEmptySheets_ShouldRemoveEmptySheets()
        {
            _workbook.CreateSheet("Empty1");
            _workbook.CreateSheet("Empty2");

            _workbook.DeleteEmptySheets();

            Assert.Equal(0, _workbook.NumberOfSheets);
        }

        [Fact]
        public void DeleteEmptySheets_ShouldRemoveSingleRowSheet()
        {
            var sheet = _workbook.CreateSheet("Data");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Test");

            _workbook.DeleteEmptySheets();

            Assert.Equal(0, _workbook.NumberOfSheets);
        }

        [Fact]
        public void DeleteEmptySheets_ShouldKeepSheetsWithMoreThanOneRow()
        {
            var sheet = _workbook.CreateSheet("Data");
            sheet.CreateRow(0);
            sheet.CreateRow(1);

            _workbook.DeleteEmptySheets();

            Assert.Equal(1, _workbook.NumberOfSheets);
        }

        // GetNextTableId 

        [Fact]
        public void GetNextTableId_EmptyWorkbook_ShouldReturn1()
        {
            var id = _workbook.GetNextTableId();

            Assert.Equal((uint)1, id);
        }

        [Fact]
        public void GetNextTableId_WithSheetsWithoutTables_ShouldReturn1()
        {
            _workbook.CreateSheet("Sheet1");
            _workbook.CreateSheet("Sheet2");

            var id = _workbook.GetNextTableId();

            Assert.Equal((uint)1, id);
        }
    }
}