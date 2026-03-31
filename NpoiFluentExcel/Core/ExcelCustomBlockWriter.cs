using NPOI.SS.UserModel;
using NpoiFluentExcel.Styling;

namespace NpoiFluentExcel.Core
{
    /// <summary>
    /// Helper for writing custom content blocks above or below a table.
    /// Allows writing values cell-by-cell and controlling row position.
    /// </summary>
    public sealed class ExcelCustomBlockWriter
    {
        private readonly ISheet _sheet;
        private readonly ExcelStyleHelper _style;

        /// <summary>
        /// Current row index where content is being written.
        /// </summary>
        public int CurrentRow { get; private set; }

        /// <summary>
        /// Starting column index for writing.
        /// </summary>
        public int StartColumn { get; }

        public ExcelStyleHelper Style => _style;

        public ExcelCustomBlockWriter(ISheet sheet, ExcelStyleHelper style, int startRow, int startColumn)
        {
            _sheet = sheet;
            _style = style;
            CurrentRow = startRow;
            StartColumn = startColumn;
        }

        /// <summary>
        /// Writes a value to a cell in the current row with column offset.
        /// </summary>
        public void WriteCell(int columnOffset, object value, ICellStyle style = null)
        {
            var row = _sheet.GetRow(CurrentRow) ?? _sheet.CreateRow(CurrentRow);
            _style.WriteCellSafe(row, StartColumn + columnOffset, value, style);
        }

        /// <summary>
        /// Moves the cursor to the next row (or multiple rows).
        /// </summary>
        public void NextRow(int count = 1)
        {
            CurrentRow += count;
        }
    }
}