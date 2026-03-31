using NPOI.SS.UserModel;
using NpoiFluentExcel.Styling;

namespace NpoiFluentExcel.Interfaces
{
    /// <summary>
    /// Row-based Excel writing mode.
    /// Columns are defined first, then data is written row by row.
    /// </summary>
    public interface IExcelRowMode<T>
    {
        /// <summary>
        /// Registers a column.
        /// </summary>
        IExcelRowMode<T> AddColumn(
            string header,
            Func<T, object> valueSelector,
            Func<bool> condition = null);

        /// <summary>
        /// Registers a column with a specific cell style.
        /// </summary>
        IExcelRowMode<T> AddColumn(
            string header,
            Func<T, object> valueSelector,
            ICellStyle style,
            Func<bool> condition = null);

        /// <summary>
        /// Registers a column using a custom writer.
        /// </summary>
        IExcelRowMode<T> AddColumn(
            string header,
            Action<IRow, T, int, ExcelStyleHelper> writer,
            Func<bool> condition = null);

        /// <summary>
        /// Marks the last defined column as a summation column.
        /// A total formula will be generated at the end of the table.
        /// </summary>
        IExcelRowMode<T> Sum();

        /// <summary>
        /// Writes a collection of rows to Excel.
        /// Can be called multiple times (supports batching).
        /// </summary>
        IExcelWriteStage<T> WriteRows(IEnumerable<T> rows);

        /// <summary>
        /// Writes a single row to Excel.
        /// Useful for streaming scenarios.
        /// </summary>
        IExcelWriteStage<T> WriteRow(T row);
    }
}