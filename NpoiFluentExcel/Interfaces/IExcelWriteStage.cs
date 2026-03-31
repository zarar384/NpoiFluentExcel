namespace NpoiFluentExcel.Interfaces
{
    /// <summary>
    /// Data writing stage Exception.
    /// </summary>
    public interface IExcelWriteStage<T> : IExcelFinalStage
    {
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