using NpoiFluentExcel.Core;

namespace NpoiFluentExcel.Interfaces
{
    /// <summary>
    /// Preparation stage for row-based Excel writing.
    /// Allows configuring layout before defining columns.
    /// </summary>
    public interface IExcelPrepareStage<T>
    {
        /// <summary>
        /// Adds a custom block above the table (between client info and header).
        /// </summary>
        IExcelPrepareStage<T> AddCustomBlock(Action<ExcelCustomBlockWriter> builder);

        /// <summary>
        /// Sets the starting column index for the table.
        /// </summary>
        IExcelPrepareStage<T> StartingCell(int column);

        /// <summary>
        /// Adds client information (e.g., name, logo) above the table.
        /// </summary>
        IExcelRowMode<T> AddClientInfo(string clientInfo, string logoPath);
    }
}