using NpoiFluentExcel.Core;

namespace NpoiFluentExcel.Interfaces
{
    /// <summary>
    /// Final stage of the Excel builder.
    /// Allows applying styles and finishing the table.
    /// </summary>
    public interface IExcelFinalStage
    {
        /// <summary>
        /// Applies a table style with optional auto-filter.
        /// </summary>
        IExcelFinalStage AddTableStyle(string styleName = "TableStyleDark9", bool disableAutoFilter = false);

        /// <summary>
        /// Automatically resizes columns based on content.
        /// </summary>
        IExcelFinalStage AutoSizeColumns();

        /// <summary>
        /// Adds a custom block below the table.
        /// </summary>
        IExcelFinalStage AddCustomBlock(Action<ExcelCustomBlockWriter> builder);
    }
}