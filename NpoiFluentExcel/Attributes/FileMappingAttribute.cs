using NPOI.SS.UserModel;

namespace NpoiFluentExcel.Mapping.Attributes
{
    /// <summary>
    /// Defines Excel sheet mapping for a DTO.
    /// Used to map a DTO to a specific worksheet during import or export.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SheetAttribute : Attribute
    {
        /// <summary>
        /// Name of the Excel sheet. 
        /// If not specified, the first sheet will be used.
        /// </summary>
        public string SheetName { get; }

        public SheetAttribute()
        {
        }

        public SheetAttribute(string sheetName)
        {
            SheetName = sheetName;
        }
    }

    /// <summary>
    /// Defines mapping between a DTO property and an Excel column.
    /// Supports both import and export scenarios.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ColumnAttribute : Attribute
    {
        /// <summary>
        /// Name of the Excel column.
        /// </summary>
        public string ColumnName { get; }

        /// <summary>
        /// Indicates whether the value is required during import.
        /// </summary>
        public bool IsRequired { get; set; } = false;

        /// <summary>
        /// Name of a custom method used to transform the value during import.
        /// </summary>
        public string ConverterMethod { get; set; }

        /// <summary>
        /// Type that contains the converter method.
        /// </summary>
        public Type ConverterType { get; set; }

        /// <summary>
        /// Cell style applied during export.
        /// </summary>
        public ICellStyle CellStyle { get; set; }

        /// <summary>
        /// Determines whether the column is included during export.
        /// </summary>
        public bool Include { get; set; } = true;

        public ColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}