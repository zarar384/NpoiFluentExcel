using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NpoiFluentExcel.Mapping.Attributes;
using NpoiFluentExcel.Builders;
using NpoiFluentExcel.Interfaces;
using System.Reflection;

namespace NpoiFluentExcel.Extensions
{
    public static class WorkbookExtensions
    {
        /// <summary>
        /// Creates or gets a sheet and returns a fluent table builder for the specified type.
        /// </summary>
        public static ExcelTableBuilder<T> AddSheet<T>(
            this XSSFWorkbook workbook,
            string sheetName)
        {
            var sheet = workbook.GetSheet(sheetName)
                        ?? workbook.CreateSheet(sheetName);

            return new ExcelTableBuilder<T>(sheet, workbook);
        }

        /// <summary>
        /// Creates a sheet from DTO definition, maps columns using attributes and writes data.
        /// </summary>
        public static IExcelWriteStage<TDto> WriteSheet<TDto>(
            this XSSFWorkbook workbook,
            IEnumerable<TDto> items)
            where TDto : class, new()
        {
            if (items == null)
                throw new InvalidOperationException("Input collection cannot be null.");

            var table = workbook.WriteSheetColumns<TDto>();

            try
            {
                return table.WriteRows(items);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Unexpected error while writing data for DTO '{typeof(TDto).Name}'.", ex);
            }
        }

        /// <summary>
        /// Creates a sheet from DTO definition and maps columns using attributes.
        /// </summary>
        public static IExcelRowMode<TDto> WriteSheetColumns<TDto>(
            this XSSFWorkbook workbook)
            where TDto : class, new()
        {
            if (workbook == null)
                throw new ArgumentNullException("Workbook is not initialized.");

            var type = typeof(TDto);

            var sheetAttr = type.GetCustomAttribute<SheetAttribute>();
            if (sheetAttr == null)
                throw new InvalidOperationException(
                    $"DTO '{type.Name}' does not define a SheetAttribute.");

            if (string.IsNullOrWhiteSpace(sheetAttr.SheetName))
                throw new InvalidOperationException(
                    $"DTO '{type.Name}' has an empty sheet name.");

            var columns = type.GetProperties()
                .Select(p => (Property: p, Attr: p.GetCustomAttribute<ColumnAttribute>()))
                .Where(x => x.Attr != null)
                .ToArray();

            if (columns.Length == 0)
                throw new InvalidOperationException(
                    $"DTO '{type.Name}' does not contain any properties with ColumnAttribute.");

            var table = workbook.AddSheet<TDto>(sheetAttr.SheetName);

            foreach (var (prop, attr) in columns)
            {
                var getter = prop.GetGetMethod();
                if (getter == null || !getter.IsPublic)
                    throw new InvalidOperationException(
                        $"Property '{prop.Name}' in DTO '{type.Name}' does not have a public getter.");

                if (string.IsNullOrWhiteSpace(attr.ColumnName))
                    throw new InvalidOperationException(
                        $"Property '{prop.Name}' in DTO '{type.Name}' has an empty column name.");

                try
                {
                    table.AddColumn(
                        attr.ColumnName,
                        dto => prop.GetValue(dto),
                        attr.CellStyle,
                        () => attr.Include);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"DTO '{type.Name}': failed to map column '{attr.ColumnName}' to property '{prop.Name}'. Error: {ex.Message}");
                }
            }

            return table;
        }

        /// <summary>
        /// Removes empty sheets from the workbook.
        /// Sheets with only a single row (row 0) are also considered empty.
        /// </summary>
        public static void DeleteEmptySheets(this XSSFWorkbook workbook)
        {
            for (int i = workbook.NumberOfSheets - 1; i >= 0; i--)
            {
                var sheet = workbook.GetSheetAt(i);
                if (sheet.LastRowNum == 0)
                    workbook.RemoveSheetAt(i);
            }
        }

        /// <summary>
        /// Returns the next available table ID across all sheets.
        /// </summary>
        public static uint GetNextTableId(this IWorkbook workbook)
        {
            uint id = 0;

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                if (workbook.GetSheetAt(i) is XSSFSheet sheet)
                {
                    id += (uint)sheet.GetTables().Count;
                }
            }

            return id + 1;
        }
    }
}