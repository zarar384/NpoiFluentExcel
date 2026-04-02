using NPOI.XSSF.UserModel;

namespace NpoiFluentExcel.ConsoleDemo.Infrastructure
{
    public static class ExcelSaver
    {
        public static void Save(XSSFWorkbook workbook, string filePath)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));

            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path is empty.", nameof(filePath));

            var directory = Path.GetDirectoryName(filePath);

            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            workbook.Write(fileStream);
        }
    }
}
