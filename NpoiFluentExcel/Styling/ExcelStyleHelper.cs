using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NpoiFluentExcel.Styling
{
    public class ExcelStyleHelper
    {
        private readonly IWorkbook _workbook;

        public IFont NormalFont { get; }
        public IFont BoldFont { get; }
        public IFont HeaderFont { get; }

        public ICellStyle Normal { get; }
        public ICellStyle Bold { get; }
        public ICellStyle Header { get; }
        public ICellStyle Date { get; }
        public ICellStyle Number { get; }
        public ICellStyle Integer { get; }
        public ICellStyle BoldNumber { get; }
        public ICellStyle BoldInteger { get; }

        public ExcelStyleHelper(IWorkbook workbook, string fontName = "Calibri", short headerFontSize = 16, short normalFontSize = 9)
        {
            _workbook = workbook;

            NormalFont = CreateFont(fontName, normalFontSize, false);
            BoldFont = CreateFont(fontName, normalFontSize, true);
            HeaderFont = CreateFont(fontName, headerFontSize, true);

            Normal = CreateStyle(NormalFont, HorizontalAlignment.Left);
            Bold = CreateStyle(BoldFont, HorizontalAlignment.Left);
            Header = CreateStyle(HeaderFont, HorizontalAlignment.Center);

            Date = CreateStyle(NormalFont, HorizontalAlignment.Right);
            Date.DataFormat = HSSFDataFormat.GetBuiltinFormat("m/d/yy");

            Number = CreateStyle(NormalFont, HorizontalAlignment.Right);
            Number.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");

            Integer = CreateStyle(NormalFont, HorizontalAlignment.Right);
            Integer.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");

            BoldNumber = CreateStyle(BoldFont, HorizontalAlignment.Right);
            BoldNumber.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0.00");

            BoldInteger = CreateStyle(BoldFont, HorizontalAlignment.Right);
            BoldInteger.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");
        }

        private IFont CreateFont(string name, short size, bool bold)
        {
            var font = _workbook.CreateFont();
            font.FontName = name;
            font.FontHeightInPoints = size;
            font.IsBold = bold;
            return font;
        }

        private ICellStyle CreateStyle(IFont font, HorizontalAlignment alignment)
        {
            var style = _workbook.CreateCellStyle();
            style.SetFont(font);
            style.Alignment = alignment;
            return style;
        }

        public void WriteCell(ICell cell, double? value, ICellStyle style)
        {
            if (!value.HasValue)
                return;

            cell.SetCellType(CellType.Numeric);
            cell.SetCellValue(value.Value);
            cell.CellStyle = style;
        }

        public void WriteCell(ICell cell, string value, ICellStyle style, bool wrapText = false)
        {
            cell.SetCellType(CellType.String);
            cell.SetCellValue(value ?? string.Empty);

            if (wrapText)
            {
                var newStyle = cell.Sheet.Workbook.CreateCellStyle();
                newStyle.CloneStyleFrom(style);
                newStyle.WrapText = true;
                cell.CellStyle = newStyle;
            }
            else
            {
                cell.CellStyle = style;
            }
        }

        public void WriteCellSafe(IRow row, int index, double? value, ICellStyle style = null)
        {
            var cell = GetOrCreateCell(row, index);

            if (value.HasValue)
            {
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(value.Value);
            }

            if (style != null)
                cell.CellStyle = style;
        }

        public void WriteCellSafe(IRow row, int index, string value, ICellStyle style)
        {
            var cell = GetOrCreateCell(row, index);

            cell.SetCellType(CellType.String);
            cell.SetCellValue(value ?? string.Empty);
            cell.CellStyle = style;
        }

        public void WriteCellSafe(IRow row, int index, DateTime? value, ICellStyle style)
        {
            var cell = GetOrCreateCell(row, index);

            if (value.HasValue)
                cell.SetCellValue(value.Value);

            cell.CellStyle = style;
        }

        public void WriteCellSafe(IRow row, int index, object value, ICellStyle style = null)
        {
            if (value == null)
            {
                WriteCellSafe(row, index, string.Empty, style ?? Normal);
                return;
            }

            switch (value)
            {
                case double d:
                    WriteCellSafe(row, index, d, style ?? Number);
                    break;

                case int i:
                    WriteCellSafe(row, index, (double)i, style ?? Integer);
                    break;

                case decimal m:
                    WriteCellSafe(row, index, (double)m, style ?? Number);
                    break;

                case long l:
                    WriteCellSafe(row, index, (double)l, style ?? Integer);
                    break;

                case float f: 
                    WriteCellSafe(row, index, (double)f, style ?? Number);
                    break;

                case DateTime dt:
                    WriteCellSafe(row, index, dt, style ?? Date);
                    break;

                default:
                    WriteCellSafe(row, index, value.ToString(), style ?? Normal);
                    break;
            }
        }

        public void WriteCellFormulaSafe(IRow row, int index, string formula, ICellStyle style)
        {
            var cell = GetOrCreateCell(row, index);

            cell.SetCellType(CellType.Formula);
            cell.SetCellFormula(formula);
            cell.CellStyle = style;
        }

        public void InsertClientLogo(ICell cell, string path)
        {
            if (string.IsNullOrEmpty(path))
                return;

            var workbook = cell.Sheet.Workbook;
            var drawing = cell.Sheet.CreateDrawingPatriarch();

            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex - 3;
            anchor.Row1 = cell.RowIndex;
            anchor.Col2 = cell.ColumnIndex - 1;
            anchor.Row2 = cell.RowIndex;
            anchor.Dx2 = (int)(5.03 * 360000);
            anchor.Dy2 = (int)(1.06 * 360000);
            anchor.AnchorType = AnchorType.DontMoveAndResize;

            var pictureIdx = workbook.AddPicture(System.IO.File.ReadAllBytes(path), PictureType.PNG);
            drawing.CreatePicture(anchor, pictureIdx);
        }

        public ICell GetOrCreateCell(IRow row, int index)
        {
            return row.GetCell(index) ?? row.CreateCell(index);
        }
    }
}