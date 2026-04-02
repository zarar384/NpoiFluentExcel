using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NpoiFluentExcel.Core;
using NpoiFluentExcel.Extensions;
using NpoiFluentExcel.Interfaces;
using NpoiFluentExcel.Styling;

namespace NpoiFluentExcel.Builders
{
    /// <summary>
    /// Fluent builder for writing tabular data into Excel (row-based mode).
    /// </summary>
    public sealed class ExcelTableBuilder<T> :
        IExcelPrepareStage<T>,
        IExcelRowMode<T>,
        IExcelWriteStage<T>,
        IExcelFinalStage
    {
        private readonly ISheet _sheet;
        private readonly ExcelStyleHelper _style;

        private int _currentRow = 1; // 0 = header
        private int _headerRowIndex = 0;

        private bool _headerWritten = false;
        private bool _isTableStyleApplied = false;

        private int _startColumn = 0;

        private string _clientInfo;
        private string _logoPath;

        private readonly List<Action<ExcelCustomBlockWriter>> _topBlocks = new();
        private readonly List<Action<ExcelCustomBlockWriter>> _bottomBlocks = new();

        private IColumnDefinition _lastColumn;
        private readonly List<IColumnDefinition> _columns = new();

        public ExcelTableBuilder(ISheet sheet, XSSFWorkbook workbook)
        {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _style = new ExcelStyleHelper(workbook);
        }

        #region Column Definitions

        private interface IColumnDefinition
        {
            int Index { get; }
            string Header { get; }
            bool IsEnabled();
            void SetIndex(int index);
            void Write(IRow row, object item, ExcelStyleHelper helper);
            bool IsSum { get; set; }
        }

        private sealed class ColumnDefinition<TItem> : IColumnDefinition
        {
            public string Header { get; }
            private readonly Action<IRow, TItem, int, ExcelStyleHelper> _writer;
            private readonly Func<bool> _condition;

            public int Index { get; private set; }
            public bool IsSum { get; set; }

            public ColumnDefinition(
                string header,
                Action<IRow, TItem, int, ExcelStyleHelper> writer,
                Func<bool> condition)
            {
                Header = header;
                _writer = writer ?? throw new ArgumentNullException(nameof(writer));
                _condition = condition;
            }

            public bool IsEnabled() => _condition == null || _condition();

            public void SetIndex(int index) => Index = index;

            public void Write(IRow row, object item, ExcelStyleHelper helper)
            {
                _writer(row, (TItem)item, Index, helper);
            }
        }

        #endregion

        #region Internal Logic

        private void WriteHeaderIfNeeded()
        {
            if (_headerWritten)
                return;

            var headerRowIndex = 0;

            if (!string.IsNullOrEmpty(_clientInfo))
            {
                var row = _sheet.GetRow(headerRowIndex) ?? _sheet.CreateRow(headerRowIndex);

                var lastColIndex = _startColumn + _columns.Count(c => c.IsEnabled());
                var logoCell = row.CreateCell(lastColIndex);

                if (!string.IsNullOrEmpty(_logoPath))
                    _style.InsertClientLogo(logoCell, _logoPath);

                var infoCell = row.CreateCell(_startColumn);
                _style.WriteCell(infoCell, _clientInfo, _style.Normal, wrapText: true);

                headerRowIndex = 2;
            }

            foreach (var block in _topBlocks)
            {
                var writer = new ExcelCustomBlockWriter(_sheet, _style, headerRowIndex, _startColumn);
                block(writer);
                headerRowIndex = writer.CurrentRow + 1;
            }

            var headerRow = _sheet.GetRow(headerRowIndex) ?? _sheet.CreateRow(headerRowIndex);

            int colIndex = _startColumn;
            foreach (var col in _columns)
            {
                if (!col.IsEnabled())
                    continue;

                col.SetIndex(colIndex);

                _style.WriteCell(
                    headerRow.CreateCell(colIndex),
                    col.Header,
                    _style.Bold
                );

                colIndex++;
            }

            _headerWritten = true;
            _headerRowIndex = headerRowIndex;
            _currentRow = _headerRowIndex + 1;
        }

        private void WriteRowInternal(T item)
        {
            var row = _sheet.GetRow(_currentRow) ?? _sheet.CreateRow(_currentRow);

            foreach (var col in _columns)
            {
                if (!col.IsEnabled())
                    continue;

                col.Write(row, item, _style);
            }

            _currentRow++;
        }

        private void WriteTotalRowIfNeeded()
        {
            var sumColumns = _columns.Where(c => c.IsEnabled() && c.IsSum).ToList();
            if (!sumColumns.Any())
                return;

            var row = _sheet.GetRow(_currentRow) ?? _sheet.CreateRow(_currentRow);

            if (_columns.FirstOrDefault(c => c.IsEnabled())?.IsSum != true)
            {
                var cell = row.CreateCell(_startColumn);
                _style.WriteCell(cell, "Total", _style.Bold);
            }

            int start = _headerRowIndex + 1;
            int end = _currentRow - 1;

            foreach (var col in sumColumns)
            {
                string letter = CellReference.ConvertNumToColString(col.Index);
                string formula = $"SUBTOTAL(109,{letter}{start + 1}:{letter}{end + 1})";

                _style.WriteCellFormulaSafe(row, col.Index, formula, _style.BoldNumber);
            }

            _currentRow++;
        }

        private void WriteBottomBlocksIfNeeded()
        {
            if (!_bottomBlocks.Any())
                return;

            int row = _currentRow;

            foreach (var block in _bottomBlocks)
            {
                var writer = new ExcelCustomBlockWriter(_sheet, _style, row, _startColumn);
                block(writer);
                row = writer.CurrentRow + 1;
            }

            _currentRow = row;
        }

        #endregion

        #region Fluent API

        public IExcelRowMode<T> AddColumn(
            string header,
            Func<T, object> selector,
            ICellStyle style,
            Func<bool> condition = null)
        {
            return AddColumn(
                header,
                (row, item, col, helper) =>
                {
                    var value = selector(item);
                    helper.WriteCellSafe(row, col, value, style);
                },
                condition
            );
        }

        public IExcelRowMode<T> AddColumn(
            string header,
            Func<T, object> selector,
            Func<bool> condition = null)
        {
            return AddColumn(header, selector, null, condition);
        }

        public IExcelRowMode<T> AddColumn(
            string header,
            Action<IRow, T, int, ExcelStyleHelper> writer,
            Func<bool> condition = null)
        {
            var col = new ColumnDefinition<T>(header, writer, condition);
            _columns.Add(col);
            _lastColumn = col;
            return this;
        }

        public IExcelWriteStage<T> WriteRows(IEnumerable<T> rows)
        {
            WriteHeaderIfNeeded();

            foreach (var row in rows)
                WriteRowInternal(row);

            WriteTotalRowIfNeeded();
            return this;
        }

        public IExcelWriteStage<T> WriteRow(T row)
        {
            WriteHeaderIfNeeded();
            WriteRowInternal(row);
            return this;
        }

        public IExcelRowMode<T> Sum()
        {
            if (_lastColumn == null)
                throw new InvalidOperationException("Sum() must follow AddColumn().");

            _lastColumn.IsSum = true;
            return this;
        }

        public IExcelFinalStage AddTableStyle(string name = "TableStyleDark9", bool disableAutoFilter = false)
        {
            if (_isTableStyleApplied)
                return this;

            if (!(_sheet is XSSFSheet xssfSheet))
                return this;

            var headerRow = _sheet.GetRow(_headerRowIndex);
            if (headerRow == null || _sheet.LastRowNum < _headerRowIndex + 1)
                return this;

            int visibleColumnCount = _columns.Count(c => c.IsEnabled());
            if (visibleColumnCount == 0)
                return this;

            int lastRowIndex = _sheet.LastRowNum;
            int lastColIndex = _startColumn + visibleColumnCount - 1;
            if (lastColIndex < 0)
                return this;

            var tableId = xssfSheet.Workbook.GetNextTableId();
            string tableName = $"Table{tableId}";

            var isCelkemRowExists = _columns.Any(c => c.IsEnabled() && c.IsSum);

            var table = xssfSheet.CreateTable();
            var ctTable = table.GetCTTable();

            var tableEndRow = isCelkemRowExists ? lastRowIndex - 1 : lastRowIndex;

            var area = new AreaReference(
                new CellReference(_headerRowIndex, _startColumn),
                new CellReference(tableEndRow, lastColIndex) // isCelkemRowExists ? lastColIndex - 1 : 
            );

            ctTable.@ref = area.FormatAsString();
            ctTable.id = tableId;
            ctTable.name = tableName;
            ctTable.displayName = tableName;
            ctTable.headerRowCount = 1;

            // TODO doesnt work
            //ctTable.totalsRowShown = isCelkemRowExists;
            //ctTable.totalsRowCount = isCelkemRowExists ? 1U : 0U;

            ctTable.tableStyleInfo = new CT_TableStyleInfo
            {
                name = name,
                showRowStripes = true,
                showColumnStripes = false
            };

            var tableColumns = new CT_TableColumns
            {
                count = (uint)visibleColumnCount,
                tableColumn = new List<CT_TableColumn>()
            };


            var headerNames = new List<string>();
            for (int i = 0; i < visibleColumnCount; i++)
            {
                int colIndex = _startColumn + i;

                headerNames.Add(headerRow.GetCell(colIndex)?.ToString() ?? $"Column{i + 1}");
            }

            var isHeaderWithDuplicate = headerNames.GroupBy(h => h)
                .Any(g => g.Count() > 1);

            if (isHeaderWithDuplicate)
                throw new Exception($"Sheet contains duplicate column headers. Table style cannot be applied. Headers: {string.Join(", ", headerNames)}");

            for (int i = 0; i < headerNames.Count; i++)
            {
                var column = new CT_TableColumn
                {
                    id = (uint)(i + 1),
                    name = headerNames[i]
                };

                tableColumns.tableColumn.Add(column);
            }

            ctTable.tableColumns = tableColumns;

            if (!disableAutoFilter)
            {
                ctTable.autoFilter = new CT_AutoFilter
                {
                    @ref = area.FormatAsString(),
                };
            }

            _isTableStyleApplied = true;
            return this;
        }

        public IExcelFinalStage AutoSizeColumns()
        {
            foreach (var col in _columns)
            {
                if (col.IsEnabled())
                    _sheet.AutoSizeColumn(col.Index);
            }

            return this;
        }

        public IExcelRowMode<T> AddClientInfo(string info, string logoPath)
        {
            _clientInfo = info;
            _logoPath = logoPath;
            return this;
        }

        public IExcelPrepareStage<T> StartingCell(int col)
        {
            if (col < 0)
                throw new ArgumentOutOfRangeException(nameof(col));

            _startColumn = col;
            return this;
        }

        public IExcelPrepareStage<T> AddCustomBlock(Action<ExcelCustomBlockWriter> block)
        {
            _topBlocks.Add(block);
            return this;
        }

        IExcelFinalStage IExcelFinalStage.AddCustomBlock(Action<ExcelCustomBlockWriter> block)
        {
            _bottomBlocks.Add(block);
            WriteBottomBlocksIfNeeded();
            return this;
        }

        #endregion
    }
}