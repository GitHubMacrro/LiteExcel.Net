using System.Collections;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace LiteExcel;

/// <summary>
/// 高层工作表模型。
/// 内部保存原始网格（含首行在内的所有行），不隐含"首行是表头"的语义；
/// 表头识别是 List&lt;T&gt;/DataTable 映射层的职责。
/// 高层坐标统一 1-based；写回低层 <see cref="SheetData"/> 时自动转换。
/// </summary>
public sealed class Worksheet
{
    private readonly List<List<Cell>> _grid = new();
    private readonly List<CellRange> _mergedRanges = new();

    /// <summary>工作表名 </summary>
    public string Name { get; set; } = "Sheet1";

    /// <summary>是否冻结首行 </summary>
    public bool FreezeHeader { get; set; }

    /// <summary>冻结行数（0 = 不冻结行）。FreezeHeader = true 等价于 1 </summary>
    public int FreezeRows { get; set; }

    /// <summary>冻结列数（0 = 不冻结列） </summary>
    public int FreezeColumns { get; set; }

    /// <summary>列宽（0-based 列索引 -> 宽度） </summary>
    public Dictionary<int, double>? ColumnWidths { get; set; }

    /// <summary>行高（0-based 行索引 -> 高度，磅） </summary>
    public Dictionary<int, double>? RowHeights { get; set; }

    /// <summary>表头样式（写出时作用于首行） </summary>
    public CellStyle? HeaderStyle { get; set; }

    /// <summary>全表默认样式（优先级最低） </summary>
    public CellStyle? DefaultStyle { get; set; }

    /// <summary>行级样式（0-based 行索引） </summary>
    public Dictionary<int, CellStyle>? RowStyles { get; set; }

    /// <summary>列级样式（0-based 列索引） </summary>
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }

    /// <summary>单元格批注（A1 格式引用 -> 文本） </summary>
    public Dictionary<string, string>? Comments { get; set; }

    /// <summary>数据验证规则 </summary>
    public List<DataValidation>? Validations { get; set; }

    /// <summary>自动筛选 </summary>
    public AutoFilter? Filter { get; set; }

    /// <summary>条件格式规则列表 </summary>
    public List<ConditionalFormat> ConditionalFormats { get; } = new();

    /// <summary>工作表宿主的 VBA 代码名（打开时捕获，保存时随 SheetData 写回 sheetPr@codeName） </summary>
    internal string? CodeName { get; set; }

    /// <summary>工作表图片（InCell / Floating） </summary>
    public List<WorksheetImage> Images { get; } = new();

    /// <summary>工作表保护（sheetProtection）。默认 null（无保护） </summary>
    public SheetProtection? Protection { get; set; }

    /// <summary>超级表（Table/ListObject）列表。打开含表文件时自动填充；新建用 <see cref="AddTable"/> </summary>
    public IReadOnlyList<XlTable> Tables => _tables;

    private readonly List<XlTable> _tables = new();

    /// <summary>合并区域（CellRange 为 0-based，与低层模型一致） </summary>
    public IReadOnlyList<CellRange> MergedRanges => _mergedRanges;

    /// <summary>整表单元格集合入口 </summary>
    public Cells Cells { get; }

    internal Worksheet()
    {
        Cells = new Cells(this);
    }

    internal Worksheet(string name) : this()
    {
        Name = name;
    }

    /// <summary>总行数（1-based 有效行数，0 表示空表） </summary>
    public int RowCount => _grid.Count;

    /// <summary>最大列数（1-based，0 表示空表） </summary>
    public int MaxColumn
    {
        get
        {
            int max = 0;
            foreach (var row in _grid)
                if (row.Count > max) max = row.Count;
            return max;
        }
    }

    // ── 单元格访问 ──

    /// <summary>添加一张浮动图片（以 row/column 左上角为锚点，默认按图片原始尺寸显示） </summary>
    public WorksheetImage AddImage(byte[] data, int row, int column, double? widthPx = null, double? heightPx = null,
        ImagePlacement placement = ImagePlacement.Floating, string? extension = null, string? name = null)
    {
        if (data is null || data.Length == 0)
            throw new ArgumentException("图片数据不能为空", nameof(data));
        if (row < 1 || column < 1)
            throw new ArgumentOutOfRangeException(nameof(row), "图片锚点行/列必须从 1 开始");

        var img = new WorksheetImage
        {
            Data = data,
            Row = row,
            Column = column,
            WidthPx = widthPx,
            HeightPx = heightPx,
            Placement = placement,
            Extension = extension,
            Name = name ?? $"图片 {Images.Count + 1}",
        };
        Images.Add(img);
        return img;
    }

    /// <summary>
    /// 添加一张浮动图片（高精度锚点：左上单元格 + EMU 偏移 + 显示尺寸 + 移动方式）。
    /// 仅 Floating；InCell 请用 row/column 重载。
    /// </summary>
    public WorksheetImage AddImage(byte[] data, ImageAnchor anchor,
        string? extension = null, string? name = null, string? altText = null)
    {
        if (data is null || data.Length == 0)
            throw new ArgumentException("图片数据不能为空", nameof(data));
        if (anchor is null)
            throw new ArgumentNullException(nameof(anchor));

        var (r, c) = CellRef.Parse(anchor.TopLeftCell);
        var img = new WorksheetImage
        {
            Data = data,
            Row = r + 1,
            Column = c + 1,
            WidthPx = anchor.WidthPixels,
            HeightPx = anchor.HeightPixels,
            Placement = ImagePlacement.Floating,
            Extension = extension,
            Name = name ?? $"图片 {Images.Count + 1}",
            Anchor = anchor,
            AltText = altText,
        };
        Images.Add(img);
        return img;
    }
    public Cell Cell(int row, int column)
    {
        if (row < 1 || column < 1)
            throw new ArgumentOutOfRangeException(nameof(row), $"行列必须从 1 开始，收到 ({row}, {column})");

        var stored = TryGetCell(row, column);
        if (stored is not null) return stored;

        return new Cell
        {
            Type = CellType.Empty,
            Owner = this,
            OwnerRow = row,
            OwnerCol = column,
        };
    }

    /// <summary>按 A1 地址访问单元格，如 "A1"、"B3" </summary>
    public Cell Cell(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("单元格地址不能为空", nameof(address));
        if (address.Contains(':'))
            throw new ArgumentException($"'{address}' 是区域地址，请使用 Range() 或 Cells 访问", nameof(address));

        var (row, col) = CellRef.Parse(address);
        return Cell(row + 1, col + 1);
    }

    /// <summary>按 A1 区域访问，如 "A1:D10"。也可用 "A1:A1" 表示单格区域 </summary>
    public ExcelRange Range(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("区域地址不能为空", nameof(address));

        if (!address.Contains(':'))
        {
            var (r, c) = CellRef.Parse(address);
            return new ExcelRange(this, r + 1, c + 1, r + 1, c + 1);
        }

        var parts = address.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"无效的区域地址：{address}", nameof(address));

        var (r1, c1) = CellRef.Parse(parts[0]);
        var (r2, c2) = CellRef.Parse(parts[1]);
        int fr = Math.Min(r1, r2) + 1, lr = Math.Max(r1, r2) + 1;
        int fc = Math.Min(c1, c2) + 1, lc = Math.Max(c1, c2) + 1;
        return new ExcelRange(this, fr, fc, lr, lc);
    }

    /// <summary>按 1-based 行列访问区域（含端点） </summary>
    public ExcelRange Range(int firstRow, int firstCol, int lastRow, int lastCol)
    {
        if (firstRow < 1 || firstCol < 1 || lastRow < firstRow || lastCol < firstCol)
            throw new ArgumentOutOfRangeException(nameof(firstRow), $"无效的区域坐标：({firstRow},{firstCol})-({lastRow},{lastCol})");
        return new ExcelRange(this, firstRow, firstCol, lastRow, lastCol);
    }

    /// <summary>设置单元格值（越界自动扩展网格）。value 为 null/DBNull 时写空单元格 </summary>
    public void SetValue(int row, int column, object? value)
    {
        var cell = Cell(row, column);
        cell.SetValue(value);
    }

    /// <summary>按 A1 地址设置单元格值 </summary>
    public void SetValue(string address, object? value)
    {
        var cell = Cell(address);
        cell.SetValue(value);
    }

    // ── 数据导入 ──

    /// <summary>
    /// 把 List&lt;T&gt; 导入本工作表：清空现有内容后从 A1 重建（首行为表头，与 <see cref="Excel.Write{T}"/> 一致）。
    /// </summary>
    public void ImportData<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IEnumerable<T> data, Action<WriteOptions<T>>? configure = null)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        var options = new WriteOptions<T>();
        configure?.Invoke(options);
        ImportSheet(XlsxWriter.ListToSheet(data, options), includeHeader: true);
    }

    /// <summary>
    /// 把 DataTable 导入本工作表：清空现有内容后从 A1 重建。首行默认写表头。
    /// </summary>
    public void ImportData(DataTable table, bool includeHeader = true)
    {
        if (table is null) throw new ArgumentNullException(nameof(table));
        ImportSheet(XlsxWriter.DataTableToSheet(table, Name), includeHeader);
    }

    private void ImportSheet(SheetData sheet, bool includeHeader)
    {
        _grid.Clear();
        _mergedRanges.Clear();

        if (sheet.FreezeHeader) FreezeHeader = true;
        if (sheet.FreezeRows > 0) FreezeRows = sheet.FreezeRows;
        if (sheet.FreezeColumns > 0) FreezeColumns = sheet.FreezeColumns;
        if (sheet.ColumnWidths is not null)
            ColumnWidths = sheet.ColumnWidths.Select((w, i) => (i, w)).ToDictionary(x => x.i, x => x.w);

        if (includeHeader && sheet.Headers is { Count: > 0 })
        {
            var headerRow = new List<Cell>(sheet.Headers.Count);
            foreach (var h in sheet.Headers)
                headerRow.Add(LiteExcel.Cell.FromText(h));
            SetRowCells(0, headerRow);
        }

        foreach (var row in sheet.Rows)
        {
            var list = new List<Cell>(row.Count);
            list.AddRange(row);
            _grid.Add(list);
        }

        RebindOwners();
    }

    // ── 超级表 ──

    /// <summary>
    /// 创建超级表（Excel 内置表）。覆盖区首行作为表头列名，至少需要表头 + 1 行数据。
    /// 样式使用 Excel 内置名（见 <see cref="TableStyleStyle"/>）；自定义名见 <paramref name="styleName"/>。
    /// </summary>
    public XlTable AddTable(string refAddress, string name, TableStyleStyle? style = null)
    {
        var tbl = CreateTableCore(refAddress, name);
        tbl.Style = style ?? TableStyleStyle.Medium9;
        AddTableInternal(tbl);
        return tbl;
    }

    /// <summary>创建超级表并指定任意样式名（未知名将经降级回调提示，Excel 打开时退化为无样式） </summary>
    public XlTable AddTable(string refAddress, string name, string styleName)
    {
        var tbl = CreateTableCore(refAddress, name);
        tbl.CustomStyleName = styleName;
        AddTableInternal(tbl);
        return tbl;
    }

    /// <summary>按表名删除超级表。存在则删除并返回 true，否则 false </summary>
    public bool RemoveTable(string name)
    {
        for (int i = 0; i < _tables.Count; i++)
        {
            if (string.Equals(_tables[i].Name, name, StringComparison.OrdinalIgnoreCase))
            {
                _tables.RemoveAt(i);
                return true;
            }
        }
        return false;
    }

    private XlTable CreateTableCore(string refAddress, string name)
    {
        if (string.IsNullOrWhiteSpace(refAddress))
            throw new ArgumentException("表区域不能为空", nameof(refAddress));
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("表名不能为空", nameof(name));
        ValidateTableName(name);

        var (firstRow, firstCol, lastRow, lastCol) = CellRef.ParseRange(refAddress);
        int rowCount = lastRow - firstRow + 1;
        int colCount = lastCol - firstCol + 1;
        if (rowCount < 2)
            throw new LiteExcelException("超级表至少需要表头 + 1 行数据（区域至少 2 行）");
        if (colCount < 1)
            throw new LiteExcelException("超级表至少需要 1 列");

        // 与既有表区域重叠校验
        foreach (var t in _tables)
        {
            var (tr0, tc0, tr1, tc1) = CellRef.ParseRange(t.Ref);
            if (firstRow <= tr1 && lastRow >= tr0 && firstCol <= tc1 && lastCol >= tc0)
                throw new LiteExcelException($"表区域与既有表「{t.Name}」重叠");
        }

        var tbl = new XlTable { Name = name, Ref = refAddress };
        // 首行文本作为列名（TryGetCell 为 1-based）
        for (int c = 0; c < colCount; c++)
        {
            var cell = TryGetCell(firstRow + 1, firstCol + c + 1);
            string colName = cell is not null && cell.Type == CellType.Text && !string.IsNullOrEmpty(cell.Text)
                ? cell.Text!
                : $"Column{c + 1}";
            tbl.AddColumn(new XlTableColumn { Name = colName });
        }
        return tbl;
    }

    private void AddTableInternal(XlTable tbl)
    {
        foreach (var t in _tables)
            if (string.Equals(t.Name, tbl.Name, StringComparison.OrdinalIgnoreCase))
                throw new LiteExcelException($"表名重复：{tbl.Name}");
        _tables.Add(tbl);
    }

    private static void ValidateTableName(string name)
    {
        // Excel 表名规则：不能以数字开头、不能含空格、不能是纯单元格地址、不能含 ". "、不能以 "_xlfn" 开头
        if (char.IsDigit(name[0]))
            throw new LiteExcelException($"表名不能以数字开头：{name}");
        if (name.Contains(" "))
            throw new LiteExcelException($"表名不能包含空格：{name}");
        if (CellRef.TryParse(name, out _))
            throw new LiteExcelException($"表名不能是单元格地址：{name}");
    }

    /// <summary>按表内现有内容估算自适应列宽，结果写入 <see cref="ColumnWidths"/>（0-based 列索引）。 </summary>
    public void AutoColumnWidths()
    {
        var sd = ToSheetData();
        XlsxWriter.AutoColumnWidths(sd);
        if (sd.ColumnWidths is not null)
            ColumnWidths = sd.ColumnWidths.Select((w, i) => (i, w)).ToDictionary(x => x.i, x => x.w);
    }

    // ── 合并 ──

    /// <summary>合并区域（1-based，含端点）。例如 Merge(1, 1, 2, 2) 合并 A1:B2 </summary>
    public void Merge(int firstRow, int firstCol, int lastRow, int lastCol)
    {
        if (firstRow < 1 || firstCol < 1 || lastRow < firstRow || lastCol < firstCol)
            throw new ArgumentOutOfRangeException(nameof(firstRow), $"无效的区域坐标：({firstRow},{firstCol})-({lastRow},{lastCol})");

        var r = new CellRange(firstRow - 1, lastRow - 1, firstCol - 1, lastCol - 1);
        if (!_mergedRanges.Any(m => m.FirstRow == r.FirstRow && m.LastRow == r.LastRow && m.FirstCol == r.FirstCol && m.LastCol == r.LastCol))
            _mergedRanges.Add(r);
    }

    /// <summary>合并区域（A1 地址），例如 Merge("A1:B2") </summary>
    public void Merge(string address)
    {
        var range = Range(address);
        Merge(range.FirstRow, range.FirstCol, range.LastRow, range.LastCol);
    }

    /// <summary>取消合并区域（A1 地址） </summary>
    public void Unmerge(string address)
    {
        var range = Range(address);
        for (int i = _mergedRanges.Count - 1; i >= 0; i--)
        {
            var m = _mergedRanges[i];
            if (m.FirstRow == range.FirstRow - 1 && m.LastRow == range.LastRow - 1 &&
                m.FirstCol == range.FirstCol - 1 && m.LastCol == range.LastCol - 1)
            {
                _mergedRanges.RemoveAt(i);
            }
        }
    }

    // ── 内部网格操作 ──

    internal Cell? TryGetCell(int row1, int col1)
    {
        if (row1 < 1 || col1 < 1 || row1 - 1 >= _grid.Count) return null;
        var row = _grid[row1 - 1];
        if (col1 - 1 >= row.Count) return null;
        return row[col1 - 1];
    }

    internal void SetCell(int row1, int col1, Cell cell)
    {
        if (row1 < 1 || col1 < 1)
            throw new ArgumentOutOfRangeException(nameof(row1), $"行列必须从 1 开始，收到 ({row1}, {col1})");

        while (_grid.Count < row1)
            _grid.Add(new List<Cell>());
        var row = _grid[row1 - 1];
        while (row.Count < col1)
            row.Add(new Cell { Type = CellType.Empty, Owner = this, OwnerRow = row1, OwnerCol = row.Count + 1 });
        row[col1 - 1] = cell;
    }

    /// <summary>单元格写回回调：把高层 Cell 落入网格（越界扩展） </summary>
    internal void OnCellChanged(Cell cell)
    {
        if (!ReferenceEquals(cell.Owner, this)) return;
        SetCell(cell.OwnerRow, cell.OwnerCol, cell);
    }

    /// <summary>遍历网格中已有的所有单元格（不含越界占位） </summary>
    internal IEnumerable<Cell> EnumerateStoredCells()
    {
        foreach (var row in _grid)
            foreach (var cell in row)
                yield return cell;
    }

    internal List<List<Cell>> Grid => _grid;

    /// <summary>转换回低层 SheetData（写回用） </summary>
    public SheetData ToSheetData()
    {
        var sheet = new SheetData
        {
            SheetName = Name,
            FreezeHeader = FreezeHeader,
            FreezeRows = FreezeRows,
            FreezeColumns = FreezeColumns,
            HeaderStyle = HeaderStyle,
            DefaultStyle = DefaultStyle,
            RowStyles = RowStyles,
            ColumnStyles = ColumnStyles,
            RowHeights = RowHeights,
            ColumnWidths = ToColumnWidthsList(ColumnWidths),
            Comments = Comments,
            Validations = Validations,
            Filter = Filter,
            CodeName = CodeName,
            ConditionalFormats = new List<ConditionalFormat>(ConditionalFormats),
            Protection = Protection,
            Tables = _tables.Select(t => CloneTable(t)).ToList(),
        };

        if (Images.Count > 0)
            sheet.Images = Images;

        foreach (var range in _mergedRanges)
            sheet.MergedRanges.Add(range);

        foreach (var row in _grid)
        {
            var cells = new Cell[row.Count];
            for (int i = 0; i < row.Count; i++)
                cells[i] = row[i];
            sheet.Rows.Add(cells);
        }

        return sheet;
    }

    /// <summary>
    /// P0-2: 把高层稀疏字典（key=0-based 列索引）转为低层 List&lt;double&gt;，按下标补齐，
    /// 未设置的列填 0（哨兵=无自定义列宽，与各写入器跳过 &lt;=0 的约定一致）。
    /// 旧实现 <c>.Select(kv =&gt; kv.Value).ToList()</c> 丢弃 key 且枚举顺序无保证，导致稀疏列宽错位。
    /// </summary>
    private static List<double>? ToColumnWidthsList(Dictionary<int, double>? widths)
    {
        if (widths is null || widths.Count == 0) return null;
        int maxKey = -1;
        foreach (var k in widths.Keys)
            if (k > maxKey) maxKey = k;
        if (maxKey < 0) return null;
        var list = new List<double>(maxKey + 1);
        for (int i = 0; i <= maxKey; i++) list.Add(0);
        foreach (var kv in widths)
            list[kv.Key] = kv.Value;
        return list;
    }

    /// <summary>从低层 SheetData 构建 Worksheet（读取时复用） </summary>
    internal static Worksheet FromSheetData(SheetData sheet)
    {
        var ws = new Worksheet(sheet.SheetName)
        {
            FreezeHeader = sheet.FreezeHeader,
            FreezeRows = sheet.FreezeRows,
            FreezeColumns = sheet.FreezeColumns,
            HeaderStyle = sheet.HeaderStyle,
            DefaultStyle = sheet.DefaultStyle,
            RowStyles = sheet.RowStyles,
            ColumnStyles = sheet.ColumnStyles,
            RowHeights = sheet.RowHeights,
            Comments = sheet.Comments,
            Validations = sheet.Validations,
            Filter = sheet.Filter,
            CodeName = sheet.CodeName,
            Protection = sheet.Protection,
        };

        if (sheet.Tables is { Count: > 0 })
        {
            foreach (var t in sheet.Tables)
                ws._tables.Add(CloneTable(t));
        }

        if (sheet.ConditionalFormats is { Count: > 0 })
        {
            foreach (var cf in sheet.ConditionalFormats)
                ws.ConditionalFormats.Add(cf);
        }

        if (sheet.Images is { Count: > 0 })
        {
            foreach (var img in sheet.Images)
                ws.Images.Add(img);
        }

        if (sheet.ColumnWidths is not null)
            ws.ColumnWidths = sheet.ColumnWidths.Select((w, i) => (i, w)).ToDictionary(x => x.i, x => x.w);

        // 低层 Headers + Rows 重组为原始网格（首行合并回数据区）
        if (sheet.Headers is { Count: > 0 })
        {
            var headerRow = new List<Cell>(sheet.Headers.Count);
            foreach (var h in sheet.Headers)
                headerRow.Add(LiteExcel.Cell.FromText(h));
            ws.SetRowCells(0, headerRow);
        }

        foreach (var row in sheet.Rows)
        {
            var list = new List<Cell>(row.Count);
            list.AddRange(row);
            ws.Grid.Add(list);
        }

        foreach (var range in sheet.MergedRanges)
            ws._mergedRanges.Add(range);

        ws.RebindOwners();
        return ws;
    }

    private void SetRowCells(int row0, List<Cell> cells)
    {
        while (_grid.Count <= row0)
            _grid.Add(new List<Cell>());
        _grid[row0] = cells;
    }

    private static XlTable CloneTable(XlTable t)
    {
        var clone = new XlTable
        {
            Name = t.Name,
            Ref = t.Ref,
            Style = t.Style,
            CustomStyleName = t.CustomStyleName,
            ShowRowStripes = t.ShowRowStripes,
            ShowFirstColumn = t.ShowFirstColumn,
            ShowLastColumn = t.ShowLastColumn,
            ShowColumnStripes = t.ShowColumnStripes,
            AutoFilter = t.AutoFilter,
            TotalsRowShown = t.TotalsRowShown,
            HeaderStyle = t.HeaderStyle?.Clone(),
        };
        foreach (var c in t.Columns)
            clone.AddColumn(c.Clone());
        return clone;
    }

    private void RebindOwners()    {
        for (int r = 0; r < _grid.Count; r++)
        {
            for (int c = 0; c < _grid[r].Count; c++)
            {
                var cell = _grid[r][c];
                cell.Owner = this;
                cell.OwnerRow = r + 1;
                cell.OwnerCol = c + 1;
            }
        }
    }

    /// <summary>填充合并区域内的非左上角单元格（FillMergedCells 选项使用） </summary>
    internal void FillMergedValues()
    {
        foreach (var m in _mergedRanges)
        {
            var topLeft = TryGetCell(m.FirstRow + 1, m.FirstCol + 1);
            if (topLeft is null || topLeft.IsEmpty) continue;
            for (int r = m.FirstRow; r <= m.LastRow; r++)
            {
                for (int c = m.FirstCol; c <= m.LastCol; c++)
                {
                    if (r == m.FirstRow && c == m.FirstCol) continue;
                    var cell = Cell(r + 1, c + 1);
                    cell.SetValue(topLeft.GetValue());
                    cell.NumberFormat = topLeft.NumberFormat;
                }
            }
        }
    }
}
