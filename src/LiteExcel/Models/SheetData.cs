namespace LiteExcel;

/// <summary>
/// 一张工作表的完整数据 
/// </summary>
public sealed class SheetData
{
    public string SheetName { get; set; } = "Sheet1";
    public List<string> Headers { get; set; } = new();
    public List<IReadOnlyList<Cell>> Rows { get; set; } = new();
    public List<CellRange> MergedRanges { get; set; } = new();
    public AutoFilter? Filter { get; set; }
    public bool FreezeHeader { get; set; }

    /// <summary>冻结行数（0 = 不冻结行） </summary>
    public int FreezeRows { get; set; }

    /// <summary>冻结列数（0 = 不冻结列） </summary>
    public int FreezeColumns { get; set; }
    public List<double>? ColumnWidths { get; set; }
    public CellStyle? HeaderStyle { get; set; }

    /// <summary>全表默认样式（优先级最低） </summary>
    public CellStyle? DefaultStyle { get; set; }

    /// <summary>行级样式（key = 0-based 行索引，对应 Rows） </summary>
    public Dictionary<int, CellStyle>? RowStyles { get; set; }

    /// <summary>列级样式（key = 0-based 列索引） </summary>
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }

    /// <summary>行高（key = 0-based 行索引，对应 Rows） 单位：磅（point） </summary>
    public Dictionary<int, double>? RowHeights { get; set; }

    /// <summary>单元格批注（key = A1 格式单元格引用，value = 批注文本） </summary>
    public Dictionary<string, string>? Comments { get; set; }

    /// <summary>数据验证规则列表 </summary>
    public List<DataValidation>? Validations { get; set; }

    /// <summary>条件格式规则列表（ConditionalFormat 目标范围 Sqref）</summary>
    public List<ConditionalFormat> ConditionalFormats { get; set; } = new();

    /// <summary>工作表宿主的 VBA 代码名（sheetPr@codeName）。带宏工作簿经保存后仍与 vbaProject 绑定，避免 Excel 重排文档模块 </summary>
    public string? CodeName { get; set; }

    /// <summary>工作表图片（InCell richData / Floating drawing） </summary>
    public List<WorksheetImage>? Images { get; set; }

    /// <summary>工作表保护（sheetProtection）。为空表示无保护 </summary>
    public SheetProtection? Protection { get; set; }

    /// <summary>超级表（Table/ListObject）列表 </summary>
    public List<XlTable> Tables { get; set; } = new();

    /// <summary>读取时捕获的 sheet tableParts rel id 列表（内部使用，写出时透传保留） </summary>
    internal List<string>? TablesRawRelIds { get; set; }

    /// <summary>
    /// 数据区起始的原始 1-based 行号（含表头 / 数据首行在内）。
    /// 默认 0 = 紧凑模式（行为同现状：首行从第 1 行写，前导空行被压缩）。
    /// &gt;0 时表示源文件首个实际行的原始行号；写出时第 1..FirstRowNumber-1 行补齐空行，
    /// 避免数据整体上移、与会话表/合并区等绝对行号引用错位。
    /// Rows / RowHeights / RowStyles 的 key 仍为 0-based、对应 Rows，不受本字段影响。
    /// </summary>
    public int FirstRowNumber { get; set; }

    /// <summary>读取时捕获的 worksheet extLst 原始 XML（含 x14:slicerList 等），写出时透传保留。
    /// 内部使用：r:id 在写出时按关系重编号映射改写 </summary>
    internal string? SheetExtLstXml { get; set; }

    /// <summary>读取时捕获的 workbook.xml 中 &lt;sheet&gt; 元素原始 sheetId 属性值。
    /// 切片器缓存等扩展部件以 tabId 引用工作表，sheetId 必须与原文件一致，否则切片器缓存无法链接到工作表。
    /// 默认空字符串 = 新建工作簿场景，写出时按 1-based 位置序号分配。</summary>
    internal string SheetId { get; set; } = "";

    /// <summary>读取时捕获的 &lt;sheet&gt; 元素 state 属性（hidden / veryHidden）；默认 null = 可见。</summary>
    internal string? SheetState { get; set; }
}