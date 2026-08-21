namespace LiteExcel;

/// <summary>
/// 可降级的 Excel 能力（写出到不支持该能力的格式时上报）。
/// </summary>
public enum DegradationCapability
{
    /// <summary>单元格批注 </summary>
    Comments,

    /// <summary>数据验证（下拉列表等） </summary>
    DataValidation,

    /// <summary>自动筛选 </summary>
    AutoFilter,

    /// <summary>图片（浮动 / InCell） </summary>
    Images,

    /// <summary>文档属性（作者/标题等 docProps） </summary>
    DocumentProperties,

    /// <summary>命名区域（definedNames） </summary>
    NamedRanges,

    /// <summary>单元格样式（字体/颜色/边框/对齐/换行/数字格式） </summary>
    Styles,

    /// <summary>合并单元格 </summary>
    MergedCells,

    /// <summary>冻结窗格（冻结行列） </summary>
    FreezePanes,

    /// <summary>超链接 </summary>
    Hyperlinks,

    /// <summary>行高 </summary>
    RowHeights,

    /// <summary>列宽 </summary>
    ColumnWidths,

    /// <summary>公式 </summary>
    Formulas,

    /// <summary>图表 </summary>
    Charts,

    /// <summary>透视表 </summary>
    PivotTables,

    /// <summary>InCell 图片（richData） </summary>
    RichData,
}

/// <summary>
/// 一次能力降级事件：写出到目标格式时，某项 Excel 能力被静默丢弃的说明。
/// 通过 <see cref="ExcelWriteOptions.OnDegradation"/> 回调上报，默认关闭（无破坏性）。
/// </summary>
public sealed class DegradationInfo
{
    /// <summary>被丢弃的能力 </summary>
    public DegradationCapability Capability { get; set; }

    /// <summary>受影响的工作表名（工作簿级能力为 null） </summary>
    public string? SheetName { get; set; }

    /// <summary>写出目标格式 </summary>
    public ExcelFormat TargetFormat { get; set; }

    /// <summary>人类可读说明 </summary>
    public string Message { get; set; } = "";
}

internal static class DegradationDetector
{
    /// <summary>工作表是否存在除 NumberFormat 之外的完整样式（字体/颜色/边框/对齐/换行） </summary>
    public static bool HasNonNumberFormatStyles(SheetData sheet)
    {
        if (HasStyleProps(sheet.HeaderStyle) || HasStyleProps(sheet.DefaultStyle)) return true;
        if (sheet.RowStyles is { Count: > 0 })
        {
            foreach (var kv in sheet.RowStyles)
                if (HasStyleProps(kv.Value)) return true;
        }
        if (sheet.ColumnStyles is { Count: > 0 })
        {
            foreach (var kv in sheet.ColumnStyles)
                if (HasStyleProps(kv.Value)) return true;
        }
        foreach (var row in sheet.Rows)
        {
            foreach (var cell in row)
                if (HasStyleProps(cell.Style)) return true;
        }
        return false;
    }

    private static bool HasStyleProps(CellStyle? s)
    {
        if (s is null) return false;
        return s.Bold || s.Italic || s.Underline || s.Strikeout
            || s.FontColor is not null || s.FillColor is not null
            || s.FontName is not null
            || (s.FontSize > 0 && s.FontSize != 11)
            || s.HorizontalAlignment != HorizontalAlignment.General
            || s.VerticalAlignment != VerticalAlignment.Bottom
            || s.WrapText
            || s.Border is not null;
    }
}
