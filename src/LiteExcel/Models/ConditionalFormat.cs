namespace LiteExcel;

/// <summary>
/// 条件格式规则类型（2.4.3 仅支持 cellIs / expression / colorScale / dataBar）。
/// </summary>
public enum ConditionalFormatType
{
    /// <summary>单元格值比较（= </> < /= /= < / >= / between / notBetween 等） </summary>
    CellIs,

    /// <summary>公式条件（formula 引用的表达式为 TRUE 时生效） </summary>
    Expression,

    /// <summary>色阶（从低值到高值平滑过渡） </summary>
    ColorScale,

    /// <summary>数据条（单元格内条形图） </summary>
    DataBar,
}

/// <summary>
/// cellIs 条件比较操作。
/// </summary>
public enum ConditionalOperator
{
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    Between,
    NotBetween,
}

/// <summary>
/// 色阶（colorScale）参数。一般采用 2 色（低/高）；Mid 可选为 3 色。
/// </summary>
public sealed class ColorScaleInfo
{
    /// <summary>低值颜色（#RRGGBB） </summary>
    public string LowColor { get; set; } = "F8696B";

    /// <summary>高值颜色（#RRGGBB） </summary>
    public string HighColor { get; set; } = "63BE7B";

    /// <summary>中间颜色（可选，设 nonNull 时启用 3 色刻度） </summary>
    public string? MidColor { get; set; }
}

/// <summary>
/// 数据条（dataBar）参数。
/// </summary>
public sealed class DataBarInfo
{
    /// <summary>条形颜色（#RRGGBB，默认 Excel 蓝 #638EC6） </summary>
    public string Color { get; set; } = "638EC6";

    /// <summary>是否同时显示文本值（默认 true）。false 时只显示条形 </summary>
    public bool ShowValue { get; set; } = true;

    /// <summary>最小长度（0-100，默认 0） </summary>
    public int MinLengthPercent { get; set; }

    /// <summary>最大长度（0-100，默认 100） </summary>
    public int MaxLengthPercent { get; set; } = 100;
}

/// <summary>
/// 一张工作表中某范围的条料格式规则。
/// </summary>
public sealed class ConditionalFormat
{
    /// <summary>适用 范围（如 "B2:B10"、"A1:A100 D2:D9"） </summary>
    public string Sqref { get; set; } = "";

    /// <summary>规则类型 </summary>
    public ConditionalFormatType Type { get; set; }

    /// <summary>
    /// 条件主体。
    /// CellIs：比较目标（数字/文本，参与 Operator）
    /// Expression：公式字符串（相对引用）
    /// Between：Formula1 / Formula2 由 Formula / Formula2 分别给出
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>CellIs Between 专用上限（between / notBetween 时） </summary>
    public string? Formula2 { get; set; }

    /// <summary>cellIs 比较操作。仅 CellIs 有效。 </summary>
    public ConditionalOperator Operator { get; set; } = ConditionalOperator.GreaterThan;

    /// <summary>条件满足时的单元格样式（字体/填充/边框；不包含对齐与数字格式） </summary>
    public CellStyle? Style { get; set; }

    /// <summary>ColorScale 专用参数 </summary>
    public ColorScaleInfo? ColorScale { get; set; }

    /// <summary>DataBar 专用参数 </summary>
    public DataBarInfo? DataBar { get; set; }

    /// <summary>优先级（各工作区域 cn 唯一；默认按注册顺序自动编号） </summary>
    public int Priority { get; set; }
}
