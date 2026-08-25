using System;

namespace LiteExcel;

/// <summary>
/// Excel 内置表格样式名枚举。样式长什么样由 Excel 程序内置渲染，文件只保存样式名字符串。
/// 共 60 种：Light 1-21 / Medium 1-28 / Dark 1-11。
/// </summary>
public enum TableStyleStyle
{
    /// <summary>无样式（无色无条纹） </summary>
    None = 0,

    Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10,
    Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20,
    Light21,

    Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10,
    Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20,
    Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28,

    Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11,
}

/// <summary>
/// 超级表（Excel Table / ListObject）的单一列定义。
/// </summary>
public sealed class XlTableColumn
{
    /// <summary>列名（= 表头单元格文本） </summary>
    public string Name { get; set; } = "";

    /// <summary>列格式（font/fill/border），写出时映射到 dxf（dataDxfId） </summary>
    public CellStyle? Style { get; set; }

    /// <summary>列数字格式（如 "yyyy/m/d" / "#,##0.00"），写出时并入列 dxf </summary>
    public string? NumberFormat { get; set; }

    /// <summary>clone（供模型复制） </summary>
    public XlTableColumn Clone() => new() { Name = Name, Style = Style?.Clone(), NumberFormat = NumberFormat };
}

/// <summary>
/// 超级表（Excel 内置表 / ListObject）：带条纹样式、表头筛选、结构范围的区域。
/// 样式由 Excel 内置渲染，本对象仅保存样式名。
/// </summary>
public sealed class XlTable
{
    /// <summary>表名（全簿唯一；允许中文；不能以数字开头、不能含空格、不能撞单元格地址） </summary>
    public string Name { get; set; } = "";

    /// <summary>覆盖区域（A1 风格，首行恒为表头） </summary>
    public string Ref { get; set; } = "";

    /// <summary>内置样式（默认 Medium9，即 Excel 插入表的默认观感）。与 <see cref="CustomStyleName"/> 互斥，CustomStyleName 优先生效 </summary>
    public TableStyleStyle Style { get; set; } = TableStyleStyle.Medium9;

    /// <summary>自定义样式名字符串（任意内置/保留名）。非空时优先生效；不在 60 个内置名内时 Excel 静默退化为无样式（经 OnDegradation 上报） </summary>
    public string? CustomStyleName { get; set; }

    /// <summary>斑马线（行条纹），默认 true </summary>
    public bool ShowRowStripes { get; set; } = true;

    /// <summary>首列强调 </summary>
    public bool ShowFirstColumn { get; set; }

    /// <summary>末列强调 </summary>
    public bool ShowLastColumn { get; set; }

    /// <summary>列条纹 </summary>
    public bool ShowColumnStripes { get; set; }

    /// <summary>表头筛选下拉，默认 true </summary>
    public bool AutoFilter { get; set; } = true;

    /// <summary>是否显示汇总行（Phase 1 恒 false；仅为读回保留） </summary>
    public bool TotalsRowShown { get; set; }

    /// <summary>可选表头样式 → headerRowDxfId </summary>
    public CellStyle? HeaderStyle { get; set; }

    private readonly List<XlTableColumn> _columns = new();

    /// <summary>列列表（由 AddTable 从表头行生成） </summary>
    public IReadOnlyList<XlTableColumn> Columns => _columns;

    /// <summary>按列名取列（大小写不敏感）。不存在抛 <see cref="LiteExcelException"/> </summary>
    public XlTableColumn Column(string name)
    {
        foreach (var c in _columns)
            if (string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
                return c;
        throw new LiteExcelException($"表中不存在列：{name}");
    }

    internal void AddColumn(XlTableColumn column) => _columns.Add(column);

    internal void ClearColumns() => _columns.Clear();

    /// <summary>写出到 tableStyleInfo 的样式名 </summary>
    internal string StyleName =>
        !string.IsNullOrEmpty(CustomStyleName) ? CustomStyleName!
        : Style == TableStyleStyle.None ? "None"
        : "TableStyle" + Style;
}
