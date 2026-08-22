namespace LiteExcel;

/// <summary>
/// 命名区域（definedNames），由工作簿持有。
/// 文件内：workbook.xml 的 &lt;definedNames&gt; 中每个 &lt;definedName&gt; 一个。
/// 名称 + 引用文本（A1 引用，含可选 sheet 限定）。
/// </summary>
public sealed class NamedRange
{
    /// <summary>名称（必须合法：英文字母或下划线开头，支持英文/数字/._） </summary>
    public string Name { get; set; } = "";

    /// <summary>引用文本（如 Sheet1!$A$1:$C$9、Sheet1!$B$2、#VALUE! 或公式） </summary>
    public string Reference { get; set; } = "";

    /// <summary>sheet-local 索引（localSheetId），-1 表示全局工作簿名称 </summary>
    public int LocalSheetId { get; set; } = -1;

    /// <summary>是否为本工作表局部名称（1-based sheet 索引） </summary>
    public bool IsLocalSheet => LocalSheetId >= 0;
}
