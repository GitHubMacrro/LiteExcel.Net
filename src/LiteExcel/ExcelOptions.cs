using System.Data;

namespace LiteExcel;

/// <summary>
/// 高层读取选项
/// </summary>
public sealed class ExcelReadOptions
{
    /// <summary>
    /// 是否读取样式。默认 true
    /// </summary>
    public bool ReadStyles { get; set; } = true;

    /// <summary>
    /// 是否把合并单元格左上角的值展开到整个合并区域。默认 false
    /// </summary>
    public bool FillMergedCells { get; set; }

    /// <summary>
    /// 读取完成后是否保持输入流打开（仅 Stream 重载有效）。默认 false
    /// </summary>
    public bool LeaveOpen { get; set; }

    /// <summary>
    /// 打开密码（文件加密）。用于解密带打开密码的 xlsx/xlsm/xlsb。
    /// 未提供时若文件已加密，将抛出明确的加密异常。
    /// </summary>
    public string? OpenPassword { get; set; }

    /// <summary>
    /// 修改密码（写保护）。用于获得编辑/保存权限。
    /// 文件设置了修改密码但未提供（或提供错误）时，工作簿以只读方式打开，不能保存。
    /// </summary>
    public string? ModifyPassword { get; set; }

    /// <summary>
    /// CSV 分隔符（可选，默认 null → 自动探测）。
    /// 仅在 ExcelFormat.Csv 生效。指定后固定使用，不再探测。
    /// 自动探测策略：首位分隔符候选在引号外的出现频率中取最多；三个候选都未出现 → 默认逗号。
    /// </summary>
    public char? Separator { get; set; }

    /// <summary>
    /// CSV 读取编码（可选，默认 null → 按 BOM 探测，无 BOM 回退 UTF-8）。仅在 ExcelFormat.Csv 生效。
    /// 显式指定时优先于 BOM（不再按 BOM 覆盖）。
    /// 编码实例由调用方提供：net48 的 BCL 自带 GBK 等代码页；net8.0 需调用方先注册
    /// <c>Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)</c>（本库不引用任何编码包）。
    /// </summary>
    public System.Text.Encoding? Encoding { get; set; }
}

/// <summary>
/// 高层写入选项
/// </summary>
public sealed class ExcelWriteOptions
{
    /// <summary>
    /// 目标文件已存在时是否覆盖。默认 true
    /// </summary>
    public bool Overwrite { get; set; } = true;

    /// <summary>
    /// 写出前自动估算列宽。默认 false
    /// </summary>
    public bool AutoFitColumns { get; set; }

    /// <summary>
    /// 写出时冻结表头。默认 false
    /// </summary>
    public bool FreezeHeader { get; set; }

    /// <summary>
    /// 覆盖工作簿文档属性（可选）
    /// </summary>
    public WorkbookProperties? Properties { get; set; }

    /// <summary>
    /// 写入完成后是否保持输出流打开（仅 Stream 重载有效）。默认 false
    /// </summary>
    public bool LeaveOpen { get; set; }

    /// <summary>
    /// 能力降级回调（可选，默认 null）。
    /// 写出到不支持某能力的格式（如 xls/xlsb/csv）时，对被静默丢弃的能力逐项回调。
    /// 默认关闭：不注册则行为与历史版本完全一致（无破坏性）。
    /// </summary>
    public Action<DegradationInfo>? OnDegradation { get; set; }

    /// <summary>
    /// CSV 写出时分隔符（可选，默认 null → 逗号）。仅在 ExcelFormat.Csv 生效。
    /// </summary>
    public char? Separator { get; set; }

    /// <summary>
    /// CSV 写出编码（可选，默认 null → UTF-8 带 BOM）。仅在 ExcelFormat.Csv 生效。
    /// BOM 由所给编码的 preamble 决定：<c>Encoding.UTF8</c> / <c>new UTF8Encoding(true)</c> 写 BOM；
    /// <c>new UTF8Encoding(false)</c> 与 GBK 等无 preamble 编码不写。
    /// 编码实例由调用方提供：net48 的 BCL 自带 GBK 等代码页；net8.0 需调用方先注册
    /// <c>Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)</c>（本库不引用任何编码包）。
    /// </summary>
    public System.Text.Encoding? Encoding { get; set; }
}
