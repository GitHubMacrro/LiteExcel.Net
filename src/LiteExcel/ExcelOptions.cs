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
}
