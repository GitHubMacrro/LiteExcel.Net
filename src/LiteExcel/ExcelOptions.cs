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
