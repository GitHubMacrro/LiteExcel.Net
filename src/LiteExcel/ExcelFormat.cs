namespace LiteExcel;

/// <summary>
/// 支持的 Excel 文件格式
/// </summary>
public enum ExcelFormat
{
    /// <summary>标准工作簿（OOXML） </summary>
    Xlsx,

    /// <summary>启用宏的工作簿（OOXML + 宏） </summary>
    Xlsm,

    /// <summary>二进制工作簿（OOXML 二进制格式，预留） </summary>
    Xlsb,

    /// <summary>旧版二进制工作簿（BIFF，预留，先只读） </summary>
    Xls,

    /// <summary>逗号分隔文本（轻量格式，仅表格数据） </summary>
    Csv,
}
