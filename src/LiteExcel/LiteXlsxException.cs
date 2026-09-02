namespace LiteExcel;

/// <summary>
/// LiteExcel 库所有异常的基类。
/// </summary>
public class LiteExcelException : Exception
{
    public LiteExcelException() { }

    public LiteExcelException(string message) : base(message) { }

    public LiteExcelException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// 旧异常名称的兼容别名。请改用 <see cref="LiteExcelException"/>。
/// </summary>
[Obsolete("Use LiteExcelException instead.")]
public class LiteXlsxException : LiteExcelException
{
    public LiteXlsxException() { }

    public LiteXlsxException(string message) : base(message) { }

    public LiteXlsxException(string message, Exception innerException) : base(message, innerException) { }
}

/// <summary>
/// 当 Sheet 名不合法（为空、超过 31 字符、包含非法字符）时抛出。
/// </summary>
public class InvalidSheetNameException : LiteExcelException
{
    public InvalidSheetNameException() { }

    public InvalidSheetNameException(string message) : base(message) { }

    public InvalidSheetNameException(string message, Exception innerException) : base(message, innerException) { }

    public InvalidSheetNameException(string sheetName, string message) : base(message)
    {
        SheetName = sheetName;
    }

    /// <summary>触发异常的非法 Sheet 名。</summary>
    public string? SheetName { get; }
}

/// <summary>
/// 流式写入时超出 Excel 单表行数上限（1,048,576）且模式为 <see cref="RowLimitExceededMode.Throw"/> 时抛出。
/// </summary>
public class RowLimitExceededException : LiteExcelException
{
    /// <summary>触发上限的行号（1 基）。</summary>
    public int RowNumber { get; }

    /// <summary>单表行数上限。</summary>
    public int MaxRows { get; }

    public RowLimitExceededException(int rowNumber, int maxRows)
        : base($"已达到 Excel 单表行数上限 {maxRows:N0}（尝试写入第 {rowNumber:N0} 行）。" +
               (rowNumber > maxRows ? "请改用 RowLimitExceededMode.SpillToNewSheet 自动分表，或自行拆分数据。" : ""))
    {
        RowNumber = rowNumber;
        MaxRows = maxRows;
    }
}

/// <summary>
/// 流式写入器在达到单表行数上限时的行为。
/// </summary>
public enum RowLimitExceededMode
{
    /// <summary>抛出 <see cref="RowLimitExceededException"/>（默认）。</summary>
    Throw,

    /// <summary>自动新建工作表（Sheet1/Sheet2/...）继续写入。</summary>
    SpillToNewSheet,

    /// <summary>停止写入后续行（截断），文件以满行为止；可通过 <see cref="XlsxStreamWriter.Truncated"/> 感知是否发生了截断。</summary>
    Truncate,
}
