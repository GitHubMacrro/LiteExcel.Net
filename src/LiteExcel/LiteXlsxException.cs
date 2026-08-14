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
