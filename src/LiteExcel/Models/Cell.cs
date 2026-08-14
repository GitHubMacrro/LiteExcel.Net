namespace LiteExcel;

/// <summary>
/// 单元格值类型
/// </summary>
public enum CellType
{
    Text,
    Number,
    Date,
    Boolean,
    Empty,
}

/// <summary>
/// 表示一个 xlsx 单元格,<see cref="Type"/> 决定哪个值字段有效
/// </summary>
public sealed class Cell
{
    public CellType Type { get; set; }
    public string? Text { get; set; }
    public double Number { get; set; }
    public DateTime Date { get; set; }
    public bool Boolean { get; set; }
    public CellStyle? Style { get; set; }
    public string? NumberFormat { get; set; }

    public bool IsEmpty => Type == CellType.Empty;

    public static Cell FromText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return new Cell { Type = CellType.Empty };
        return new Cell { Type = CellType.Text, Text = text };
    }

    public static Cell FromNumber(double number, string? numberFormat = null)
    {
        return new Cell { Type = CellType.Number, Number = number, NumberFormat = numberFormat };
    }

    public static Cell FromDate(DateTime date, string? numberFormat = null)
    {
        return new Cell
        {
            Type = CellType.Date,
            Date = date,
            NumberFormat = numberFormat ?? "yyyy-MM-dd",
        };
    }

    public static Cell FromBoolean(bool value)
    {
        return new Cell { Type = CellType.Boolean, Boolean = value };
    }

    public static Cell Empty => new() { Type = CellType.Empty };
}
