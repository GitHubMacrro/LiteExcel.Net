using System.Globalization;

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
    public CellStyle? Style
    {
        get => _style;
        set
        {
            _style = value;
            Owner?.OnCellChanged(this);
        }
    }
    private CellStyle? _style;

    public string? NumberFormat
    {
        get => _numberFormat;
        set
        {
            _numberFormat = value;
            Owner?.OnCellChanged(this);
        }
    }
    private string? _numberFormat;

    /// <summary>
    /// 是否为公式字符串（写出时按公式处理）
    /// </summary>
    public bool IsFormula { get; set; }

    /// <summary>单元格超链接（可选） </summary>
    public Hyperlink? Hyperlink
    {
        get => _hyperlink;
        set
        {
            _hyperlink = value;
            Owner?.OnCellChanged(this);
        }
    }
    private Hyperlink? _hyperlink;

    /// <summary>所属工作表（由高层 Worksheet 挂接，用于写回） </summary>
    internal Worksheet? Owner { get; set; }

    /// <summary>Owner 坐标系下 1-based 行号 </summary>
    internal int OwnerRow { get; set; }

    /// <summary>Owner 坐标系下 1-based 列号 </summary>
    internal int OwnerCol { get; set; }

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

    /// <summary>创建公式单元格（仅写公式字符串，不计算结果） </summary>
    public static Cell FromFormula(string formula)
    {
        return new Cell { Type = CellType.Text, Text = formula, IsFormula = true };
    }

    public static Cell Empty => new() { Type = CellType.Empty };

    /// <summary>设置单元格值。空单元格的 null/空串转为 Empty </summary>
    public void SetValue(object? value)
    {
        if (value is Cell cell)
        {
            CopyFrom(cell);
        }
        else if (value is null || value == DBNull.Value)
        {
            Type = CellType.Empty;
            Text = null;
            Number = 0;
            Date = default;
            Boolean = false;
            IsFormula = false;
        }
        else
        {
            switch (value)
            {
                case bool b:
                    Type = CellType.Boolean; Boolean = b; Text = null; Number = 0; Date = default; IsFormula = false;
                    break;
                case DateTime dt:
                    Type = CellType.Date; Date = dt; NumberFormat ??= "yyyy-MM-dd"; Text = null; Number = 0; Boolean = false; IsFormula = false;
                    break;
                case sbyte n: SetNumber(n); break;
                case byte n: SetNumber(n); break;
                case short n: SetNumber(n); break;
                case ushort n: SetNumber(n); break;
                case int n: SetNumber(n); break;
                case uint n: SetNumber(n); break;
                case long n: SetNumber(n); break;
                case ulong n: SetNumber(n); break;
                case float n: SetNumber(n); break;
                case double n: SetNumber(n); break;
                case decimal n: SetNumber((double)n); break;
                default:
                    Type = CellType.Text; Text = value.ToString(); Number = 0; Date = default; Boolean = false; IsFormula = false;
                    break;
            }
        }

        Owner?.OnCellChanged(this);
    }

    private void SetNumber(double n)
    {
        Type = CellType.Number;
        Number = n;
        Text = null;
        Date = default;
        Boolean = false;
        IsFormula = false;
    }

    private void CopyFrom(Cell other)
    {
        Type = other.Type;
        Text = other.Text;
        Number = other.Number;
        Date = other.Date;
        Boolean = other.Boolean;
        Style = other.Style;
        NumberFormat = other.NumberFormat;
        IsFormula = other.IsFormula;
        Hyperlink = other.Hyperlink?.Clone();
    }
    /// <summary>以字符串读取值。Empty 返回 null，Number/Date/Boolean 按惯例格式化 </summary>
    public string? GetString()
    {
        return Type switch
        {
            CellType.Text => Text,
            CellType.Number => Number.ToString(CultureInfo.InvariantCulture),
            CellType.Date => Date.ToString(NumberFormat ?? "yyyy-MM-dd", CultureInfo.InvariantCulture),
            CellType.Boolean => Boolean ? "TRUE" : "FALSE",
            _ => null,
        };
    }

    /// <summary>以 double 读取值。类型不匹配抛 <see cref="InvalidCastException"/> </summary>
    public double GetDouble()
    {
        if (Type != CellType.Number) throw new InvalidCastException($"单元格类型为 {Type}，不能读取为 double。");
        return Number;
    }

    /// <summary>以 DateTime 读取值。类型不匹配抛 <see cref="InvalidCastException"/> </summary>
    public DateTime GetDateTime()
    {
        if (Type != CellType.Date) throw new InvalidCastException($"单元格类型为 {Type}，不能读取为 DateTime。");
        return Date;
    }

    /// <summary>以 bool 读取值。类型不匹配抛 <see cref="InvalidCastException"/> </summary>
    public bool GetBoolean()
    {
        if (Type != CellType.Boolean) throw new InvalidCastException($"单元格类型为 {Type}，不能读取为 bool。");
        return Boolean;
    }

    /// <summary>尝试以 string 读取值。成功返回 true，空单元格返回 false 且 value 为 null </summary>
    public bool TryGetString(out string? value)
    {
        if (IsEmpty) { value = null; return false; }
        value = GetString();
        return true;
    }

    /// <summary>尝试以 double 读取值 </summary>
    public bool TryGetDouble(out double value)
    {
        if (Type == CellType.Number) { value = Number; return true; }
        value = 0;
        return false;
    }

    /// <summary>尝试以 DateTime 读取值 </summary>
    public bool TryGetDateTime(out DateTime value)
    {
        if (Type == CellType.Date) { value = Date; return true; }
        value = default;
        return false;
    }

    /// <summary>尝试以 bool 读取值 </summary>
    public bool TryGetBoolean(out bool value)
    {
        if (Type == CellType.Boolean) { value = Boolean; return true; }
        value = false;
        return false;
    }

    /// <summary>读取原始值对象（object），Empty 返回 null </summary>
    public object? GetValue()
    {
        return Type switch
        {
            CellType.Text => Text,
            CellType.Number => Number,
            CellType.Date => Date,
            CellType.Boolean => Boolean,
            _ => null,
        };
    }
}
