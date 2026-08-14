namespace LiteExcel;

/// <summary>
/// Auto filter condition types.
/// </summary>
public enum FilterType
{
    Equals,
    Compare,
    Contains,
    BeginsWith,
    EndsWith,
    Blank,
}

/// <summary>
/// Comparison operators for <see cref="FilterType.Compare"/>.
/// </summary>
public enum FilterOperator
{
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Between,
}

/// <summary>
/// Auto filter definition for a worksheet.
/// Range: e.g. "A1:D542".
/// Columns: filter conditions per column.
/// HiddenRows: 0-based row indices within Rows that should be hidden.
/// </summary>
public sealed class AutoFilter
{
    public string Range { get; set; } = "";
    public List<FilterColumn> Columns { get; set; } = new();
    public HashSet<int> HiddenRows { get; set; } = new();
}

/// <summary>
/// Filter condition for a single column.
/// </summary>
public sealed class FilterColumn
{
    public int ColumnIndex { get; set; }
    public FilterType Type { get; set; }
    public List<string> Values { get; set; } = new();
    public FilterOperator Operator { get; set; }

    /// <summary>Lower bound for Between operator.</summary>
    public string? MinValue { get; set; }

    /// <summary>Upper bound for Between operator.</summary>
    public string? MaxValue { get; set; }
}
