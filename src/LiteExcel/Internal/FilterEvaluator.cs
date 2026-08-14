using System.Globalization;

namespace LiteExcel.Internal;

/// <summary>
/// Evaluates filter conditions against row data to compute hidden rows.
/// </summary>
internal static class FilterEvaluator
{
    /// <summary>
    /// Evaluate filter conditions and return the set of row indices that should be hidden.
    /// A row is hidden if it does NOT match ALL column conditions.
    /// </summary>
    public static HashSet<int> EvaluateHiddenRows(SheetData sheet)
    {
        var hidden = new HashSet<int>();

        if (sheet.Filter is null || sheet.Filter.Columns.Count == 0)
            return hidden;

        var filter = sheet.Filter;

        for (int rowIdx = 0; rowIdx < sheet.Rows.Count; rowIdx++)
        {
            var row = sheet.Rows[rowIdx];
            bool visible = true;

            foreach (var col in filter.Columns)
            {
                if (col.ColumnIndex < 0 || col.ColumnIndex >= row.Count)
                {
                    // Column doesn't exist in this row; treat as blank
                    if (!MatchesCondition(col, ""))
                    {
                        visible = false;
                        break;
                    }
                    continue;
                }

                var cell = row[col.ColumnIndex];
                string value = CellToString(cell);

                if (!MatchesCondition(col, value))
                {
                    visible = false;
                    break;
                }
            }

            if (!visible)
                hidden.Add(rowIdx);
        }

        return hidden;
    }

    private static string CellToString(Cell cell)
    {
        return cell.Type switch
        {
            CellType.Text => cell.Text ?? "",
            CellType.Number => cell.Number.ToString(CultureInfo.InvariantCulture),
            CellType.Date => cell.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            CellType.Boolean => cell.Boolean ? "TRUE" : "FALSE",
            _ => "",
        };
    }

    private static bool MatchesCondition(FilterColumn col, string value)
    {
        return col.Type switch
        {
            FilterType.Equals => MatchesEquals(col.Values, value),
            FilterType.Compare => MatchesCompare(col, value),
            FilterType.Contains => MatchesContains(col.Values, value),
            FilterType.BeginsWith => MatchesBeginsWith(col.Values, value),
            FilterType.EndsWith => MatchesEndsWith(col.Values, value),
            FilterType.Blank => MatchesBlank(col, value),
            _ => true,
        };
    }

    private static bool MatchesEquals(List<string> values, string value)
    {
        if (values.Count == 0) return true;
        foreach (var v in values)
        {
            if (string.Equals(v, value, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool MatchesCompare(FilterColumn col, string value)
    {
        // Between uses MinValue/MaxValue, not Values
        if (col.Operator == FilterOperator.Between)
        {
            if (col.MinValue is null || col.MaxValue is null) return true;
            double? betweenValue = TryParseDouble(value);
            if (betweenValue is null) return false;
            var min = TryParseDouble(col.MinValue);
            var max = TryParseDouble(col.MaxValue);
            if (min is null || max is null) return false;
            return betweenValue >= min && betweenValue <= max;
        }

        if (col.Values.Count == 0) return true;

        double? numValue = TryParseDouble(value);
        if (numValue is null) return false;

        switch (col.Operator)
        {
            case FilterOperator.GreaterThan:
                foreach (var v in col.Values)
                {
                    var cmp = TryParseDouble(v);
                    if (cmp is not null && numValue > cmp) return true;
                }
                return false;

            case FilterOperator.GreaterThanOrEqual:
                foreach (var v in col.Values)
                {
                    var cmp = TryParseDouble(v);
                    if (cmp is not null && numValue >= cmp) return true;
                }
                return false;

            case FilterOperator.LessThan:
                foreach (var v in col.Values)
                {
                    var cmp = TryParseDouble(v);
                    if (cmp is not null && numValue < cmp) return true;
                }
                return false;

            case FilterOperator.LessThanOrEqual:
                foreach (var v in col.Values)
                {
                    var cmp = TryParseDouble(v);
                    if (cmp is not null && numValue <= cmp) return true;
                }
                return false;

            default:
                return true;
        }
    }

    private static bool MatchesContains(List<string> values, string value)
    {
        if (values.Count == 0) return true;
        foreach (var v in values)
        {
            if (value.IndexOf(v, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        return false;
    }

    private static bool MatchesBeginsWith(List<string> values, string value)
    {
        if (values.Count == 0) return true;
        foreach (var v in values)
        {
            if (value.StartsWith(v, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool MatchesEndsWith(List<string> values, string value)
    {
        if (values.Count == 0) return true;
        foreach (var v in values)
        {
            if (value.EndsWith(v, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool MatchesBlank(FilterColumn col, string value)
    {
        // Values.Count == 0 means "is blank"; Values.Count > 0 means "is not blank"
        if (col.Values.Count == 0)
            return string.IsNullOrEmpty(value);
        return !string.IsNullOrEmpty(value);
    }

    private static double? TryParseDouble(string s)
    {
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            return d;
        return null;
    }
}
