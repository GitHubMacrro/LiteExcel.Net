using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;

namespace LiteExcel;

/// <summary>
/// xlsx 读写器（高层反射 API） 
/// List&lt;T&gt; 映射使用反射，已用 [DynamicallyAccessedMembers] 标注，AOT/裁剪安全 
/// </summary>
public static partial class XlsxReader
{
    /// <summary>
    /// 将指定工作表读为 List&lt;T&gt; 第一行作为表头 
    /// </summary>
    public static List<T> Read<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        string path, int sheetIndex = 0, Action<ReadOptions<T>>? configure = null) where T : new()
    {
        var sheet = Read(path, sheetIndex, firstRowIsHeader: true);
        return SheetToList<T>(sheet, configure);
    }

    /// <summary>
    /// 将指定工作表读为 List&lt;T&gt; 第一行作为表头 
    /// </summary>
    public static List<T> Read<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        string path, string sheetName, Action<ReadOptions<T>>? configure = null) where T : new()
    {
        var sheet = Read(path, sheetName, firstRowIsHeader: true);
        return SheetToList<T>(sheet, configure);
    }

    private static List<T> SheetToList<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        SheetData sheet, Action<ReadOptions<T>>? configure) where T : new()
    {
        var options = new ReadOptions<T>();
        configure?.Invoke(options);

        var propMap = BuildReadPropertyMap<T>(options, sheet.Headers);

        var list = new List<T>(sheet.Rows.Count);
        foreach (var row in sheet.Rows)
        {
            var item = new T();
            foreach (var (prop, colIdx) in propMap)
            {
                if (colIdx < 0 || colIdx >= row.Count) continue;
                var cell = row[colIdx];
                if (cell.Type == CellType.Empty) continue;
                var value = CellToProperty(cell, prop.PropertyType);
                if (value is not null) prop.SetValue(item, value);
            }
            list.Add(item);
        }
        return list;
    }

    private static List<(PropertyInfo prop, int colIdx)> BuildReadPropertyMap<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        ReadOptions<T> options, List<string> headers)
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var result = new List<(PropertyInfo, int)>(props.Length);

        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<LiteColumnAttribute>();
            if (attr is not null && attr.Ignore) continue;

            // 优先用 Fluent 配置的表头名，其次特性 Name，最后属性名
            string headerName;
            var fluentName = options.GetHeaderName(prop.Name);
            if (fluentName is not null) headerName = fluentName;
            else if (attr?.Name is not null) headerName = attr.Name;
            else headerName = prop.Name;

            int colIdx = headers.IndexOf(headerName);
            result.Add((prop, colIdx));
        }
        return result;
    }

    private static object? CellToProperty(Cell cell, Type targetType)
    {
        var value = cell.Type switch
        {
            CellType.Text => cell.Text,
            CellType.Number => ConvertNumber(cell.Number, targetType),
            CellType.Date => cell.Date,
            CellType.Boolean => cell.Boolean,
            _ => null,
        };

        if (value is null) return null;

        var valueType = value.GetType();
        if (targetType.IsAssignableFrom(valueType)) return value;

        // 类型转换
        try
        {
            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static object? ConvertNumber(double num, Type targetType)
    {
        if (targetType == typeof(double) || targetType == typeof(double?)) return num;
        if (targetType == typeof(float) || targetType == typeof(float?)) return (float)num;
        if (targetType == typeof(decimal) || targetType == typeof(decimal?)) return (decimal)num;
        if (targetType == typeof(int) || targetType == typeof(int?)) return (int)num;
        if (targetType == typeof(long) || targetType == typeof(long?)) return (long)num;
        if (targetType == typeof(short) || targetType == typeof(short?)) return (short)num;
        if (targetType == typeof(byte) || targetType == typeof(byte?)) return (byte)num;
        if (targetType == typeof(uint) || targetType == typeof(uint?)) return (uint)num;
        if (targetType == typeof(ulong) || targetType == typeof(ulong?)) return (ulong)num;
        if (targetType == typeof(ushort) || targetType == typeof(ushort?)) return (ushort)num;
        if (targetType == typeof(sbyte) || targetType == typeof(sbyte?)) return (sbyte)num;
        if (targetType == typeof(string)) return num.ToString(CultureInfo.InvariantCulture);
        return num;
    }
}

/// <summary>
/// xlsx 写出器（高层反射 API） 
/// List&lt;T&gt; 映射使用反射，已用 [DynamicallyAccessedMembers] 标注，AOT/裁剪安全 
/// </summary>
public static partial class XlsxWriter
{
    /// <summary>
    /// 将 List&lt;T&gt; 写入 xlsx 文件 
    /// </summary>
    public static void Write<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string path, IEnumerable<T> data, Action<WriteOptions<T>>? configure = null)
    {
        var options = new WriteOptions<T>();
        configure?.Invoke(options);

        var sheet = ListToSheet(data, options);
        Write(path, sheet);
    }

    internal static SheetData ListToSheet<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IEnumerable<T> data, WriteOptions<T> options)
    {
        var sheet = new SheetData
        {
            SheetName = options.SheetName ?? "Sheet1",
            FreezeHeader = options.FreezeHeader,
            ColumnWidths = options.ColumnWidths,
        };

        var columns = BuildWriteColumns<T>(options);

        // 表头
        sheet.Headers = columns.Select(c => c.HeaderName).ToList();

        // 数据行
        foreach (var item in data)
        {
            var cells = new Cell[columns.Count];
            for (int i = 0; i < columns.Count; i++)
            {
                var col = columns[i];
                var value = col.Property.GetValue(item);
                cells[i] = ObjectToCell(value, col.Format, col.IsFormula);
            }
            sheet.Rows.Add(cells);
        }

        return sheet;
    }

    private sealed class WriteColumn
    {
        public PropertyInfo Property = null!;
        public string HeaderName = "";
        public string? Format;
        public bool IsFormula;
    }

    private static List<WriteColumn> BuildWriteColumns<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        WriteOptions<T> options)
    {
        var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        var columns = new List<WriteColumn>(props.Length);

        foreach (var prop in props)
        {
            var attr = prop.GetCustomAttribute<LiteColumnAttribute>();
            if (attr is not null && attr.Ignore) continue;
            if (options.IsIgnored(prop.Name)) continue;

            var headerName = options.GetColumnName(prop.Name) ?? attr?.Name ?? prop.Name;
            var format = options.GetFormat(prop.Name) ?? attr?.Format;
            bool isFormula = options.IsFormulaColumn(prop.Name) || (attr is not null && attr.IsFormula);

            columns.Add(new WriteColumn
            {
                Property = prop,
                HeaderName = headerName,
                Format = format,
                IsFormula = isFormula,
            });
        }

        // 处理 Order：有 Order 特性的按 Order 排，没有的保持声明顺序
        var ordered = new List<WriteColumn>();
        var withOrder = new List<(WriteColumn col, int order)>();

        foreach (var col in columns)
        {
            var attr = col.Property.GetCustomAttribute<LiteColumnAttribute>();
            if (attr is not null && attr.Order >= 0)
                withOrder.Add((col, attr.Order));
            else
                ordered.Add(col);
        }

        var result = new List<WriteColumn>(columns.Count);
        // 先按 Order 排序的有顺序列
        result.AddRange(withOrder.OrderBy(x => x.order).Select(x => x.col));
        // 再追加无顺序列（保持声明顺序）
        result.AddRange(ordered);
        return result;
    }

    private static Cell ObjectToCell(object? value, string? format, bool isFormula = false)
    {
        if (value is null) return Cell.Empty;

        // 公式列：字符串值按公式写出（去掉前导 "="）
        if (isFormula && value is string formulaText)
        {
            var f = formulaText.TrimStart();
            if (f.StartsWith("=", StringComparison.Ordinal)) f = f.Substring(1);
            return Cell.FromFormula(f);
        }

        return value switch
        {
            bool b => Cell.FromBoolean(b),
            DateTime dt => Cell.FromDate(dt, format),
            DateTimeOffset dto => Cell.FromDate(dto.DateTime, format),
            sbyte n => Cell.FromNumber(n, format),
            byte n => Cell.FromNumber(n, format),
            short n => Cell.FromNumber(n, format),
            ushort n => Cell.FromNumber(n, format),
            int n => Cell.FromNumber(n, format),
            uint n => Cell.FromNumber(n, format),
            long n => Cell.FromNumber(n, format),
            ulong n => Cell.FromNumber(n, format),
            float n => Cell.FromNumber(n, format),
            double n => Cell.FromNumber(n, format),
            decimal n => Cell.FromNumber((double)n, format),
            _ => Cell.FromText(value.ToString()),
        };
    }
}
