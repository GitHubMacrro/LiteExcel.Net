using System.Linq.Expressions;

namespace LiteExcel;

/// <summary>
/// List&lt;T&gt; 写出配置 支持 Fluent 回调和字典映射 
/// </summary>
public sealed class WriteOptions<T>
{
    /// <summary>工作表名，默认 "Sheet1" </summary>
    public string? SheetName { get; set; }

    /// <summary>是否冻结表头 </summary>
    public bool FreezeHeader { get; set; }

    /// <summary>列宽列表（可选） </summary>
    public List<double>? ColumnWidths { get; set; }

    private readonly Dictionary<string, string> _columnNames = new();
    private readonly Dictionary<string, string> _columnFormats = new();
    private readonly HashSet<string> _ignored = new();

    /// <summary>指定属性对应的列名和可选格式（Fluent，引用类型） </summary>
    public WriteOptions<T> Column(Expression<Func<T, object?>> prop, string name, string? format = null)
    {
        var propName = ExpressionHelper.GetPropertyName(prop);
        _columnNames[propName] = name;
        if (format is not null) _columnFormats[propName] = format;
        return this;
    }

    /// <summary>指定属性对应的列名和可选格式（Fluent，值类型） </summary>
    public WriteOptions<T> Column<TProp>(Expression<Func<T, TProp>> prop, string name, string? format = null)
    {
        var propName = ExpressionHelper.GetPropertyName(prop);
        _columnNames[propName] = name;
        if (format is not null) _columnFormats[propName] = format;
        return this;
    }

    /// <summary>忽略指定属性（Fluent，引用类型） </summary>
    public WriteOptions<T> Ignore(Expression<Func<T, object?>> prop)
    {
        _ignored.Add(ExpressionHelper.GetPropertyName(prop));
        return this;
    }

    /// <summary>忽略指定属性（Fluent，值类型） </summary>
    public WriteOptions<T> Ignore<TProp>(Expression<Func<T, TProp>> prop)
    {
        _ignored.Add(ExpressionHelper.GetPropertyName(prop));
        return this;
    }

    /// <summary>批量映射属性名 -> 列名（字典映射，老项目常见） </summary>
    public WriteOptions<T> Map(IDictionary<string, string> mapping)
    {
        foreach (var kv in mapping) _columnNames[kv.Key] = kv.Value;
        return this;
    }

    internal string? GetColumnName(string propName) =>
        _columnNames.TryGetValue(propName, out var name) ? name : null;

    internal string? GetFormat(string propName) =>
        _columnFormats.TryGetValue(propName, out var fmt) ? fmt : null;

    internal bool IsIgnored(string propName) => _ignored.Contains(propName);
}

/// <summary>
/// List&lt;T&gt; 读取配置 支持 Fluent 回调和字典映射 
/// </summary>
public sealed class ReadOptions<T>
{
    private readonly Dictionary<string, string> _headerNames = new();

    /// <summary>指定属性对应的表头名（Fluent，引用类型） </summary>
    public ReadOptions<T> Column(Expression<Func<T, object?>> prop, string headerName)
    {
        _headerNames[ExpressionHelper.GetPropertyName(prop)] = headerName;
        return this;
    }

    /// <summary>指定属性对应的表头名（Fluent，值类型） </summary>
    public ReadOptions<T> Column<TProp>(Expression<Func<T, TProp>> prop, string headerName)
    {
        _headerNames[ExpressionHelper.GetPropertyName(prop)] = headerName;
        return this;
    }

    /// <summary>批量映射属性名 -> 表头名（字典映射） </summary>
    public ReadOptions<T> Map(IDictionary<string, string> mapping)
    {
        foreach (var kv in mapping) _headerNames[kv.Key] = kv.Value;
        return this;
    }

    internal string? GetHeaderName(string propName) =>
        _headerNames.TryGetValue(propName, out var name) ? name : null;
}

/// <summary>表达式辅助工具 </summary>
internal static class ExpressionHelper
{
    /// <summary>从 x => x.Prop 或 x => (object)x.Prop 中提取属性名 </summary>
    public static string GetPropertyName<T>(Expression<Func<T, object?>> expr)
    {
        var body = expr.Body;
        // 剥开值类型到 object 的 UnaryExpression(Convert)
        if (body is UnaryExpression unary) body = unary.Operand;
        if (body is MemberExpression member) return member.Member.Name;
        throw new ArgumentException("表达式必须是属性访问，如 x => x.Name", nameof(expr));
    }

    /// <summary>从 x => x.Prop 中提取属性名（泛型版本） </summary>
    public static string GetPropertyName<T, TProp>(Expression<Func<T, TProp>> expr)
    {
        var body = expr.Body;
        if (body is UnaryExpression unary) body = unary.Operand;
        if (body is MemberExpression member) return member.Member.Name;
        throw new ArgumentException("表达式必须是属性访问，如 x => x.Name", nameof(expr));
    }
}
