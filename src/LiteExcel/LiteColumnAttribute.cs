namespace LiteExcel;

/// <summary>
/// 标在属性上，控制 List&lt;T&gt; 映射时的列名/顺序/格式/忽略 
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class LiteColumnAttribute : Attribute
{
    /// <summary>列名，默认用属性名 </summary>
    public string? Name { get; set; }

    /// <summary>列顺序，-1 按声明顺序 </summary>
    public int Order { get; set; } = -1;

    /// <summary>数字/日期格式，如 "0.00" / "yyyy-MM-dd" </summary>
    public string? Format { get; set; }

    /// <summary>true 则把该字符串属性当作公式写出（值可带或不带前导 "="） </summary>
    public bool IsFormula { get; set; }

    /// <summary>true 则不输出该列 </summary>
    public bool Ignore { get; set; }
}
