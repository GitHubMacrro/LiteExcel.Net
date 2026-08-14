namespace LiteExcel;

/// <summary>
/// 数据验证（Data Validation）类型 
/// </summary>
public enum DataValidationType
{
    /// <summary>下拉列表验证 </summary>
    List,

    /// <summary>整数验证 </summary>
    WholeNumber,

    /// <summary>小数验证 </summary>
    Decimal,

    /// <summary>日期验证 </summary>
    Date,
}

/// <summary>
/// 单元格数据验证规则 
/// </summary>
public sealed class DataValidation
{
    /// <summary>验证类型 </summary>
    public DataValidationType Type { get; set; }

    /// <summary>应用范围，如 A1:A10 </summary>
    public string Sqref { get; set; } = "";

    /// <summary>
    /// 验证公式 1 
    /// List 类型示例：用引号包裹的逗号分隔列表；
    /// 数值/日期区间验证时为下限公式 
    /// </summary>
    public string Formula1 { get; set; } = "";

    /// <summary>
    /// 验证公式 2，用于区间验证（Between）的上限公式 非区间验证时可为 null 
    /// </summary>
    public string? Formula2 { get; set; }

    /// <summary>是否允许空白 </summary>
    public bool AllowBlank { get; set; }

    /// <summary>输入提示标题 </summary>
    public string? PromptTitle { get; set; }

    /// <summary>输入提示正文 </summary>
    public string? Prompt { get; set; }
}