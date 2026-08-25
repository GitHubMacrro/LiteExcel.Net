namespace LiteExcel;

/// <summary>
/// 图标集样式（Excel 内置集合名）。图标长什么样由 Excel 程序内置渲染，文件只存集合名字符串。
/// 共 17 种：3/4/5 图标系列。
/// </summary>
public enum IconSetStyle
{
    /// <summary>3Arrows 三色箭头 </summary>
    ThreeArrows,
    ThreeArrowsGray,
    ThreeFlags,
    ThreeTrafficLights,
    ThreeTrafficLights2,
    ThreeSigns,
    ThreeSymbols,
    ThreeSymbols2,
    FourArrows,
    FourArrowsGray,
    FourRedToBlack,
    FourRating,
    FourTrafficLights,
    FiveArrows,
    FiveArrowsGray,
    FiveRating,
    FiveQuarters,
}

/// <summary>
/// 图标集（iconSet）条件格式参数。
/// </summary>
public sealed class IconSetInfo
{
    /// <summary>内置集合（默认三色箭头） </summary>
    public IconSetStyle Style { get; set; } = IconSetStyle.ThreeArrows;

    /// <summary>任意集合名字符串（如 Excel 未来新增集合）。非空时优先生效 </summary>
    public string? CustomStyleName { get; set; }

    /// <summary>阈值是否按百分比（true）或绝对数值（false），默认 true </summary>
    public bool Percent { get; set; } = true;

    /// <summary>单元格内是否同显数值，默认 true </summary>
    public bool ShowValue { get; set; } = true;

    /// <summary>自定义阈值（图标数 - 1 个，升序）。为空则按图标数均分百分比 </summary>
    public double[]? Thresholds { get; set; }

    /// <summary>写出到 iconSet 的集合名字符串 </summary>
    internal string SetName =>
        !string.IsNullOrEmpty(CustomStyleName) ? CustomStyleName!
        : Style switch
        {
            IconSetStyle.ThreeArrows => "3Arrows",
            IconSetStyle.ThreeArrowsGray => "3ArrowsGray",
            IconSetStyle.ThreeFlags => "3Flags",
            IconSetStyle.ThreeTrafficLights => "3TrafficLights1",
            IconSetStyle.ThreeTrafficLights2 => "3TrafficLights2",
            IconSetStyle.ThreeSigns => "3Signs",
            IconSetStyle.ThreeSymbols => "3Symbols",
            IconSetStyle.ThreeSymbols2 => "3Symbols2",
            IconSetStyle.FourArrows => "4Arrows",
            IconSetStyle.FourArrowsGray => "4ArrowsGray",
            IconSetStyle.FourRedToBlack => "4RedToBlack",
            IconSetStyle.FourRating => "4Rating",
            IconSetStyle.FourTrafficLights => "4TrafficLights",
            IconSetStyle.FiveArrows => "5Arrows",
            IconSetStyle.FiveArrowsGray => "5ArrowsGray",
            IconSetStyle.FiveRating => "5Rating",
            _ => "5Quarters",
        };

    /// <summary>该集合的图标数（3/4/5） </summary>
    internal int IconCount => SetName.StartsWith("3", System.StringComparison.Ordinal) ? 3
        : SetName.StartsWith("4", System.StringComparison.Ordinal) ? 4
        : 5;

    /// <summary>生成 cfvo 的 percent 值序列（图标数 - 1 个阈值 + 起始 0）。用户未给 Thresholds 时均分 </summary>
    internal double[] EffectiveThresholds()
    {
        if (Thresholds is { Length: > 0 })
            return Thresholds;
        int count = IconCount;
        var result = new double[count];
        for (int i = 0; i < count; i++)
            result[i] = i * 100.0 / count;
        return result;
    }
}
