using System;

namespace LiteExcel;

/// <summary>
/// 单元格超链接。
/// 支持外部 URL/文件链接（<see cref="IsInternal"/> = false）与工作簿内部跳转（= true）。
/// </summary>
public sealed class Hyperlink
{
    /// <summary>链接目标。内部链接格式如 "#SheetName!A1"；外部为完整 URL 或文件路径 </summary>
    public string Target { get; set; } = "";

    /// <summary>鼠标悬停提示文本（可选） </summary>
    public string? Tooltip { get; set; }

    /// <summary>是否工作簿内部跳转。true 时 Target 以 '#' 开头指向内部单元格 </summary>
    public bool IsInternal { get; set; }

    /// <summary>深拷贝 </summary>
    public Hyperlink Clone() => new()
    {
        Target = Target,
        Tooltip = Tooltip,
        IsInternal = IsInternal,
    };

    public override bool Equals(object? obj) =>
        obj is Hyperlink other &&
        string.Equals(Target, other.Target, StringComparison.Ordinal) &&
        string.Equals(Tooltip, other.Tooltip, StringComparison.Ordinal) &&
        IsInternal == other.IsInternal;

    public override int GetHashCode()
    {
        int hash = 17;
        hash = hash * 31 + (Target?.GetHashCode() ?? 0);
        hash = hash * 31 + (Tooltip?.GetHashCode() ?? 0);
        hash = hash * 31 + (IsInternal ? 1 : 0);
        return hash;
    }
}
