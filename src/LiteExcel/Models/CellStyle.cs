namespace LiteExcel;

/// <summary>
/// 单元格样式 颜色统一使用 "#RRGGBB" 格式
/// </summary>
public sealed class CellStyle
{
    public string? FontName { get; set; }
    public double FontSize { get; set; } = 11;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public bool Strikeout { get; set; }
    public string? FontColor { get; set; }
    public string? FillColor { get; set; }
    public HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.General;
    public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Bottom;
    public bool WrapText { get; set; }
    public BorderStyle? Border { get; set; }

    public CellStyle Clone()
    {
        return new CellStyle
        {
            FontName = FontName,
            FontSize = FontSize,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            Strikeout = Strikeout,
            FontColor = FontColor,
            FillColor = FillColor,
            HorizontalAlignment = HorizontalAlignment,
            VerticalAlignment = VerticalAlignment,
            WrapText = WrapText,
            Border = Border?.Clone(),
        };
    }

    public override bool Equals(object? obj)
    {
        if (obj is not CellStyle other) return false;
        return FontName == other.FontName
            && FontSize == other.FontSize
            && Bold == other.Bold
            && Italic == other.Italic
            && Underline == other.Underline
            && Strikeout == other.Strikeout
            && FontColor == other.FontColor
            && FillColor == other.FillColor
            && HorizontalAlignment == other.HorizontalAlignment
            && VerticalAlignment == other.VerticalAlignment
            && WrapText == other.WrapText
            && Equals(Border, other.Border);
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(FontName);
        hash.Add(FontSize);
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(Underline);
        hash.Add(Strikeout);
        hash.Add(FontColor);
        hash.Add(FillColor);
        hash.Add(HorizontalAlignment);
        hash.Add(VerticalAlignment);
        hash.Add(WrapText);
        hash.Add(Border);
        return hash.ToHashCode();
    }
}

public enum HorizontalAlignment
{
    General,
    Left,
    Center,
    Right,
}

public enum VerticalAlignment
{
    Top,
    Center,
    Bottom,
}

public sealed class BorderStyle
{
    public BorderEdge? Top { get; set; }
    public BorderEdge? Bottom { get; set; }
    public BorderEdge? Left { get; set; }
    public BorderEdge? Right { get; set; }

    public BorderStyle Clone()
    {
        return new BorderStyle
        {
            Top = Top != null ? new BorderEdge { Style = Top.Style, Color = Top.Color } : null,
            Bottom = Bottom != null ? new BorderEdge { Style = Bottom.Style, Color = Bottom.Color } : null,
            Left = Left != null ? new BorderEdge { Style = Left.Style, Color = Left.Color } : null,
            Right = Right != null ? new BorderEdge { Style = Right.Style, Color = Right.Color } : null,
        };
    }

    public override bool Equals(object? obj)
    {
        if (obj is not BorderStyle other) return false;
        return EdgeEquals(Top, other.Top)
            && EdgeEquals(Bottom, other.Bottom)
            && EdgeEquals(Left, other.Left)
            && EdgeEquals(Right, other.Right);
    }

    private static bool EdgeEquals(BorderEdge? a, BorderEdge? b)
    {
        if (a is null && b is null) return true;
        if (a is null || b is null) return false;
        return a.Style == b.Style && a.Color == b.Color;
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(Top?.Style); hash.Add(Top?.Color);
        hash.Add(Bottom?.Style); hash.Add(Bottom?.Color);
        hash.Add(Left?.Style); hash.Add(Left?.Color);
        hash.Add(Right?.Style); hash.Add(Right?.Color);
        return hash.ToHashCode();
    }
}

public sealed class BorderEdge
{
    public string Style { get; set; } = "thin";
    public string? Color { get; set; }
}
