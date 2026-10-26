namespace LiteExcel;

/// <summary>富文本 run 的垂直位置（OOXML vertAlign）</summary>
public enum TextVerticalAlign
{
    Baseline,
    Superscript,
    Subscript,
}

/// <summary>富文本中的一段连续文本及其字体格式</summary>
public sealed class TextRun
{
    public string Text { get; set; } = "";

    /// <summary>本段字体（复用 CellStyle 的字体字段；填充、边框、对齐被忽略） </summary>
    public CellStyle? Style { get; set; }

    public TextVerticalAlign? VerticalAlign { get; set; }

    public TextRun Clone() => new()
    {
        Text = Text,
        Style = Style?.Clone(),
        VerticalAlign = VerticalAlign,
    };
}

/// <summary>单元格富文本：由若干格式不同的文本段组成，写出后仍是同一个单元格</summary>
public sealed class RichText
{
    public List<TextRun> Runs { get; } = new();

    /// <summary>各段拼接后的纯文本 </summary>
    public string Text
    {
        get
        {
            var sb = new System.Text.StringBuilder();
            foreach (var r in Runs) sb.Append(r.Text);
            return sb.ToString();
        }
    }

    public RichText() { }

    public RichText(IEnumerable<TextRun> runs)
    {
        Runs.AddRange(runs);
    }

    public RichText Add(string text, CellStyle? style = null, TextVerticalAlign? verticalAlign = null)
    {
        Runs.Add(new TextRun { Text = text, Style = style, VerticalAlign = verticalAlign });
        return this;
    }

    public RichText Clone()
    {
        var clone = new RichText();
        foreach (var r in Runs) clone.Runs.Add(r.Clone());
        return clone;
    }
}
