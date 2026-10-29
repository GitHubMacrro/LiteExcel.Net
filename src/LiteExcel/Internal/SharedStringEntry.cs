namespace LiteExcel.Internal;

/// <summary>共享字符串表条目：压平文本 + 可选富文本段</summary>
internal readonly struct SharedStringEntry
{
    public readonly string Text;
    public readonly RichText? Rich;

    public SharedStringEntry(string text, RichText? rich)
    {
        Text = text;
        Rich = rich;
    }
}
