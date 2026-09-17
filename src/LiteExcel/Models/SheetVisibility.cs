namespace LiteExcel;

/// <summary>
/// 工作表可见性。
/// 对应 OOXML <c>workbook.xml</c> 的 <c>&lt;sheet state&gt;</c>、XLSB 的 BrtBundleSh.hsState、XLS 的 BOUNDSHEET.grbit。
/// </summary>
public enum SheetVisibility
{
    /// <summary>可见（默认） </summary>
    Visible = 0,

    /// <summary>隐藏（用户可在 Excel 中取消隐藏） </summary>
    Hidden = 1,

    /// <summary>深度隐藏（只能在 VBA/编辑器里恢复） </summary>
    VeryHidden = 2,
}

/// <summary>工作表可见性在 OOXML / BIFF 之间的取值映射（内部使用）。</summary>
internal static class SheetVisibilityMap
{
    /// <summary>workbook.xml sheet@state → 枚举（缺失/未知按可见处理，与 Excel 一致）。</summary>
    public static SheetVisibility FromOoxml(string? state) => state switch
    {
        "hidden" => SheetVisibility.Hidden,
        "veryHidden" => SheetVisibility.VeryHidden,
        _ => SheetVisibility.Visible,
    };

    /// <summary>枚举 → workbook.xml sheet@state（可见返回 null，不写属性）。</summary>
    public static string? ToOoxml(SheetVisibility visibility) => visibility switch
    {
        SheetVisibility.Hidden => "hidden",
        SheetVisibility.VeryHidden => "veryHidden",
        _ => null,
    };

    /// <summary>BIFF hsState/grbit 低字节（0=visible/1=hidden/2=veryHidden） → 枚举。</summary>
    public static SheetVisibility FromBiff(int hsState) => hsState switch
    {
        1 => SheetVisibility.Hidden,
        2 => SheetVisibility.VeryHidden,
        _ => SheetVisibility.Visible,
    };

    /// <summary>枚举 → BIFF hsState/grbit 低字节（0/1/2）。</summary>
    public static int ToBiff(SheetVisibility visibility) => (int)visibility;
}
