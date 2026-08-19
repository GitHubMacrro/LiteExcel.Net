namespace LiteExcel;

/// <summary>
/// 图片放置方式：InCell = 嵌入单元格（richData），Floating = 浮动图片（drawing）。
/// </summary>
public enum ImagePlacement
{
    /// <summary>嵌入单元格（Excel 365 InCell 图片，richData 体系）</summary>
    InCell,

    /// <summary>浮动图片（传统 drawing 锚点，可自由移动/缩放）</summary>
    Floating,
}

/// <summary>
/// 工作表中的一张图片。InCell 放置到指定单元格，Floating 以指定行列左上角为锚点。
/// </summary>
public sealed class WorksheetImage
{
    /// <summary>图片二进制数据（PNG/JPEG 等）</summary>
    public byte[] Data { get; set; } = Array.Empty<byte>();

    /// <summary>扩展名（不含点）。为 null 时按 magic bytes 自动探测 </summary>
    public string? Extension { get; set; }

    /// <summary>锚点行（1-based）</summary>
    public int Row { get; set; } = 1;

    /// <summary>锚点列（1-based）</summary>
    public int Column { get; set; } = 1;

    /// <summary>放置方式</summary>
    public ImagePlacement Placement { get; set; } = ImagePlacement.Floating;

    /// <summary>显示宽度（像素）。InCell 忽略；Floating 为 null 时按图片原始尺寸</summary>
    public double? WidthPx { get; set; }

    /// <summary>显示高度（像素）。InCell 忽略；Floating 为 null 时按图片原始尺寸</summary>
    public double? HeightPx { get; set; }

    /// <summary>图片名称（drawing 中的 cNvPr@name，可选）</summary>
    public string? Name { get; set; }

    /// <summary>有效图片扩展名（探测后）。仅写回时内部使用 </summary>
    internal string EffectiveExtension => NormalizeExtension(Extension) ?? DetectExtension(Data);

    /// <summary>全局 media 序号（image1..imageN），写回时由 XlsxWriter 统一分配 </summary>
    internal int MediaNumber { get; set; }

    /// <summary>解析后的像素尺寸（width, height）。探测失败返回 (0, 0) </summary>
    internal (int Width, int Height) PixelSize
    {
        get
        {
            var ext = EffectiveExtension;
            if (ext == "png") return ImageHeaders.ParsePng(Data);
            if (ext == "jpg" || ext == "jpeg") return ImageHeaders.ParseJpeg(Data);
            if (ext == "gif") return ImageHeaders.ParseGif(Data);
            if (ext == "bmp") return ImageHeaders.ParseBmp(Data);
            return (0, 0);
        }
    }

    private static string? NormalizeExtension(string? ext)
    {
        if (string.IsNullOrWhiteSpace(ext)) return null;
        ext = ext!.TrimStart('.').ToLowerInvariant();
        if (ext is "jpg" or "jpeg" or "png" or "gif" or "bmp") return ext == "jpeg" ? "jpg" : ext;
        return null;
    }

    private static string DetectExtension(byte[] data)
    {
        if (data is null || data.Length < 8) return "png";
        if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47) return "png";
        if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF) return "jpg";
        if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46) return "gif";
        if (data[0] == 0x42 && data[1] == 0x4D) return "bmp";
        return "png";
    }

    /// <summary>每英寸 96 像素（Excel 默认 DPI），1 英寸 = 914400 EMU </summary>
    internal const double EmuPerPixel = 914400.0 / 96.0;
}
