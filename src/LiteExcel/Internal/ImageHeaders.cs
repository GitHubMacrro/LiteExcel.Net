namespace LiteExcel;

/// <summary>
/// 图片文件头解析：从 magic bytes 读取像素尺寸（无需 System.Drawing）。
/// </summary>
internal static class ImageHeaders
{
    /// <summary>PNG：8 字节签名 + 4 字节长度 + "IHDR" + 宽(4) + 高(4) </summary>
    public static (int W, int H) ParsePng(byte[] d)
    {
        if (d is null || d.Length < 24) return (0, 0);
        return (ReadBE32(d, 16), ReadBE32(d, 20));
    }

    /// <summary>JPEG：SOI 后扫描 SOF0/1/2/3/5/6/7/9/10/11 段 </summary>
    public static (int W, int H) ParseJpeg(byte[] d)
    {
        if (d is null || d.Length < 4 || d[0] != 0xFF || d[1] != 0xD8) return (0, 0);
        int i = 2;
        while (i + 4 <= d.Length)
        {
            if (d[i] != 0xFF) { i++; continue; }
            byte marker = d[i + 1];
            if (marker == 0xD8 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
            {
                i += 2;
                continue;
            }
            int len = ReadBE16(d, i + 2);
            if (len < 2 || i + 2 + len > d.Length) break;
            bool isSof = marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB;
            if (isSof && i + 9 <= d.Length)
            {
                int h = ReadBE16(d, i + 5);
                int w = ReadBE16(d, i + 7);
                return (w, h);
            }
            i += 2 + len;
        }
        return (0, 0);
    }

    /// <summary>GIF：6 字节签名后宽(2 小端) + 高(2 小端) </summary>
    public static (int W, int H) ParseGif(byte[] d)
    {
        if (d is null || d.Length < 10) return (0, 0);
        return (ReadLE16(d, 6), ReadLE16(d, 8));
    }

    /// <summary>BMP：偏移 18 宽(4 小端) + 偏移 22 高(4 小端) </summary>
    public static (int W, int H) ParseBmp(byte[] d)
    {
        if (d is null || d.Length < 26) return (0, 0);
        int h = ReadLE32(d, 22);
        return (ReadLE32(d, 18), Math.Abs(h));
    }

    private static int ReadBE16(byte[] d, int i) => (d[i] << 8) | d[i + 1];
    private static int ReadBE32(byte[] d, int i) => (d[i] << 24) | (d[i + 1] << 16) | (d[i + 2] << 8) | d[i + 3];
    private static int ReadLE16(byte[] d, int i) => d[i] | (d[i + 1] << 8);
    private static int ReadLE32(byte[] d, int i) => d[i] | (d[i + 1] << 8) | (d[i + 2] << 16) | (d[i + 3] << 24);
}
