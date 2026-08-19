using LiteExcel.Internal.Cfb;
using System.IO;

namespace LiteExcel.Internal;

/// <summary>
/// 加密工作簿识别器。
/// 带打开密码的 xlsx/xlsm/xlsb 实际是 OLE CFB 复合文档（内含 EncryptionInfo + EncryptedPackage 流），
/// 而非普通 zip 包。用于未提供打开密码时给出明确异常（提示提供 OpenPassword），
/// 避免误报为"zip 损坏"之类的无关错误。
/// </summary>
internal static class EncryptionDetector
{
    private static readonly byte[] CfbSignature =
        { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    /// <summary>
    /// 检查流是否为加密的 OOXML 工作簿（CFB 容器 + EncryptionInfo 流）。
    /// 是则返回 true；否则返回 false（可能是普通 zip 或其他格式）。
    /// 流位置在方法返回时已复位到起始位置，调用方可继续使用。
    /// <paramref name="path"/> 用于错误信息中的文件名展示；Stream 场景可传显示名（如 "&lt;stream&gt;"）。
    /// </summary>
    public static bool IsEncryptedOoxml(Stream stream, string path)
    {
        if (!LooksLikeCfb(stream)) return false;

        CfbFile? cfb = null;
        try
        {
            stream.Position = 0;
            cfb = CfbFile.Open(stream);
        }
        catch (LiteExcelException)
        {
            // CFB 结构非法：按原有路径报格式错误，不在这里拦截
            return false;
        }
        finally
        {
            // CfbFile.Open 会 CopyTo 读走整个流，无论识别结果如何都把位置复位，
            // 避免后续 zip 读取在流末尾开始而误报损坏
            stream.Position = 0;
        }

        return cfb.GetStream("EncryptionInfo") is not null;
    }

    /// <summary>
    /// 检查流是否为加密的 OOXML 工作簿，是则抛 <see cref="LiteExcelException"/>（未提供密码时的明确错误）。
    /// 流位置复位到起始位置。
    /// </summary>
    public static void ThrowIfEncryptedOoxml(Stream stream, string path)
    {
        if (IsEncryptedOoxml(stream, path))
        {
            throw new LiteExcelException(
                $"文件 '{SafeDisplayName(path)}' 已加密（带打开密码）。请通过 Excel.Open 的 ExcelReadOptions.OpenPassword 提供正确的打开密码。");
        }
    }

    /// <summary>安全提取用于错误信息的文件名。net48 的 Path.GetFileName 对非法路径字符（如 &lt;&gt;）会抛异常，这里兜底 </summary>
    private static string SafeDisplayName(string path)
    {
        if (string.IsNullOrEmpty(path)) return "文件";
        try
        {
            var name = Path.GetFileName(path);
            return string.IsNullOrEmpty(name) ? path : name;
        }
        catch (ArgumentException)
        {
            // 显示名不是合法路径（如 "<stream>"），原样展示
            return path;
        }
    }

    private static bool LooksLikeCfb(Stream stream)
    {
        long pos = stream.Position;
        try
        {
            var buf = new byte[8];
            int read = 0;
            while (read < 8)
            {
                int n = stream.Read(buf, read, 8 - read);
                if (n <= 0) break;
                read += n;
            }
            if (read < 8) return false;
            for (int i = 0; i < 8; i++)
                if (buf[i] != CfbSignature[i]) return false;
            return true;
        }
        finally
        {
            stream.Position = pos;
        }
    }
}
