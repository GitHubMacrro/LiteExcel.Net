using System.IO;
using System.IO.Compression;

namespace LiteExcel;

public static partial class XlsxReader
{
    /// <summary>
    /// 以原始模式（firstRowIsHeader=false）读取所有工作表。
    /// 首行不拆分进 Headers，所有行都保留在 Rows，供高层 Workbook/Worksheet 直接使用。
    /// </summary>
    internal static List<SheetData> ReadAllRaw(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadAllRaw(fs);
    }

    /// <summary>原始模式读取所有工作表（Stream 重载） </summary>
    internal static List<SheetData> ReadAllRaw(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var shared = ReadSharedStrings(zip);
        var styles = ReadStyles(zip);
        var sheets = ReadWorkbook(zip);

        var result = new List<SheetData>(sheets.Count);
        foreach (var info in sheets)
            result.Add(ReadWorksheet(zip, info.Path, info.Name, shared, styles, firstRowIsHeader: false));
        return result;
    }
}
