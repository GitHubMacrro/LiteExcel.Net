using System.IO;
using System.Text;

namespace LiteExcel;

/// <summary>
/// CSV 格式后端（轻量，仅表格数据，不支持样式/合并/批注等 Excel 专有能力）。
/// 实现 RFC 4180 基础子集：双引号包裹含分隔符/换行/引号的字段。
/// </summary>
internal static class CsvBackend
{
    /// <summary>读取 CSV 文件为单张工作表的原始数据（首行不拆分为表头） </summary>
    public static SheetData Read(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Read(fs, Path.GetFileNameWithoutExtension(path));
    }

    /// <summary>从流读取 CSV 为单张工作表的原始数据。sheetName 用于工作表命名 </summary>
    public static SheetData Read(Stream stream, string sheetName = "Sheet1")
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;
        using var reader = new StreamReader(ms, DetectEncoding(ms) ?? Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return Read(reader, sheetName);
    }

    internal static SheetData Read(TextReader reader, string sheetName)
    {
        var sheet = new SheetData { SheetName = sheetName };
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (string.IsNullOrEmpty(line)) continue;
            var fields = ParseLine(line);
            var cells = new List<Cell>(fields.Count);
            foreach (var f in fields)
                cells.Add(Cell.FromText(f));
            sheet.Rows.Add(cells);
        }
        return sheet;
    }

    /// <summary>写入 CSV 文件 </summary>
    public static void Write(string path, SheetData sheet)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        Write(fs, sheet);
    }

    internal static void Write(Stream stream, SheetData sheet)
    {
        var sb = new StringBuilder();

        if (sheet.Headers is { Count: > 0 })
        {
            AppendRow(sb, sheet.Headers);
        }

        foreach (var row in sheet.Rows)
        {
            var fields = new List<string>(row.Count);
            foreach (var cell in row)
                fields.Add(cell.GetString() ?? "");
            AppendRow(sb, fields);
        }

        var bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        var text = sb.ToString();
        var data = bytes.GetBytes(text);
        stream.Write(data, 0, data.Length);
    }

    private static void AppendRow(StringBuilder sb, IReadOnlyList<string> fields)
    {
        for (int i = 0; i < fields.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendField(sb, fields[i]);
        }
        sb.Append('\n');
    }

    private static void AppendField(StringBuilder sb, string field)
    {
        bool needQuote = field.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0;
        if (!needQuote)
        {
            sb.Append(field);
            return;
        }
        sb.Append('"');
        sb.Append(field.Replace("\"", "\"\""));
        sb.Append('"');
    }

    private static List<string> ParseLine(string line)
    {
        var fields = new List<string>();
        var sb = new StringBuilder();
        bool inQuotes = false;
        bool fieldStarted = false;

        for (int i = 0; i < line.Length; i++)
        {
            char ch = line[i];
            if (inQuotes)
            {
                if (ch == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    sb.Append(ch);
                }
            }
            else
            {
                if (ch == '"' && !fieldStarted)
                {
                    inQuotes = true;
                    fieldStarted = true;
                }
                else if (ch == ',')
                {
                    fields.Add(sb.ToString());
                    sb.Clear();
                    fieldStarted = false;
                }
                else
                {
                    sb.Append(ch);
                    fieldStarted = true;
                }
            }
        }

        fields.Add(sb.ToString());
        return fields;
    }

    private static Encoding? DetectEncoding(Stream fs)
    {
        // BOM 检测
        if (!fs.CanSeek) return null;
        var bom = new byte[3];
        int read = fs.Read(bom, 0, 3);
        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
        fs.Position = 0;
        return null;
    }
}
