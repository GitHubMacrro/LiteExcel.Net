using System.IO;
using System.Text;

namespace LiteExcel;

/// <summary>
/// CSV 格式后端（轻量，仅表格数据，不支持样式/合并/批注等 Excel 专有能力）。
/// 实现 RFC 4180 基础子集：双引号包裹含分隔符/换行/引号的字段。
/// </summary>
internal static class CsvBackend
{
    /// <summary>读取 CSV 文件为单张工作表的原始数据（首行不拆分为表头）。separator 为 null 时自动探测逗号/分号/Tab。</summary>
    public static SheetData Read(string path, char? separator = null)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Read(fs, Path.GetFileNameWithoutExtension(path), separator);
    }

    /// <summary>从流读取 CSV 为单张工作表的原始数据。sheetName 用于工作表命名，separator 为 null 时自动探测。</summary>
    public static SheetData Read(Stream stream, string sheetName = "Sheet1", char? separator = null)
    {
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        var encoding = DetectEncoding(ms) ?? Encoding.UTF8;
        var probe = new string(encoding.GetString(ms.ToArray()).Take(8192).ToArray());
        char actual = separator ?? DetectSeparator(probe);
        ms.Position = 0;
        using var reader = new StreamReader(ms, encoding, detectEncodingFromByteOrderMarks: true);
        return Read(reader, sheetName, actual);
    }

    internal static SheetData Read(TextReader reader, string sheetName, char separator)
    {
        var sheet = new SheetData { SheetName = sheetName };
        foreach (var fields in ReadRecords(reader, separator))
        {
            var cells = new List<Cell>(fields.Count);
            foreach (var f in fields)
                cells.Add(Cell.FromText(f));
            sheet.Rows.Add(cells);
        }
        return sheet;
    }

    /// <summary>
    /// 探测 CSV 分隔符（仅统计引号之外的分隔符频率，引号内不算）。
    /// 候选按 逗号 &gt; 分号 &gt; Tab 考察，取最多；三个候选都为零时回退逗号。
    /// </summary>
    internal static char DetectSeparator(string preview)
    {
        int comma = 0, semicolon = 0, tab = 0;
        bool inQuotes = false;
        foreach (char c in preview)
        {
            if (c == '"') inQuotes = !inQuotes;
            else if (!inQuotes)
            {
                if (c == ',') comma++;
                else if (c == ';') semicolon++;
                else if (c == '\t') tab++;
            }
        }
        if (semicolon > comma && semicolon >= tab) return ';';
        if (tab > comma && tab > semicolon) return '\t';
        return ',';
    }

    private static IEnumerable<List<string>> ReadRecords(TextReader reader, char separator)
    {
        var fields = new List<string>();
        var field = new StringBuilder();
        bool inQuotes = false;
        bool fieldStarted = false;
        bool recordHasCharacters = false;

        while (true)
        {
            int value = reader.Read();
            if (value < 0)
            {
                if (inQuotes)
                    throw new FormatException("CSV 字段缺少结束引号。");
                if (recordHasCharacters || fields.Count > 0 || field.Length > 0)
                {
                    fields.Add(field.ToString());
                    yield return fields;
                }
                yield break;
            }

            char ch = (char)value;
            recordHasCharacters = true;

            if (inQuotes)
            {
                if (ch == '"')
                {
                    int next = reader.Peek();
                    if (next == '"')
                    {
                        reader.Read();
                        field.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else if (ch == '\r' || ch == '\n')
                {
                    if (ch == '\r' && reader.Peek() == '\n') reader.Read();
                    field.Append('\n');
                }
                else
                {
                    field.Append(ch);
                }
                continue;
            }

            if (ch == '"' && !fieldStarted)
            {
                inQuotes = true;
                fieldStarted = true;
            }
            else if (ch == separator)
            {
                fields.Add(field.ToString());
                field.Clear();
                fieldStarted = false;
            }
            else if (ch == '\r' || ch == '\n')
            {
                if (ch == '\r' && reader.Peek() == '\n') reader.Read();
                fields.Add(field.ToString());
                yield return fields;
                fields = new List<string>();
                field.Clear();
                fieldStarted = false;
                recordHasCharacters = false;
            }
            else
            {
                field.Append(ch);
                fieldStarted = true;
            }
        }
    }

    /// <summary>写入 CSV 文件 </summary>
    public static void Write(string path, SheetData sheet, char? separator = null)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        Write(fs, sheet, onDegradation: null, targetFormat: ExcelFormat.Csv, separator: separator);
    }

    internal static void Write(Stream stream, SheetData sheet, Action<DegradationInfo>? onDegradation = null, ExcelFormat targetFormat = ExcelFormat.Csv, char? separator = null)
    {
        char sep = separator ?? ',';
        ReportDegradations(sheet, onDegradation, targetFormat);

        var sb = new StringBuilder();

        if (sheet.Headers is { Count: > 0 })
        {
            AppendRow(sb, sheet.Headers, sep);
        }

        foreach (var row in sheet.Rows)
        {
            var fields = new List<string>(row.Count);
            foreach (var cell in row)
                fields.Add(cell.GetString() ?? "");
            AppendRow(sb, fields, sep);
        }

        var bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        var text = sb.ToString();
        var data = bytes.GetBytes(text);
        stream.Write(data, 0, data.Length);
    }

    /// <summary>写出 CSV 时对静默丢弃的 Excel 专有能力逐项上报（P0-19 显式化） </summary>
    private static void ReportDegradations(SheetData sheet, Action<DegradationInfo>? onDegradation, ExcelFormat targetFormat)
    {
        if (onDegradation is null) return;

        void Report(DegradationCapability cap, string msg)
            => onDegradation(new DegradationInfo
            {
                Capability = cap,
                SheetName = sheet.SheetName,
                TargetFormat = targetFormat,
                Message = msg,
            });

        bool sheetHasStyles = sheet.DefaultStyle is not null
            || sheet.HeaderStyle is not null
            || (sheet.RowStyles is { Count: > 0 })
            || (sheet.ColumnStyles is { Count: > 0 });
        if (sheetHasStyles)
            Report(DegradationCapability.Styles, $"CSV 不支持样式，工作表 '{sheet.SheetName}' 的行/列/默认/表头样式已丢弃。");
        if (sheet.RowHeights is { Count: > 0 })
            Report(DegradationCapability.RowHeights, $"CSV 不支持行高，工作表 '{sheet.SheetName}' 的行高已丢弃。");
        if (sheet.ColumnWidths is { Count: > 0 })
            Report(DegradationCapability.ColumnWidths, $"CSV 不支持列宽，工作表 '{sheet.SheetName}' 的列宽已丢弃。");
        if (sheet.MergedRanges is { Count: > 0 })
            Report(DegradationCapability.MergedCells, $"CSV 不支持合并单元格，工作表 '{sheet.SheetName}' 的合并已丢弃。");
        if (sheet.FreezeRows > 0 || sheet.FreezeColumns > 0)
            Report(DegradationCapability.FreezePanes, $"CSV 不支持冻结窗格，工作表 '{sheet.SheetName}' 的冻结已丢弃。");
        if (sheet.Filter is not null)
            Report(DegradationCapability.AutoFilter, $"CSV 不支持自动筛选，工作表 '{sheet.SheetName}' 的筛选已丢弃。");
        if (sheet.Comments is { Count: > 0 })
            Report(DegradationCapability.Comments, $"CSV 不支持批注，工作表 '{sheet.SheetName}' 的批注已丢弃。");
        if (sheet.Validations is { Count: > 0 })
            Report(DegradationCapability.DataValidation, $"CSV 不支持数据验证，工作表 '{sheet.SheetName}' 的数据验证已丢弃。");
        if (sheet.Images is { Count: > 0 })
        {
            Report(DegradationCapability.Images, $"CSV 不支持图片，工作表 '{sheet.SheetName}' 的图片已丢弃。");
            if (sheet.Images.Any(i => i.Placement == ImagePlacement.InCell))
                Report(DegradationCapability.RichData, $"CSV 不支持 InCell 图片，工作表 '{sheet.SheetName}' 的 InCell 图片已丢弃。");
        }
        if (sheet.Tables is { Count: > 0 })
            Report(DegradationCapability.Tables, $"CSV 不支持超级表，工作表 '{sheet.SheetName}' 的超级表已丢弃。");

        foreach (var row in sheet.Rows)
        {
            foreach (var cell in row)
            {
                if (cell.Style is not null || !string.IsNullOrEmpty(cell.NumberFormat))
                {
                    if (!sheetHasStyles)
                        Report(DegradationCapability.Styles, $"CSV 不支持单元格样式，工作表 '{sheet.SheetName}' 的样式已丢弃。");
                    sheetHasStyles = true;
                }
                if (cell.Hyperlink is not null)
                    Report(DegradationCapability.Hyperlinks, $"CSV 不支持超链接，工作表 '{sheet.SheetName}' 的超链接已丢弃。");
                if (cell.IsFormula || !string.IsNullOrEmpty(cell.Formula))
                    Report(DegradationCapability.Formulas, $"CSV 不支持公式，工作表 '{sheet.SheetName}' 的公式以文本值写出。");
            }
        }
        if (sheet.ConditionalFormats is { Count: > 0 })
            Report(DegradationCapability.ConditionalFormatting, $"CSV 不支持条件格式，工作表 '{sheet.SheetName}' 的条件格式已丢弃。");
    }

    private static void AppendRow(StringBuilder sb, IReadOnlyList<string> fields, char separator)
    {
        for (int i = 0; i < fields.Count; i++)
        {
            if (i > 0) sb.Append(separator);
            AppendField(sb, fields[i], separator);
        }
        sb.Append('\n');
    }

    private static void AppendField(StringBuilder sb, string field, char separator)
    {
        bool needQuote = field.IndexOf(separator) >= 0
            || field.IndexOfAny(new[] { '"', '\n', '\r' }) >= 0;
        if (!needQuote)
        {
            sb.Append(field);
            return;
        }
        sb.Append('"');
        sb.Append(field.Replace("\"", "\"\""));
        sb.Append('"');
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
