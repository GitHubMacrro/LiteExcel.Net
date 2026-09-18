using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using LiteExcel.Internal.Biff12;

namespace LiteExcel.Internal;

/// <summary>
/// .xlsb（BIFF12 二进制 OOXML 变体）读取后端。
/// 容器仍是 ZIP（与 xlsx 相同的 OPC 包），部件内为二进制记录流。
/// 仅读取：数据单元格（含公式缓存值）、共享字符串、日期识别、合并单元格、列宽、行高、冻结表头。
/// </summary>
internal static class XlsbBackend
{
    // workbook.bin
    private const int BrtBundleSh = 0x009C;   // 工作表清单条目
    private const int BrtWbProp = 0x0099;     // 工作簿属性（date1904 标志）
    private const int BrtExternSheet = 0x016A; // 外部表引用（ixti → itab 映射）
    private const int BrtDefinedName = 0x0027; // 定义名称
    private const int BrtFileSharing = 0x0224;      // 写保护（旧式）
    private const int BrtFileSharingIso = 0x02A4;   // 写保护（ISO 盐化哈希）

    // sharedStrings.bin
    private const int BrtBeginSst = 0x009F;
    private const int BrtSSTItem = 0x0013;
    private const int BrtEndSst = 0x00A0;

    // styles.bin
    private const int BrtFmt = 0x002C;        // 自定义数字格式
    private const int BrtXf = 0x002F;         // 单元格样式 XF
    private const int BrtBeginCellXfs = 0x0269;
    private const int BrtEndCellXfs = 0x026A;
    private const int BrtFont = 0x002B;       // 字体
    private const int BrtBeginFonts = 0x0263;
    private const int BrtEndFonts = 0x0264;
    private const int BrtFill = 0x002D;       // 填充
    private const int BrtBeginFills = 0x025B;
    private const int BrtEndFills = 0x025C;
    private const int BrtBorder = 0x002E;     // 边框
    private const int BrtBeginBorders = 0x0265;
    private const int BrtEndBorders = 0x0266;

    // table（超级表）部件记录
    private const int BrtBeginList = 0x0157;        // 表属性
    private const int BrtBeginListCol = 0x015B;     // 列定义
    private const int BrtTableStyleClient = 0x0201; // 表样式
    // worksheet 表引用
    private const int BrtTablePart = 0x0295;        // 表部件引用（XLWideString rId）
    // 数据验证
    private const int BrtBeginDVs = 0x023D;         // cDVs(4) + reserved(4) + ...
    private const int BrtEndDVs = 0x023E;
    private const int BrtDVal = 0x0040;             // 单条数据验证

    // worksheet.bin
    private const int BrtRowHdr = 0x0000;
    private const int BrtCellBlank = 0x0001;
    private const int BrtCellRk = 0x0002;
    private const int BrtCellError = 0x0003;
    private const int BrtCellBool = 0x0004;
    private const int BrtCellReal = 0x0005;
    private const int BrtCellSt = 0x0006;
    private const int BrtCellIsst = 0x0007;
    private const int BrtFmlaString = 0x0008;
    private const int BrtFmlaNum = 0x0009;
    private const int BrtFmlaBool = 0x000A;
    private const int BrtFmlaError = 0x000B;
    private const int BrtShortBlank = 0x000C;
    private const int BrtShortRk = 0x000D;
    private const int BrtShortError = 0x000E;
    private const int BrtShortBool = 0x000F;
    private const int BrtShortReal = 0x0010;
    private const int BrtShortSt = 0x0011;
    private const int BrtShortIsst = 0x0012;
    private const int BrtColInfo = 0x003C;
    private const int BrtWsDim = 0x0094;
    private const int BrtPane = 0x0097;
    private const int BrtMergeCell = 0x00B0;
    private const int BrtBeginAFilter = 0x00A1;
    private const int BrtEndAFilter = 0x00A2;
    private const int BrtHLink = 0x01EE;

    public static List<SheetData> ReadAll(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadAll(fs);
    }

    /// <summary>读取 .xlsb 包中的 VBA 宏工程原始字节（xl/vbaProject.bin），无宏返回 null </summary>
    public static byte[]? ReadVbaProject(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadVbaProject(fs);
    }

    /// <summary>从流读取 .xlsb 包中的 VBA 宏工程原始字节，无宏返回 null。流必须可读 </summary>
    public static byte[]? ReadVbaProject(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        return ReadEntry(zip, "xl/vbaProject.bin");
    }

    /// <summary>读取 .xlsb 工作簿宿主的 VBA 代码名（BrtWbProp 内 codeName），无则返回 null </summary>
    public static string? ReadWorkbookCodeName(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadWorkbookCodeName(fs);
    }

    /// <summary>从流读取 .xlsb 工作簿宿主的 VBA 代码名，无则返回 null </summary>
    public static string? ReadWorkbookCodeName(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        return ReadWorkbookCodeNameCore(zip);
    }

    private static string? ReadWorkbookCodeNameCore(ZipArchive zip)
    {
        var wbBytes = ReadEntry(zip, "xl/workbook.bin");
        if (wbBytes is null) return null;
        var records = Biff12Records.ReadAll(wbBytes);
        foreach (var rec in records)
        {
            if (rec.Rt != BrtWbProp || rec.Data.Length < 9) continue;
            int off = 8;
            uint cch = Biff12Records.ReadU32(rec.Data, off);
            off += 4;
            if (cch == 0 || cch == 0xFFFFFFFF || off + (int)cch * 2 > rec.Data.Length) continue;
            return System.Text.Encoding.Unicode.GetString(rec.Data, off, (int)cch * 2);
        }
        return null;
    }

    /// <summary>读取 .xlsb 工作簿的 1904 日期系统标志（BrtWbProp flags bit0） </summary>
    public static bool ReadDate1904(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadDate1904(fs);
    }

    /// <summary>从流读取 .xlsb 工作簿的 1904 日期系统标志 </summary>
    public static bool ReadDate1904(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var wbBytes = ReadEntry(zip, "xl/workbook.bin");
        if (wbBytes is null) return false;
        var records = Biff12Records.ReadAll(wbBytes);
        foreach (var rec in records)
        {
            if (rec.Rt != BrtWbProp || rec.Data.Length < 4) continue;
            return (Biff12Records.ReadU32(rec.Data, 0) & 0x01) != 0;
        }
        return false;
    }

    /// <summary>读取 .xlsb 工作簿的写保护（fileSharing）信息。无则返回 null </summary>
    public static Internal.Encryption.FileSharingInfo? ReadFileSharing(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadFileSharing(fs);
    }

    /// <summary>从流读取 .xlsb 工作簿的写保护（fileSharing）信息。无则返回 null </summary>
    public static Internal.Encryption.FileSharingInfo? ReadFileSharing(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var wbBytes = ReadEntry(zip, "xl/workbook.bin");
        if (wbBytes is null) return null;
        var records = Biff12Records.ReadAll(wbBytes);
        foreach (var rec in records)
        {
            if (rec.Rt != BrtFileSharing && rec.Rt != BrtFileSharingIso) continue;
            var info = ParseFileSharing(rec.Data);
            if (info is not null) return info;
        }
        return null;
    }

    /// <summary>
    /// 解析 BrtFileSharingIso(0x02A4) / BrtFileSharing(0x0224) 记录。
    /// Iso 布局（样本）：spinCount(4) + flags(2) + stUser(XLWideString) + hashValue(4+bytes) + saltValue(4+bytes) + algorithmName(XLWideString)。
    /// 旧式 0x0224 无哈希参数（仅标记写保护）。
    /// </summary>
    private static Internal.Encryption.FileSharingInfo? ParseFileSharing(byte[] d)
    {
        if (d.Length < 6) return null;
        int spin = Biff12Records.ReadS32(d, 0);
        bool readOnlyRecommended = (Biff12Records.ReadU16(d, 4) & 0x01) != 0;

        int off = 6;
        if (off + 4 > d.Length) return null;
        int userCch = Biff12Records.ReadS32(d, off); off += 4;
        off += userCch * 2;
        if (off > d.Length) return null;

        // 旧式 0x0224：仅 stUser + flags，无哈希参数
        if (off + 4 > d.Length) return null;
        int hashLen = Biff12Records.ReadS32(d, off); off += 4;
        if (hashLen <= 0 || hashLen > 512 || off + hashLen > d.Length) return null;
        var hash = new byte[hashLen];
        Array.Copy(d, off, hash, 0, hashLen);
        off += hashLen;

        if (off + 4 > d.Length) return null;
        int saltLen = Biff12Records.ReadS32(d, off); off += 4;
        byte[]? salt = null;
        if (saltLen > 0 && saltLen <= 512 && off + saltLen <= d.Length)
        {
            salt = new byte[saltLen];
            Array.Copy(d, off, salt, 0, saltLen);
            off += saltLen;
        }

        string algorithmName = "";
        if (off + 4 <= d.Length)
        {
            int algoCch = Biff12Records.ReadS32(d, off); off += 4;
            if (algoCch > 0 && algoCch < 64 && off + algoCch * 2 <= d.Length)
                algorithmName = System.Text.Encoding.Unicode.GetString(d, off, algoCch * 2);
        }

        if (string.IsNullOrEmpty(algorithmName)) algorithmName = "SHA-512";
        return new Internal.Encryption.FileSharingInfo(hash, salt, algorithmName, spin, readOnlyRecommended);
    }

    public static List<SheetData> ReadAll(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var wbBytes = ReadEntry(zip, "xl/workbook.bin")
            ?? throw new LiteExcelException(".xlsb 文件中缺少 xl/workbook.bin");
        var (sheets, date1904) = ParseWorkbook(wbBytes);

        var sstBytes = ReadEntry(zip, "xl/sharedStrings.bin");
        var sst = sstBytes is not null ? ParseSharedStrings(sstBytes) : new List<string>();

        var stylesBytes = ReadEntry(zip, "xl/styles.bin");
        var (formats, cellXfs, cellStyles) = stylesBytes is not null
            ? ParseStyles(stylesBytes)
            : (new Dictionary<int, string>(), new List<int> { 0 }, new List<CellStyle?> { null });

        var sheetPaths = MapSheetPaths(zip, sheets);

        var result = new List<SheetData>(sheets.Count);
        for (int i = 0; i < sheets.Count; i++)
        {
            var data = ReadEntry(zip, sheetPaths[i]);
            if (data is null)
                throw new LiteExcelException($"缺少工作表文件: {sheetPaths[i]}");
            var rels = ReadSheetHyperlinkRels(zip, sheetPaths[i]);
            var sd = ParseWorksheet(data, sheets[i].Name, sst, formats, cellXfs, cellStyles, date1904, rels);
            sd.SheetState = SheetVisibilityMap.ToOoxml(SheetVisibilityMap.FromBiff(sheets[i].HsState));
            ReadCommentsForSheet(zip, sheetPaths[i], sd);
            ReadTablesForSheet(zip, sheetPaths[i], data, sd);
            result.Add(sd);
        }

        if (result.Count == 0)
            throw new LiteExcelException("这不是有效的 .xlsb 文件（未找到任何工作表）");
        return result;
    }

    private static byte[]? ReadEntry(ZipArchive zip, string name)
    {
        var entry = zip.GetEntry(name);
        if (entry is null) return null;
        using var s = entry.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>将工作簿清单中的 rId 映射到实际工作表部件路径。</summary>
    private static List<string> MapSheetPaths(ZipArchive zip, List<(string Name, string RelId, int HsState)> sheets)
    {
        var relMap = new Dictionary<string, string>();
        var relsEntry = zip.GetEntry("xl/_rels/workbook.bin.rels");
        if (relsEntry is not null)
        {
            try
            {
                var rels = XElement.Load(relsEntry.Open());
                var relNs = rels.Name.Namespace;
                foreach (var rel in rels.Elements(relNs + "Relationship"))
                {
                    var id = rel.Attribute("Id")?.Value;
                    var target = rel.Attribute("Target")?.Value ?? "";
                    if (id is not null) relMap[id] = target;
                }
            }
            catch
            {
                relMap.Clear(); // 关系文件损坏时退回按序号猜测
            }
        }

        var result = new List<string>(sheets.Count);
        for (int i = 0; i < sheets.Count; i++)
        {
            string path = "";
            if (relMap.TryGetValue(sheets[i].RelId, out var target) && !string.IsNullOrEmpty(target))
            {
                path = target.StartsWith("/")
                    ? target.TrimStart('/')
                    : target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
                        ? target
                        : "xl/" + target;
            }
            if (string.IsNullOrEmpty(path))
                path = $"xl/worksheets/sheet{i + 1}.bin";
            result.Add(path);
        }
        return result;
    }

    /// <summary>
    /// 为流式读取预加载工作簿级共享数据：工作表清单、SST、样式、日期系统。
    /// 返回 (sheets, sst, formats, cellXfs, date1904)。
    /// </summary>
    public static (List<(string Name, string RelId, int HsState)> sheets, List<string> sst,
        Dictionary<int, string> formats, List<int> cellXfs, bool date1904)
        PrepareStreaming(ZipArchive zip)
    {
        var wbBytes = ReadEntry(zip, "xl/workbook.bin")
            ?? throw new LiteExcelException(".xlsb 文件中缺少 xl/workbook.bin");
        var (sheets, date1904) = ParseWorkbook(wbBytes);

        var sstBytes = ReadEntry(zip, "xl/sharedStrings.bin");
        var sst = sstBytes is not null ? ParseSharedStrings(sstBytes) : new List<string>();

        var stylesBytes = ReadEntry(zip, "xl/styles.bin");
        var (formats, cellXfs, _) = stylesBytes is not null
            ? ParseStyles(stylesBytes)
            : (new Dictionary<int, string>(), new List<int> { 0 }, new List<CellStyle?> { null });

        return (sheets, sst, formats, cellXfs, date1904);
    }

    /// <summary>将工作簿清单中的 rId 映射到实际工作表部件路径（公开给流式读取器）。</summary>
    public static List<string> MapSheetPathsPublic(ZipArchive zip, List<(string Name, string RelId, int HsState)> sheets)
        => MapSheetPaths(zip, sheets);

    private static (List<(string Name, string RelId, int HsState)> Sheets, bool Date1904) ParseWorkbook(byte[] wb)
    {
        var records = Biff12Records.ReadAll(wb);
        if (records.Count == 0)
            throw new LiteExcelException("这不是有效的 .xlsb 文件（workbook.bin 为空）");

        var sheets = new List<(string, string, int)>();
        bool date1904 = false;

        foreach (var rec in records)
        {
            switch (rec.Rt)
            {
                case BrtBundleSh:
                {
                    var d = rec.Data;
                    if (d.Length < 8) break;
                    uint hsState = Biff12Records.ReadU32(d, 0);
                    int off = 8; // Hidden(4) + iTabID(4)
                    var relId = Biff12Records.ReadWideString(d, ref off);
                    var name = Biff12Records.ReadWideString(d, ref off);
                    sheets.Add((name, relId, (int)hsState));
                    break;
                }
                case BrtWbProp:
                    if (rec.Data.Length >= 4)
                        date1904 = (Biff12Records.ReadU32(rec.Data, 0) & 0x01) != 0;
                    break;
            }
        }

        if (sheets.Count == 0)
            throw new LiteExcelException("这不是有效的 .xlsb 文件（未找到任何工作表）");
        return (sheets, date1904);
    }

    /// <summary>从已打开的 zip 读取定义名称（供 Excel.cs 打开路径复用同一文件快照）。</summary>
    public static (List<NamedRange> Names, bool HasUnsupported) ReadDefinedNames(ZipArchive zip, List<string> sheetNames)
    {
        var wbBytes = ReadEntry(zip, "xl/workbook.bin");
        if (wbBytes is null) return (new List<NamedRange>(), false);
        return ParseDefinedNames(wbBytes, sheetNames);
    }

    /// <summary>
    /// 解析 workbook.bin 的定义名称（BrtDefinedName）。rgce 仅解码单引用/区域/常量；
    /// 复合表达式跳过（不产出错误引用）。返回 (名称, 是否含未支持项)。
    /// </summary>
    public static (List<NamedRange> Names, bool HasUnsupported) ParseDefinedNames(byte[] wb, List<string> sheetNames)
    {
        var records = Biff12Records.ReadAll(wb);
        var result = new List<NamedRange>();
        bool unsupported = false;

        // 先扫 BrtExternSheet 构建 ixti → sheet 名
        var sheetNameByIxti = new List<string>();
        foreach (var rec in records)
        {
            if (rec.Rt != BrtExternSheet || rec.Data.Length < 4) continue;
            int cXti = (int)Biff12Records.ReadU32(rec.Data, 0);
            for (int i = 0; i < cXti; i++)
            {
                int o = 4 + i * 12;
                if (o + 12 > rec.Data.Length) break;
                int itabFirst = (int)Biff12Records.ReadU32(rec.Data, o + 4);
                // itab 为 0-based sheet 索引；越界（如 -1 表示删除的表）记为空
                sheetNameByIxti.Add(itabFirst >= 0 && itabFirst < sheetNames.Count ? sheetNames[itabFirst] : "");
            }
            break;
        }

        foreach (var rec in records)
        {
            if (rec.Rt != BrtDefinedName || rec.Data.Length < 13) continue;
            var d = rec.Data;
            int itab = (int)Biff12Records.ReadU32(d, 5);
            int off = 9;
            var name = Biff12Records.ReadWideString(d, ref off);
            if (off + 4 > d.Length) continue;
            int cce = (int)Biff12Records.ReadU32(d, off); off += 4;
            if (cce < 0 || off + cce > d.Length) continue;
            var rgce = new byte[cce];
            Array.Copy(d, off, rgce, 0, cce);

            var reference = RgceDecoder.Decode(rgce, sheetNameByIxti);
            if (reference is null) { unsupported = true; continue; }

            result.Add(new NamedRange
            {
                Name = name,
                Reference = reference,
                LocalSheetId = itab, // 0-based；-1 = 全局
            });
        }

        return (result, unsupported);
    }

    private static List<string> ParseSharedStrings(byte[] data)
    {
        var records = Biff12Records.ReadAll(data);
        var result = new List<string>();
        foreach (var rec in records)
        {
            if (rec.Rt == BrtSSTItem)
            {
                var d = rec.Data;
                if (d.Length < 5) continue;
                // BrtSSTItem = RichStr: flags(1) + XLWideString；富文本 run 数据忽略
                int off = 1;
                result.Add(Biff12Records.ReadWideString(d, ref off));
            }
            else if (rec.Rt == BrtEndSst)
            {
                break;
            }
        }
        return result;
    }

    private static (Dictionary<int, string> Formats, List<int> CellXfs, List<CellStyle?> CellStyles) ParseStyles(byte[] data)
    {
        var records = Biff12Records.ReadAll(data);
        var formats = new Dictionary<int, string>();
        var cellXfs = new List<int>(); // BrtBeginCellXFs 内按序排列，索引 0 即默认样式
        var fonts = new List<CellStyle?>();
        var fills = new List<string?>();
        var borders = new List<BorderStyle?>();
        var xfRefs = new List<(int FontId, int FillId, int BorderId)>();
        bool inCellXfs = false, inFonts = false, inFills = false, inBorders = false;

        foreach (var rec in records)
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case BrtFmt:
                    if (d.Length >= 6)
                    {
                        int numFmtId = Biff12Records.ReadU16(d, 0);
                        int off = 2;
                        formats[numFmtId] = Biff12Records.ReadWideString(d, ref off);
                    }
                    break;
                case BrtBeginFonts: inFonts = true; break;
                case BrtEndFonts: inFonts = false; break;
                case BrtFont: if (inFonts) fonts.Add(ParseFont(d)); break;
                case BrtBeginFills: inFills = true; break;
                case BrtEndFills: inFills = false; break;
                case BrtFill: if (inFills) fills.Add(ParseFillColor(d)); break;
                case BrtBeginBorders: inBorders = true; break;
                case BrtEndBorders: inBorders = false; break;
                case BrtBorder: if (inBorders) borders.Add(ParseBorder(d)); break;
                case BrtBeginCellXfs: inCellXfs = true; break;
                case BrtEndCellXfs: inCellXfs = false; break;
                case BrtXf:
                    if (inCellXfs && d.Length >= 10)
                    {
                        int ifmt = Biff12Records.ReadU16(d, 2);
                        int fontId = Biff12Records.ReadU16(d, 4);
                        int fillId = Biff12Records.ReadU16(d, 6);
                        int borderId = Biff12Records.ReadU16(d, 8);
                        cellXfs.Add(ifmt);
                        xfRefs.Add((fontId, fillId, borderId));
                    }
                    break;
            }
        }

        // 每个 ixfe → CellStyle（默认 XF 0/0/0 返回 null，保持轻量）
        var styles = new List<CellStyle?>(xfRefs.Count);
        foreach (var (fontId, fillId, borderId) in xfRefs)
        {
            var font = fontId >= 0 && fontId < fonts.Count ? fonts[fontId] : null;
            var fill = fillId >= 0 && fillId < fills.Count ? fills[fillId] : null;
            var border = borderId >= 0 && borderId < borders.Count ? borders[borderId] : null;
            if (font is null && fill is null && border is null) { styles.Add(null); continue; }
            var s = font ?? new CellStyle();
            s.FillColor = fill;
            s.Border = border;
            styles.Add(s);
        }

        return (formats, cellXfs, styles);
    }

    /// <summary>解析 BrtFont（29+ 字节）。布局：sz(2 twips) grbit(2) weight(2) vertAlign(2) underline(1)
    /// family(1) charset(1) pad(1) BrtColor(8) scheme(1) name(XLWideString)。</summary>
    private static CellStyle? ParseFont(byte[] d)
    {
        if (d.Length < 25) return null;
        int szTwips = Biff12Records.ReadU16(d, 0);
        int grbit = Biff12Records.ReadU16(d, 2);
        int weight = Biff12Records.ReadU16(d, 4);
        int underline = d[8];
        int off = 21;
        var name = Biff12Records.ReadWideString(d, ref off);
        return new CellStyle
        {
            FontName = string.IsNullOrEmpty(name) ? null : name,
            FontSize = szTwips > 0 ? szTwips / 20.0 : 11,
            Bold = (grbit & 0x01) != 0 || weight >= 700,
            Italic = (grbit & 0x02) != 0,
            Strikeout = (grbit & 0x08) != 0,
            Underline = underline != 0,
            FontColor = ColorRgb(d, 12),
        };
    }

    /// <summary>解析 BrtFill 的填充色：fls(4) + fgColor(BrtColor 8) + bgColor(8) + 12×u32。仅 solid(fls=1) 返回颜色。</summary>
    private static string? ParseFillColor(byte[] d)
    {
        if (d.Length < 12) return null;
        int fls = (int)Biff12Records.ReadU32(d, 0);
        if (fls != 1) return null; // 1 = solid；0=无、0x11=gray125 等不映射
        return ColorRgb(d, 4);
    }

    /// <summary>解析 BrtBorder：diagonal(1) + 5×(style(1)+reserved(1)+BrtColor(8))，边序 left/right/top/bottom/diagonal。</summary>
    private static BorderStyle? ParseBorder(byte[] d)
    {
        if (d.Length < 51) return null;
        var left = ParseBorderEdge(d, 1);
        var right = ParseBorderEdge(d, 11);
        var top = ParseBorderEdge(d, 21);
        var bottom = ParseBorderEdge(d, 31);
        if (left is null && right is null && top is null && bottom is null) return null;
        return new BorderStyle { Left = left, Right = right, Top = top, Bottom = bottom };
    }

    private static BorderEdge? ParseBorderEdge(byte[] d, int off)
    {
        var name = BorderStyleName(d[off]);
        if (name is null) return null;
        return new BorderEdge { Style = name, Color = ColorRgb(d, off + 2) };
    }

    private static string? BorderStyleName(int style) => style switch
    {
        0 => null,
        1 => "thin",
        2 => "medium",
        3 => "dashed",
        4 => "dotted",
        5 => "thick",
        6 => "double",
        7 => "hair",
        8 => "mediumDashed",
        9 => "dashDot",
        10 => "mediumDashDot",
        11 => "dashDotDot",
        12 => "mediumDashDotDot",
        13 => "slantDashDot",
        _ => "thin",
    };

    /// <summary>BrtColor（8 字节）→ "#RRGGBB"；RGB 固定位于颜色记录 +4（theme/rgb 均已含解析后 RGB）。
    /// 类型字节 0x00=auto、0x01=无颜色时返回 null。</summary>
    private static string? ColorRgb(byte[] d, int off)
    {
        if (off + 7 > d.Length) return null;
        byte type = d[off];
        if (type == 0x00 || type == 0x01) return null;
        return $"#{d[off + 4]:X2}{d[off + 5]:X2}{d[off + 6]:X2}";
    }

    /// <summary>
    /// 读取工作表的超级表：从 sheetN.bin 的 BrtTablePart(0x0295) 取 relId，
    /// 经 sheetN.bin.rels 定位 xl/tables/tableN.bin，解析 BrtBeginList/BrtBeginListCol/BrtTableStyleClient。
    /// </summary>
    private static void ReadTablesForSheet(ZipArchive zip, string sheetPath, byte[] sheetBin, SheetData sheet)
    {
        // 收集表部件 relId（BrtTablePart 位于 BrtBeginTableParts 段内）
        var tableRelIds = new List<string>();
        int pos = 0;
        while (pos < sheetBin.Length)
        {
            int rt = Biff12Records.ReadVarInt(sheetBin, ref pos);
            int cb = Biff12Records.ReadVarInt(sheetBin, ref pos);
            if (cb < 0 || pos + cb > sheetBin.Length) break;
            if (rt == BrtTablePart)
            {
                var d = sheetBin;
                int off = pos;
                var rid = Biff12Records.ReadWideString(d, ref off);
                if (!string.IsNullOrEmpty(rid)) tableRelIds.Add(rid);
            }
            pos += cb;
        }
        if (tableRelIds.Count == 0) return;

        // sheet rels：relId → table 部件路径
        var slash = sheetPath.LastIndexOf('/');
        var dir = slash < 0 ? "" : sheetPath.Substring(0, slash);
        var file = slash < 0 ? sheetPath : sheetPath.Substring(slash + 1);
        var relsPath = $"{dir}/_rels/{file}.rels";
        var relsEntry = zip.GetEntry(relsPath);
        if (relsEntry is null) return;

        var relMap = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            var rels = XElement.Load(relsEntry.Open());
            var relNs = rels.Name.Namespace;
            foreach (var rel in rels.Elements(relNs + "Relationship"))
            {
                var id = rel.Attribute("Id")?.Value;
                var type = rel.Attribute("Type")?.Value ?? "";
                var target = rel.Attribute("Target")?.Value;
                if (id is not null && target is not null && type.EndsWith("/table", StringComparison.Ordinal))
                    relMap[id] = target;
            }
        }
        catch { return; }

        foreach (var rid in tableRelIds)
        {
            if (!relMap.TryGetValue(rid, out var target)) continue;
            var entry = ResolveTableTarget(sheetPath, target);
            var te = zip.GetEntry(entry);
            if (te is null) continue;
            byte[] bin;
            using (var s = te.Open())
            using (var ms = new MemoryStream()) { s.CopyTo(ms); bin = ms.ToArray(); }
            try
            {
                var tbl = ParseTableBin(bin);
                if (tbl is not null)
                {
                    tbl.OriginEntry = entry;
                    sheet.Tables.Add(tbl);
                }
            }
            catch { /* 单个表失败不影响整表 */ }
        }
    }

    /// <summary>解析 tableN.bin（BIFF12）→ XlTable。</summary>
    private static XlTable? ParseTableBin(byte[] data)
    {
        XlTable? tbl = null;
        int pos = 0;
        while (pos < data.Length)
        {
            int rt = Biff12Records.ReadVarInt(data, ref pos);
            int cb = Biff12Records.ReadVarInt(data, ref pos);
            if (cb < 0 || pos + cb > data.Length) break;
            int off = pos;

            switch (rt)
            {
                case BrtBeginList:
                {
                    // rwFirst(4) rwLast(4) colFirst(4) colLast(4) lt(4) idList(4)
                    // crwHeader(4) crwTotals(4) flags(4) nDxf*6(4*6) dwConnID(4) stName stDisplayName ...
                    int rwF = Biff12Records.ReadS32(data, off); off += 4;
                    int rwL = Biff12Records.ReadS32(data, off); off += 4;
                    int colF = Biff12Records.ReadS32(data, off); off += 4;
                    int colL = Biff12Records.ReadS32(data, off); off += 4;
                    off += 4; // lt
                    off += 4; // idList
                    int crwHeader = Biff12Records.ReadS32(data, off); off += 4;
                    int crwTotals = Biff12Records.ReadS32(data, off); off += 4;
                    off += 4;  // flags
                    off += 4 * 6; // nDxf*
                    off += 4;  // dwConnID
                    var name = ReadNullableWide(data, ref off) ?? "";
                    off += 4; // stDisplayName cch 由 ReadNullableWide 已消费；此处不再解析
                    tbl = new XlTable
                    {
                        Name = name,
                        Ref = CellRefText(colF, rwF) + ":" + CellRefText(colL, rwL),
                        TotalsRowShown = crwTotals != 0,
                        AutoFilter = true,
                    };
                    _ = crwHeader;
                    break;
                }
                case BrtBeginListCol:
                {
                    // idField(4) ilta(4) nDxfHdr(4) nDxfInsertRow(4) nDxfAgg(4) idqsif(4) stName stCaption ...
                    off += 4;      // idField
                    off += 4 * 4;  // ilta + 3×nDxf
                    off += 4;      // idqsif
                    ReadNullableWide(data, ref off);            // stName
                    var cap = ReadNullableWide(data, ref off) ?? ""; // stCaption
                    if (tbl is not null && cap.Length > 0)
                        tbl.AddColumn(new XlTableColumn { Name = cap });
                    break;
                }
                case BrtTableStyleClient:
                {
                    if (data.Length - off >= 2)
                    {
                        int flags = data[off] | (data[off + 1] << 8);
                        off += 2;
                        var styleName = ReadNullableWide(data, ref off);
                        if (tbl is not null && !string.IsNullOrEmpty(styleName))
                        {
                            tbl.CustomStyleName = styleName;
                            if (styleName.StartsWith("TableStyle", StringComparison.Ordinal)
                                && Enum.TryParse<TableStyleStyle>(styleName.Substring("TableStyle".Length), out var parsed))
                                tbl.Style = parsed;
                            tbl.ShowFirstColumn = (flags & 0x01) != 0;
                            tbl.ShowLastColumn = (flags & 0x02) != 0;
                            tbl.ShowRowStripes = (flags & 0x04) != 0;
                            tbl.ShowColumnStripes = (flags & 0x08) != 0;
                        }
                    }
                    break;
                }
            }
            pos += cb;
        }
        return tbl;
    }

    /// <summary>解析 sheet rels 中的 table Target → 包内绝对路径（相对 xl/worksheets 或 xl/…）。</summary>
    private static string ResolveTableTarget(string sheetPath, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        if (target.StartsWith("../", StringComparison.Ordinal)) return "xl/" + target.Substring(3);
        var slash = sheetPath.LastIndexOf('/');
        var dir = slash < 0 ? "xl" : sheetPath.Substring(0, slash);
        return dir + "/" + target;
    }

    /// <summary>XLNullableWideString：cch(4)，0xFFFFFFFF 为 null。</summary>
    private static string? ReadNullableWide(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return null;
        uint cch = Biff12Records.ReadU32(d, off); off += 4;
        if (cch == 0xFFFFFFFF) return null;
        if (cch > 0x7FFF || off + (int)cch * 2 > d.Length) return null;
        var s = System.Text.Encoding.Unicode.GetString(d, off, (int)cch * 2);
        off += (int)cch * 2;
        return s;
    }

    /// <summary>0-based 列/行 → A1 单元格引用（用于表 Ref 构造）。</summary>
    private static string CellRefText(int col, int row) => RgceDecoder.ColumnName(col) + (row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>读取工作表 rels 中 hyperlink 关系（外部超链接 rId → Target）</summary>
    private static Dictionary<string, string> ReadSheetHyperlinkRels(ZipArchive zip, string sheetPath)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(sheetPath)) return result;
        var slash = sheetPath.LastIndexOf('/');
        var dir = slash < 0 ? "" : sheetPath.Substring(0, slash);
        var file = slash < 0 ? sheetPath : sheetPath.Substring(slash + 1);
        var relsPath = $"{dir}/_rels/{file}.rels";

        var relsEntry = zip.GetEntry(relsPath);
        if (relsEntry is null) return result;
        try
        {
            var rels = XElement.Load(relsEntry.Open());
            var relNs = rels.Name.Namespace;
            foreach (var rel in rels.Elements(relNs + "Relationship"))
            {
                var type = rel.Attribute("Type")?.Value ?? "";
                if (!type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
                var id = rel.Attribute("Id")?.Value;
                var target = rel.Attribute("Target")?.Value ?? "";
                if (id is not null && !string.IsNullOrEmpty(target)) result[id] = target;
            }
        }
        catch
        {
            result.Clear();
        }
        return result;
    }

    private static SheetData ParseWorksheet(byte[] data, string sheetName, List<string> sst,
        Dictionary<int, string> formats, List<int> cellXfs, List<CellStyle?> cellStyles, bool date1904, Dictionary<string, string>? hlinkRels = null)
    {
        var records = Biff12Records.ReadAll(data);
        var sheet = new SheetData { SheetName = sheetName };
        var cells = new Dictionary<int, Dictionary<int, Cell>>();
        int maxRow = -1;
        int maxCol = -1;
        var colWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();
        int freezeRows = 0;
        int freezeCols = 0;
        int currentRow = -1;
        int prevCol = -1;

        foreach (var rec in records)
        {
            var d = rec.Data;
            bool isShort = rec.Rt >= BrtShortBlank;
            switch (rec.Rt)
            {
                case BrtRowHdr:
                    currentRow = ParseRowHdr(d, rowHeights);
                    prevCol = -1;
                    break;
                case BrtCellBlank:
                case BrtShortBlank:
                    PutCell(cells, d, isShort, ref prevCol, currentRow, (_) => Cell.Empty, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellRk:
                case BrtShortRk:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => FormatDetector.CellFromNumber(BiffShared.DecodeRk(ReadS32(d, valOff)), StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904),
                        ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellError:
                case BrtShortError:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromText(BiffShared.ErrorCode(d[valOff])), ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellBool:
                case BrtShortBool:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromBoolean(d[valOff] != 0), ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellReal:
                case BrtShortReal:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => FormatDetector.CellFromNumber(BitConverter.ToDouble(d, valOff), StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904),
                        ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellSt:
                case BrtShortSt:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromText(ReadStringAt(d, valOff)), ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtCellIsst:
                case BrtShortIsst:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            int idx = ReadS32(d, valOff);
                            return idx >= 0 && idx < sst.Count ? Cell.FromText(sst[idx]) : Cell.Empty;
                        }, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtFmlaString:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromText(ReadStringAt(d, valOff));
                            ApplyFormula(d, valOff + 4 + ReadWideLen(d, valOff), cell);
                            return cell;
                        }, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtFmlaNum:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = FormatDetector.CellFromNumber(BitConverter.ToDouble(d, valOff), StyleRef(d, 4), cellXfs, formats, date1904);
                            ApplyFormula(d, valOff + 8, cell);
                            return cell;
                        }, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtFmlaBool:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromBoolean(d[valOff] != 0);
                            ApplyFormula(d, valOff + 1, cell);
                            return cell;
                        }, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtFmlaError:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromText(BiffShared.ErrorCode(d[valOff]));
                            ApplyFormula(d, valOff + 1, cell);
                            return cell;
                        }, ref maxRow, ref maxCol, cellStyles);
                    break;
                case BrtColInfo:
                    ParseColInfo(d, colWidths);
                    break;
                case BrtMergeCell:
                    ParseMergeCell(d, sheet);
                    break;
                case BrtHLink:
                    ParseHLink(d, cells, hlinkRels, ref maxRow, ref maxCol);
                    break;
                case BrtPane:
                    // colFrozen(Xnum 8) + rowFrozen(Xnum 8)
                    if (d.Length >= 16)
                    {
                        double colFrozen = BitConverter.ToDouble(d, 0);
                        double rowFrozen = BitConverter.ToDouble(d, 8);
                        if (colFrozen >= 1.0 || rowFrozen >= 1.0)
                        {
                            freezeRows = (int)Math.Round(rowFrozen);
                            freezeCols = (int)Math.Round(colFrozen);
                        }
                    }
                    break;
                case BrtBeginAFilter:
                    ParseBeginAFilter(d, sheet);
                    break;
                case BrtDVal:
                    ParseDVal(d, sheet);
                    break;
            }
        }

        // 组装行
        for (int row = 0; row <= maxRow; row++)
        {
            var arr = new Cell[maxCol + 1];
            for (int c = 0; c <= maxCol; c++) arr[c] = Cell.Empty;
            if (cells.TryGetValue(row, out var rowCells))
            {
                foreach (var kv in rowCells)
                    arr[kv.Key] = kv.Value;
            }
            sheet.Rows.Add(arr);
        }

        if (colWidths.Count > 0)
        {
            var widths = new List<double>(maxCol + 1);
            for (int c = 0; c <= maxCol; c++)
                widths.Add(colWidths.TryGetValue(c, out var w) ? w : 8.43);
            sheet.ColumnWidths = widths;
        }

        if (rowHeights.Count > 0)
            sheet.RowHeights = rowHeights;

        sheet.FreezeHeader = freezeRows == 1 && freezeCols == 0;
        sheet.FreezeRows = freezeRows;
        sheet.FreezeColumns = freezeCols;
        return sheet;
    }

    /// <summary>解析行头，返回行号；flags 含 0x20 时 miyRw 表示显式行高（缇）。</summary>
    private static int ParseRowHdr(byte[] d, Dictionary<int, double> rowHeights)
    {
        if (d.Length < 12) return 0;
        int rw = ReadS32(d, 0);
        int miyRw = Biff12Records.ReadU16(d, 8);
        byte flags = d[11];
        if ((flags & 0x20) != 0 && miyRw != 0 && miyRw != 0xFF)
            rowHeights[rw] = miyRw / 20.0;
        return rw;
    }

    private static void ParseHLink(byte[] d, Dictionary<int, Dictionary<int, Cell>> cells,
        Dictionary<string, string>? hlinkRels, ref int maxRow, ref int maxCol)
    {
        // BrtHLink = RfX(16) + relId(XLNullableWideString) + location + tooltip + display
        if (d.Length < 16) return;
        int off = 0;
        int rwFirst = ReadS32(d, off); off += 4;
        int rwLast = ReadS32(d, off); off += 4;
        int colFirst = ReadS32(d, off); off += 4;
        int colLast = ReadS32(d, off); off += 4;

        var relId = ReadNullableWideString(d, ref off);
        var location = Biff12Records.ReadWideString(d, ref off);
        var tooltip = Biff12Records.ReadWideString(d, ref off);
        Biff12Records.ReadWideString(d, ref off); // display（忽略）

        string target;
        bool isInternal;
        if (!string.IsNullOrEmpty(location))
        {
            target = location.StartsWith("#", StringComparison.Ordinal) ? location : "#" + location;
            isInternal = true;
        }
        else if (hlinkRels is not null && !string.IsNullOrEmpty(relId) && hlinkRels.TryGetValue(relId, out var t))
        {
            target = t;
            isInternal = false;
        }
        else
        {
            return;
        }

        for (int r = rwFirst; r <= rwLast; r++)
        {
            for (int c = colFirst; c <= colLast; c++)
            {
                if (!cells.TryGetValue(r, out var rowCells))
                {
                    rowCells = new Dictionary<int, Cell>();
                    cells[r] = rowCells;
                }
                if (!rowCells.TryGetValue(c, out var cell))
                {
                    cell = Cell.Empty;
                    rowCells[c] = cell;
                }
                cell.Hyperlink = new Hyperlink { Target = target, Tooltip = string.IsNullOrEmpty(tooltip) ? null : tooltip, IsInternal = isInternal };
                if (r > maxRow) maxRow = r;
                if (c > maxCol) maxCol = c;
            }
        }
    }

    private static string ReadNullableWideString(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return "";
        uint cch = Biff12Records.ReadU32(d, off);
        off += 4;
        if (cch == 0 || cch == 0xFFFFFFFF) return "";
        int bytes = (int)cch * 2;
        if (off + bytes > d.Length) return "";
        var s = System.Text.Encoding.Unicode.GetString(d, off, bytes);
        off += bytes;
        return s;
    }

    private static void PutCell(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d, bool shortCell,
        ref int prevCol, int currentRow, Func<int, Cell> factory, ref int maxRow, ref int maxCol,
        List<CellStyle?>? cellStyles = null)
    {
        int valueOff = shortCell ? 4 : 8;
        int col;
        if (shortCell)
        {
            col = prevCol + 1;
        }
        else
        {
            col = ReadS32(d, 0);
            if (col < 0) col = 0;
        }
        prevCol = col;

        var cell = factory(valueOff);
        // 应用单元格样式（字体/填充/边框）：ixfe → CellStyle。默认样式返回 null，保持轻量。
        if (cellStyles is not null)
        {
            int ixfe = StyleRef(d, shortCell ? 0 : 4);
            if (ixfe >= 0 && ixfe < cellStyles.Count && cellStyles[ixfe] is { } st)
                cell.Style = st;
        }
        if (!cells.TryGetValue(currentRow, out var rowCells))
        {
            rowCells = new Dictionary<int, Cell>();
            cells[currentRow] = rowCells;
        }
        rowCells[col] = cell;
        if (currentRow > maxRow) maxRow = currentRow;
        if (col > maxCol) maxCol = col;
    }

    private static int StyleRef(byte[] d, int off)
    {
        if (off + 3 > d.Length) return 0;
        return d[off] | (d[off + 1] << 8) | (d[off + 2] << 16);
    }

    private static string ReadStringAt(byte[] d, int off)
    {
        int o = off;
        return Biff12Records.ReadWideString(d, ref o);
    }

    /// <summary>XLWideString 的字节长度（cch(4) + 字符数据）</summary>
    private static int ReadWideLen(byte[] d, int off)
    {
        if (off + 4 > d.Length) return 0;
        uint cch = Biff12Records.ReadU32(d, off);
        return 4 + (int)cch * 2;
    }

    /// <summary>xlsb 公式记录尾部解析：value 之后 2 字节跳过 + cce(4) + RPN。</summary>
    private static void ApplyFormula(byte[] d, int valueEnd, Cell cell)
    {
        int fOff = valueEnd + 2;
        if (fOff + 4 > d.Length) return;
        int cce = ReadS32(d, fOff);
        if (cce <= 0 || fOff + 4 + cce > d.Length) return;
        var rpn = new byte[cce];
        Array.Copy(d, fOff + 4, rpn, 0, cce);
        var text = Biff.FormulaParser.Parse(rpn, biff12: true);
        if (!string.IsNullOrEmpty(text))
        {
            // 将公式文本存入 Formula，不覆盖缓存值。
            cell.IsFormula = true;
            cell.Formula = text;
        }
    }

    private static void ParseColInfo(byte[] d, Dictionary<int, double> colWidths)
    {
        if (d.Length < 12) return;
        int colFirst = ReadS32(d, 0);
        int colLast = ReadS32(d, 4);
        uint width = Biff12Records.ReadU32(d, 8);
        double w = width / 256.0;
        for (int c = colFirst; c <= colLast; c++)
            colWidths[c] = w;
    }

    private static void ParseMergeCell(byte[] d, SheetData sheet)
    {
        if (d.Length < 16) return;
        int rwFirst = ReadS32(d, 0);
        int rwLast = ReadS32(d, 4);
        int colFirst = ReadS32(d, 8);
        int colLast = ReadS32(d, 12);
        sheet.MergedRanges.Add(new CellRange(rwFirst, rwLast, colFirst, colLast));
    }

    /// <summary>BrtBeginAFilter：rfx = rwFirst(4) + rwLast(4) + colFirst(4) + colLast(4) = 16 字节。</summary>
    private static void ParseBeginAFilter(byte[] d, SheetData sheet)
    {
        if (d.Length < 16) return;
        int rwFirst = ReadS32(d, 0);
        int rwLast = ReadS32(d, 4);
        int colFirst = ReadS32(d, 8);
        int colLast = ReadS32(d, 12);
        sheet.Filter = new AutoFilter
        {
            Range = CellRef.ToString(rwFirst, colFirst) + ":" + CellRef.ToString(rwLast, colLast),
        };
    }

    private static int ReadS32(byte[] d, int off) => Biff12Records.ReadS32(d, off);

    /// <summary>数组切片（net48 无 Range 语法）。</summary>
    private static byte[] Slice(byte[] d, int off, int len)
    {
        var r = new byte[len];
        Array.Copy(d, off, r, 0, len);
        return r;
    }

    /// <summary>
    /// BrtDVal：flags(4) + cRefs(4) + cRefs×(rwFirst,rwLast,colFirst,colLast)(各4)
    /// + promptTitle + prompt + errorTitle + errorMessage（XLNullableWideString）
    /// + cce(4) + rgce[formula1] + (可选)cce(4) + rgce[formula2] + reserved(4)。
    /// flags 位 0-3 = 类型：1=whole 2=decimal 3=list 4=date 5=time 6=textLength 7=custom。
    /// </summary>
    private static void ParseDVal(byte[] d, SheetData sheet)
    {
        if (d.Length < 8) return;
        int o = 0;
        int flags = ReadS32(d, o); o += 4;
        int typeCode = flags & 0xF;
        int cRefs = ReadS32(d, o); o += 4;
        if (cRefs < 0 || o + cRefs * 16 > d.Length) return;

        var refs = new List<string>(cRefs);
        for (int i = 0; i < cRefs; i++)
        {
            int rwF = ReadS32(d, o); int rwL = ReadS32(d, o + 4);
            int colF = ReadS32(d, o + 8); int colL = ReadS32(d, o + 12);
            o += 16;
            refs.Add(CellRef.ToString(rwF, colF) + ":" + CellRef.ToString(rwL, colL));
        }
        if (o + 16 > d.Length) return;
        var promptTitle = ReadNullableWide(d, ref o);
        var prompt = ReadNullableWide(d, ref o);
        var errorTitle = ReadNullableWide(d, ref o);
        var errorMsg = ReadNullableWide(d, ref o);

        string? f1 = null, f2 = null;
        if (o + 4 <= d.Length)
        {
            int cce1 = ReadS32(d, o); o += 4;
            if (cce1 >= 0 && o + cce1 <= d.Length)
            {
                f1 = RgceDecoder.Decode(Slice(d, o, cce1), Array.Empty<string>()) ?? DecodeDvLiteral(d, o, cce1);
                o += cce1;
            }
        }
        o += 4; // reserved（rgce1 与 rgce2 之间）
        if (o + 4 <= d.Length)
        {
            int cce2 = ReadS32(d, o); o += 4;
            if (cce2 > 0 && o + cce2 <= d.Length)
                f2 = RgceDecoder.Decode(Slice(d, o, cce2), Array.Empty<string>()) ?? DecodeDvLiteral(d, o, cce2);
        }

        var dv = new DataValidation
        {
            Type = typeCode switch
            {
                1 => DataValidationType.WholeNumber,
                2 => DataValidationType.Decimal,
                4 => DataValidationType.Date,
                3 => DataValidationType.List,
                _ => DataValidationType.List, // 6=textLength 等无对应模型，按列表占位（值仍可读）
            },
            Sqref = string.Join(" ", refs),
            Formula1 = f1 ?? "",
            Formula2 = f2,
            AllowBlank = (flags & (1 << 8)) != 0,
            PromptTitle = promptTitle ?? errorTitle,
            Prompt = prompt ?? errorMsg,
        };
        sheet.Validations ??= new List<DataValidation>();
        sheet.Validations.Add(dv);
    }

    /// <summary>列表验证的公式是字符串常量（PtgStr 0x17），RgceDecoder 只支持单令牌，此处直接解字符串。</summary>
    private static string? DecodeDvLiteral(byte[] d, int off, int cce)
    {
        if (cce >= 3 && d[off] == 0x17)
        {
            int cch = d[off + 1] | (d[off + 2] << 8);
            if (off + 3 + cch * 2 <= d.Length)
                return "\"" + System.Text.Encoding.Unicode.GetString(d, off + 3, cch * 2) + "\"";
        }
        return null;
    }

    private const int BrtBeginComments = 0x0274;
    private const int BrtEndComments = 0x0275;
    private const int BrtBeginCommentAuthors = 0x0276;
    private const int BrtEndCommentAuthors = 0x0277;
    private const int BrtCommentAuthor = 0x0278;
    private const int BrtBeginCommentList = 0x0279;
    private const int BrtCommentText = 0x027D;

    /// <summary>
    /// 从 commentsN.bin 和 VML 读取批注。
    /// commentsN.bin 是独立 BIFF12 部件（不在 sheetN.bin 内），包含作者表和批注文本。
    /// 单元格位置从 VML 的 <x:Row><x:Column> 获取（与 XLSX 相同的 VML 格式）。
    /// </summary>
    private static void ReadCommentsForSheet(ZipArchive zip, string sheetPath, SheetData sheet)
    {
        var (dir, file) = SplitSheetPath(sheetPath);
        var relsPath = $"{dir}/_rels/{file}.rels";
        var relsEntry = zip.GetEntry(relsPath);
        if (relsEntry is null) return;

        string? commentsTarget = null;
        string? vmlTarget = null;
        try
        {
            var rels = XElement.Load(relsEntry.Open());
            var relNs = rels.Name.Namespace;
            foreach (var rel in rels.Elements(relNs + "Relationship"))
            {
                var type = rel.Attribute("Type")?.Value ?? "";
                if (type.EndsWith("/comments", StringComparison.OrdinalIgnoreCase))
                    commentsTarget = rel.Attribute("Target")?.Value;
                else if (type.EndsWith("/vmlDrawing", StringComparison.OrdinalIgnoreCase))
                    vmlTarget = rel.Attribute("Target")?.Value;
            }
        }
        catch { return; }

        if (commentsTarget is null) return;

        var commentsPath = ResolveRelativePath(sheetPath, commentsTarget);
        var commentsEntry = zip.GetEntry(commentsPath);
        if (commentsEntry is null) return;

        var commentsData = ReadEntry(zip, commentsPath);
        if (commentsData is null) return;

        var texts = ParseCommentsBin(commentsData);

        if (vmlTarget is not null)
        {
            var vmlPath = ResolveRelativePath(sheetPath, vmlTarget);
            var vmlEntry = zip.GetEntry(vmlPath);
            if (vmlEntry is not null)
            {
                var vmlBytes = ReadEntry(zip, vmlPath);
                if (vmlBytes is not null)
                {
                    var vmlText = System.Text.Encoding.UTF8.GetString(vmlBytes);
                    var positions = ParseVmlPositions(vmlText);
                    for (int i = 0; i < texts.Count && i < positions.Count; i++)
                    {
                        var (row, col) = positions[i];
                        var a1Ref = CellRef.ToString(row, col);
                        sheet.Comments ??= new Dictionary<string, string>();
                        sheet.Comments[a1Ref] = texts[i];
                    }
                }
            }
        }
    }

    /// <summary>解析 comments1.bin 的 BIFF12 记录，提取批注文本列表（按出现顺序）。</summary>
    private static List<string> ParseCommentsBin(byte[] data)
    {
        var texts = new List<string>();
        var records = Biff12Records.ReadAll(data);
        foreach (var rec in records)
        {
            if (rec.Rt == BrtCommentText && rec.Data.Length >= 1)
            {
                int off = 1;
                var text = Biff12Records.ReadWideString(rec.Data, ref off);
                texts.Add(text);
            }
        }
        return texts;
    }

    /// <summary>解析 VML 中 <x:Row> 和 <x:Column> 的值，返回 (row, col) 列表（0-based）。</summary>
    private static List<(int row, int col)> ParseVmlPositions(string vmlText)
    {
        var positions = new List<(int row, int col)>();
        try
        {
            var doc = XElement.Parse(vmlText);
            var ns = doc.Name.Namespace;
            var xNs = ns.GetName("x");
            foreach (var shape in doc.Descendants())
            {
                if (!shape.Name.LocalName.Equals("ClientData", StringComparison.OrdinalIgnoreCase))
                    continue;
                int row = -1, col = -1;
                foreach (var child in shape.Elements())
                {
                    if (child.Name.LocalName == "Row" && int.TryParse(child.Value, out var r))
                        row = r;
                    if (child.Name.LocalName == "Column" && int.TryParse(child.Value, out var c))
                        col = c;
                }
                if (row >= 0 && col >= 0)
                    positions.Add((row, col));
            }
        }
        catch { }
        return positions;
    }

    private static (string dir, string file) SplitSheetPath(string sheetPath)
    {
        var slash = sheetPath.LastIndexOf('/');
        if (slash < 0) return ("", sheetPath);
        return (sheetPath.Substring(0, slash), sheetPath.Substring(slash + 1));
    }

    private static string ResolveRelativePath(string basePath, string relative)
    {
        if (relative.StartsWith("/", StringComparison.Ordinal))
            return relative.TrimStart('/');
        var baseDir = System.IO.Path.GetDirectoryName(basePath)?.Replace('\\', '/') ?? "";
        while (relative.StartsWith("../", StringComparison.Ordinal))
        {
            relative = relative.Substring(3);
            var lastSlash = baseDir.LastIndexOf('/');
            if (lastSlash >= 0) baseDir = baseDir.Substring(0, lastSlash);
        }
        return string.IsNullOrEmpty(baseDir) ? relative : baseDir + "/" + relative;
    }
}
