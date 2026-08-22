using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using LiteExcel.Internal.Biff12;

namespace LiteExcel.Internal;

/// <summary>
/// .xlsb（BIFF12 二进制 OOXML 变体）写入后端。
/// 容器仍是 ZIP（与 xlsx 相同的 OPC 包），部件内为二进制记录流。
/// 记录头 = RecordType(LEB128) + RecordSize(LEB128)，与读取侧一致。
/// 支持：多工作表（中文名）、文本/数字/日期/布尔、共享字符串表、
/// 样式与数字格式、合并单元格、列宽、行高、冻结表头、公式缓存值。
/// 公式文本不保留（按缓存值写出）；图片/图表等高级能力不在范围内。
/// </summary>
internal static class XlsbWriter
{
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    // workbook.bin
    private const int BrtBeginBook = 0x0083;
    private const int BrtFileVersion = 0x0080;
    private const int BrtWbProp = 0x0099;
    private const int BrtFileSharingIso = 0x02A4;
    private const int BrtBeginBookViews = 0x0087;
    private const int BrtBookView = 0x009E;
    private const int BrtEndBookViews = 0x0088;
    private const int BrtBeginBundleShs = 0x008F;
    private const int BrtBundleSh = 0x009C;
    private const int BrtEndBundleShs = 0x0090;
    private const int BrtEndBook = 0x0084;

    // sharedStrings.bin
    private const int BrtBeginSst = 0x009F;
    private const int BrtSSTItem = 0x0013;
    private const int BrtEndSst = 0x00A0;

    // styles.bin
    private const int BrtBeginStyleSheet = 0x0116;
    private const int BrtBeginFmts = 0x0267;
    private const int BrtFmt = 0x002C;
    private const int BrtEndFmts = 0x0268;
    private const int BrtBeginFonts = 0x0263;
    private const int BrtFont = 0x002B;
    private const int BrtEndFonts = 0x0264;
    private const int BrtBeginFills = 0x025B;
    private const int BrtFill = 0x002D;
    private const int BrtEndFills = 0x025C;
    private const int BrtBeginBorders = 0x0265;
    private const int BrtBorder = 0x002E;
    private const int BrtEndBorders = 0x0266;
    private const int BrtBeginCellStyleXFs = 0x0272;
    private const int BrtXF = 0x002F;
    private const int BrtEndCellStyleXFs = 0x0273;
    private const int BrtBeginCellXFs = 0x0269;
    private const int BrtEndCellXFs = 0x026A;
    private const int BrtBeginStyles = 0x026B;
    private const int BrtStyle = 0x0030;
    private const int BrtEndStyles = 0x026C;
    private const int BrtBeginDXFs = 0x01F9;
    private const int BrtEndDXFs = 0x01FA;
    private const int BrtBeginTableStyles = 0x01FC;
    private const int BrtEndTableStyles = 0x01FD;
    private const int BrtEndStyleSheet = 0x0117;

    // worksheet.bin
    private const int BrtBeginSheet = 0x0081;
    private const int BrtWsProp = 0x0093;
    private const int BrtWsDim = 0x0094;
    private const int BrtBeginWsViews = 0x0085;
    private const int BrtBeginWsView = 0x0089;
    private const int BrtPane = 0x0097;
    private const int BrtEndWsView = 0x008A;
    private const int BrtEndWsViews = 0x0086;
    private const int BrtBeginColInfos = 0x0186;
    private const int BrtColInfo = 0x003C;
    private const int BrtEndColInfos = 0x0187;
    private const int BrtBeginSheetData = 0x0091;
    private const int BrtRowHdr = 0x0000;
    private const int BrtCellBlank = 0x0001;
    private const int BrtCellRk = 0x0002;
    private const int BrtCellBool = 0x0004;
    private const int BrtCellReal = 0x0005;
    private const int BrtCellSt = 0x0006;
    private const int BrtCellIsst = 0x0007;
    private const int BrtShortBlank = 0x000C;
    private const int BrtShortRk = 0x000D;
    private const int BrtShortBool = 0x000F;
    private const int BrtShortReal = 0x0010;
    private const int BrtShortSt = 0x0011;
    private const int BrtShortIsst = 0x0012;
    private const int BrtEndSheetData = 0x0092;
    private const int BrtBeginMergeCells = 0x00B1;
    private const int BrtMergeCell = 0x00B0;
    private const int BrtEndMergeCells = 0x00B2;
    private const int BrtHLink = 0x01EE;
    private const int BrtEndSheet = 0x0082;

    private const int DefaultFontId = 0;
    private const int DefaultFillId = 0;
    private const int DefaultBorderId = 0;
    private const int DefaultCellStyleXf = 0;
    private const int DefaultCellXf = 0;
    private const int BuiltinDateFmtId = 14;
    private const int FirstCustomFmtId = 164;

    /// <summary>写出 .xlsb 工作簿到流。vbaProject 为源工作簿捕获的宏工程字节（可为 null）；workbookCodeName 为宿主代码名（可为 null）。
    /// <paramref name="preserved"/> 为打开时捕获的未重建 OOXML 部件（图表/透视表/主题/绘图等），保存时透传；
    /// <paramref name="properties"/> 为文档属性，非 null 时写出 docProps。</summary>
    public static void Write(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject = null, string? workbookCodeName = null, bool date1904 = false,
        string? fileSharingHash = null, string? fileSharingSalt = null, int? fileSharingSpin = null, bool fileSharingReadOnlyRecommended = false,
        Action<DegradationInfo>? onDegradation = null, ExcelFormat targetFormat = ExcelFormat.Xlsb,
        OoxmlPreservedParts? preserved = null, WorkbookProperties? properties = null)
    {
        if (sheets is null || sheets.Count == 0)
            throw new ArgumentException("至少需要一张工作表", nameof(sheets));

        ReportSheetDegradations(sheets, onDegradation, targetFormat);

        var sst = new List<string>();
        var sstIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var cellXfs = new List<(int Ifmt, string? FmtCode)>();
        var fmtCodeToXf = new Dictionary<string, int>(StringComparer.Ordinal);
        cellXfs.Add((0, null)); // 索引 0 = General 默认样式

        int GetXf(string? fmtCode)
        {
            if (string.IsNullOrEmpty(fmtCode)) return DefaultCellXf;
            if (fmtCodeToXf.TryGetValue(fmtCode, out var idx)) return idx;
            idx = cellXfs.Count;
            fmtCodeToXf[fmtCode] = idx;
            cellXfs.Add((ResolveFmtId(fmtCode), fmtCode));
            return idx;
        }

        void ScanCell(Cell cell)
        {
            if (cell.IsEmpty) return;
            switch (cell.Type)
            {
                case CellType.Text:
                    if (cell.Text is not null && !sstIndex.ContainsKey(cell.Text))
                    {
                        sstIndex[cell.Text] = sst.Count;
                        sst.Add(cell.Text);
                    }
                    break;
                case CellType.Number:
                case CellType.Date:
                    GetXf(cell.NumberFormat);
                    break;
            }
        }

        foreach (var sheet in sheets)
            foreach (var row in sheet.Rows)
                foreach (var cell in row)
                    ScanCell(cell);

        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        // P0-14(xlsb): 先写保留部件（blob），写入器重建的条目在 rebuilt 集合中，避免重名
        if (preserved is not null)
        {
            var rebuilt = OoxmlPreservedParts.BuildRebuiltEntries(sheets.Count, binary: true);
            foreach (var kv in preserved.Parts)
            {
                if (rebuilt.Contains(kv.Key)) continue;
                WriteEntry(zip, kv.Key, kv.Value);
            }
        }

        // 包结构
        WriteEntry(zip, "[Content_Types].xml", ContentTypesXml(sheets.Count, sst.Count > 0, vbaProject is not null, properties is not null, preserved));
        WriteEntry(zip, "_rels/.rels", RootRelsXml(properties is not null));
        WriteEntry(zip, "xl/workbook.bin", BuildWorkbookBin(sheets, workbookCodeName, date1904, fileSharingHash, fileSharingSalt, fileSharingSpin, fileSharingReadOnlyRecommended));
        WriteEntry(zip, "xl/_rels/workbook.bin.rels", WorkbookRelsXml(sheets.Count, sst.Count > 0, vbaProject is not null, preserved));
        WriteEntry(zip, "xl/styles.bin", BuildStylesBin(cellXfs));
        if (vbaProject is not null && vbaProject.Length > 0)
            WriteEntry(zip, "xl/vbaProject.bin", vbaProject);
        if (sst.Count > 0)
            WriteEntry(zip, "xl/sharedStrings.bin", BuildSharedStringsBin(sst, sstIndex));
        if (properties is not null)
        {
            WriteEntry(zip, "docProps/core.xml", XlsxWriter.CorePropsXml(properties));
            WriteEntry(zip, "docProps/app.xml", XlsxWriter.AppPropsXml(properties, sheets));
        }

        for (int i = 0; i < sheets.Count; i++)
        {
            var extLinks = CollectExternalHyperlinks(sheets[i]);
            WriteEntry(zip, $"xl/worksheets/sheet{i + 1}.bin", BuildWorksheetBin(sheets[i], sstIndex, GetXf, date1904, extLinks));
            var sheetRels = BuildSheetRelsXml(i + 1, extLinks, preserved);
            if (!string.IsNullOrEmpty(sheetRels))
                WriteEntry(zip, $"xl/worksheets/_rels/sheet{i + 1}.bin.rels", sheetRels);
        }
    }

    /// <summary>合并工作表级保留 rels（图表/透视表等）与重建的超链接 rels </summary>
    private static string? BuildSheetRelsXml(int sheetNumber, List<string> extLinks, OoxmlPreservedParts? preserved)
    {
        var relParts = new List<XlsxWriter.RelInfo>();
        for (int k = 0; k < extLinks.Count; k++)
        {
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rIdH{k + 1}",
                Type = $"{OfficeRelNs}/hyperlink",
                Target = extLinks[k],
                TargetMode = "External",
            });
        }
        string rebuilt = XlsxWriter.RelsXml(relParts);
        string original = "";
        if (preserved is not null && preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetNumber}.bin.rels", out var r))
            original = r;
        return XlsxWriter.MergeRelsXml(original, "xl/worksheets", new HashSet<string>(StringComparer.Ordinal), rebuilt);
    }

    /// <summary>写出 xlsb 时对静默丢弃的 sheet 级能力逐项上报（P0-4/15/16 显式化） </summary>
    private static void ReportSheetDegradations(IReadOnlyList<SheetData> sheets,
        Action<DegradationInfo>? onDegradation, ExcelFormat targetFormat)
    {
        if (onDegradation is null) return;
        foreach (var sheet in sheets)
        {
            void Report(DegradationCapability cap, string msg)
                => onDegradation(new DegradationInfo
                {
                    Capability = cap,
                    SheetName = sheet.SheetName,
                    TargetFormat = targetFormat,
                    Message = msg,
                });
            if (sheet.Comments is { Count: > 0 })
                Report(DegradationCapability.Comments, $"xlsb 不支持批注，工作表 '{sheet.SheetName}' 的批注已丢弃。");
            if (sheet.Validations is { Count: > 0 })
                Report(DegradationCapability.DataValidation, $"xlsb 不支持数据验证，工作表 '{sheet.SheetName}' 的数据验证已丢弃。");
            if (sheet.Filter is not null)
                Report(DegradationCapability.AutoFilter, $"xlsb 不支持自动筛选，工作表 '{sheet.SheetName}' 的筛选已丢弃。");
            if (sheet.Images is { Count: > 0 })
                Report(DegradationCapability.Images, $"xlsb 不支持图片，工作表 '{sheet.SheetName}' 的图片已丢弃。");
            if (DegradationDetector.HasNonNumberFormatStyles(sheet))
                Report(DegradationCapability.Styles, $"xlsb 仅支持数字格式，工作表 '{sheet.SheetName}' 的完整样式（字体/颜色/边框/对齐/换行）已降级。");
            if (sheet.ConditionalFormats is { Count: > 0 })
                Report(DegradationCapability.ConditionalFormatting, $"xlsb 不支持条件格式，工作表 '{sheet.SheetName}' 的条件格式已丢弃。");
        }
    }

    /// <summary>收集工作表中外部超链接（按行序遍历，保持稳定顺序） </summary>
    private static List<string> CollectExternalHyperlinks(SheetData sheet)
    {
        var list = new List<string>();
        foreach (var row in sheet.Rows)
        {
            foreach (var cell in row)
            {
                if (cell.Hyperlink is null || cell.Hyperlink.IsInternal) continue;
                if (string.IsNullOrEmpty(cell.Hyperlink.Target)) continue;
                if (!list.Contains(cell.Hyperlink.Target)) list.Add(cell.Hyperlink.Target);
            }
        }
        return list;
    }

    private static string XmlEscape(string s)
    {
        if (string.IsNullOrEmpty(s)) return s ?? "";
        return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }

    // ── 包 XML 部件 ──

    private static string ContentTypesXml(int sheetCount, bool hasSst, bool hasVba, bool hasProps, OoxmlPreservedParts? preserved)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        for (int i = 1; i <= sheetCount; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
        sb.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
        if (hasSst)
            sb.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
        if (hasVba)
            sb.Append("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
        if (hasProps)
        {
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }
        // P0-14(xlsb): 合并保留的 content types 声明（图表/透视表/主题等）
        if (preserved is not null)
        {
            foreach (var (ext, ct) in preserved.DefaultTypes)
                if (ext != "bin" && ext != "rels" && ext != "xml" && !sb.ToString().Contains($"Extension=\"{ext}\""))
                    sb.Append($"<Default Extension=\"{ext}\" ContentType=\"{ct}\"/>");
            foreach (var (part, ct) in preserved.OverrideTypes)
                if (!part.StartsWith("/xl/worksheets/") && !part.StartsWith("/xl/workbook") && !part.StartsWith("/xl/styles")
                    && !part.StartsWith("/xl/sharedStrings") && part != "/xl/vbaProject.bin"
                    && !sb.ToString().Contains($"PartName=\"{part}\""))
                    sb.Append($"<Override PartName=\"{part}\" ContentType=\"{ct}\"/>");
        }
        sb.Append("</Types>");
        return sb.ToString();
    }

    private static string RootRelsXml(bool hasProps)
    {
        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>");
        if (hasProps)
        {
            sb.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
            sb.Append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    private static string WorkbookRelsXml(int sheetCount, bool hasSst, bool hasVba, OoxmlPreservedParts? preserved)
    {
        var relParts = new List<XlsxWriter.RelInfo>();
        for (int i = 1; i <= sheetCount; i++)
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rId{i}",
                Type = $"{OfficeRelNs}/worksheet",
                Target = $"worksheets/sheet{i}.bin",
            });
        relParts.Add(new XlsxWriter.RelInfo
        {
            Id = $"rId{sheetCount + 1}",
            Type = $"{OfficeRelNs}/styles",
            Target = "styles.bin",
        });
        if (hasSst)
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rId{sheetCount + 2}",
                Type = $"{OfficeRelNs}/sharedStrings",
                Target = "sharedStrings.bin",
            });
        if (hasVba)
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rId{sheetCount + 3}",
                Type = "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                Target = "vbaProject.bin",
            });
        string rebuilt = XlsxWriter.RelsXml(relParts);
        string original = "";
        if (preserved is not null && preserved.Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var r))
            original = r;
        var rebuiltTargets = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 1; i <= sheetCount; i++)
        {
            rebuiltTargets.Add($"xl/worksheets/sheet{i}.bin");
            rebuiltTargets.Add($"xl/worksheets/_rels/sheet{i}.bin.rels");
        }
        rebuiltTargets.Add("xl/styles.bin");
        rebuiltTargets.Add("xl/sharedStrings.bin");
        rebuiltTargets.Add("xl/vbaProject.bin");
        rebuiltTargets.Add("xl/workbook.bin");
        return XlsxWriter.MergeRelsXml(original, "xl", rebuiltTargets, rebuilt) ?? rebuilt;
    }

    private static void WriteEntry(ZipArchive zip, string name, byte[] data)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var s = entry.Open();
        s.Write(data, 0, data.Length);
    }

    private static void WriteEntry(ZipArchive zip, string name, string xml)
    {
        WriteEntry(zip, name, new UTF8Encoding(false).GetBytes(xml));
    }

    // ── workbook.bin ──

    private static byte[] BuildWorkbookBin(IReadOnlyList<SheetData> sheets, string? workbookCodeName, bool date1904,
        string? fileSharingHash = null, string? fileSharingSalt = null, int? fileSharingSpin = null, bool fileSharingReadOnlyRecommended = false)
    {
        var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginBook, Array.Empty<byte>());
        WriteRecord(ms, BrtFileVersion, FileVersion());
        WriteRecord(ms, BrtWbProp, WbProp(workbookCodeName, date1904)); // 日期系统由 flags bit0 指定
        // 写保护（修改密码）：BrtFileSharingIso（对齐 Excel 样本布局）
        if (!string.IsNullOrEmpty(fileSharingHash))
            WriteRecord(ms, BrtFileSharingIso, FileSharingIso(fileSharingHash, fileSharingSalt, fileSharingSpin ?? 100000, fileSharingReadOnlyRecommended));
        WriteRecord(ms, BrtBeginBookViews, Array.Empty<byte>());
        WriteRecord(ms, BrtBookView, BookView());
        WriteRecord(ms, BrtEndBookViews, Array.Empty<byte>());
        WriteRecord(ms, BrtBeginBundleShs, Array.Empty<byte>());
        for (int i = 0; i < sheets.Count; i++)
            WriteRecord(ms, BrtBundleSh, BundleSh(i, sheets[i].SheetName));
        WriteRecord(ms, BrtEndBundleShs, Array.Empty<byte>());
        WriteRecord(ms, BrtEndBook, Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] BookView()
    {
        // 对照 SheetJS write_BrtBookView：29 字节
        // xwPos(4)=0 xwLen(4)=460 xwGap(4)=28800 xwCalcMode? 实际：
        // s32(4)=0 + s32(4)=460 + u32(4)=28800 + u32(4)=17600 + u32(4)=500 + u32(4)=idx + u32(4)=idx + flags(1)=0x78
        var ms = new MemoryStream();
        WriteS32(ms, 0);
        WriteS32(ms, 460);
        WriteU32(ms, 28800);
        WriteU32(ms, 17600);
        WriteU32(ms, 500);
        WriteU32(ms, 0); // 激活表索引
        WriteU32(ms, 0);
        ms.WriteByte(0x78);
        return ms.ToArray();
    }

    private static byte[] FileVersion()
    {
        // 4 个 u32 0 + "LiteExcel" + "2.2.6" + "2.2.6" + "7262"
        var ms = new MemoryStream();
        for (int i = 0; i < 4; i++) WriteU32(ms, 0);
        WriteWideString(ms, "LiteExcel");
        WriteWideString(ms, "2.2.6");
        WriteWideString(ms, "2.2.6");
        WriteWideString(ms, "7262");
        return ms.ToArray();
    }

    private static byte[] WbProp(string? codeName, bool date1904)
    {
        // 对照 Excel：flags(4) + defaultThemeVersion(4) + CodeName(XLWideString，可为空 → cch=0 占 4 字节)
        // flags bit0 = 1 表示 1904 日期系统（读取侧 XlsbBackend 用 &0x01 判断）
        var ms = new MemoryStream();
        WriteU32(ms, date1904 ? 0x00010021u : 0x00010020u);
        WriteU32(ms, 0x0003163C);
        WriteWideString(ms, codeName ?? "");
        return ms.ToArray();
    }

    private static byte[] BundleSh(int index, string name)
    {
        var ms = new MemoryStream();
        WriteU32(ms, 0); // Hidden
        WriteU32(ms, (uint)(index + 1)); // iTabID
        WriteNullableWideString(ms, "rId" + (index + 1));
        WriteWideString(ms, name.Length > 31 ? name.Substring(0, 31) : name);
        return ms.ToArray();
    }

    /// <summary>
    /// BrtFileSharingIso（0x02A4）写保护记录，对齐 Excel 样本布局：
    /// spinCount(4) + flags(2)[bit0=readOnlyRecommended] + stUser(XLWideString "Admin") +
    /// hashValue(4+bytes, base64 解码) + saltValue(4+bytes) + algorithmName(XLWideString "SHA-512")。
    /// </summary>
    private static byte[] FileSharingIso(string hashB64, string? saltB64, int spinCount, bool readOnlyRecommended)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)spinCount);
        WriteU16(ms, (ushort)(readOnlyRecommended ? 1 : 0));
        WriteWideString(ms, "Admin");
        var hash = Convert.FromBase64String(hashB64);
        WriteU32(ms, (uint)hash.Length);
        ms.Write(hash, 0, hash.Length);
        if (!string.IsNullOrEmpty(saltB64))
        {
            var salt = Convert.FromBase64String(saltB64);
            WriteU32(ms, (uint)salt.Length);
            ms.Write(salt, 0, salt.Length);
        }
        else
        {
            WriteU32(ms, 0);
        }
        WriteWideString(ms, "SHA-512");
        return ms.ToArray();
    }

    // ── sharedStrings.bin ──

    private static byte[] BuildSharedStringsBin(List<string> sst, Dictionary<string, int> sstIndex)
    {
        var ms = new MemoryStream();
        var head = new byte[8];
        WriteU32To(head, 0, (uint)sst.Count);   // Count
        WriteU32To(head, 4, (uint)sst.Count);   // Unique
        WriteRecord(ms, BrtBeginSst, head);
        foreach (var s in sst)
            WriteRecord(ms, BrtSSTItem, RichStr(s));
        WriteRecord(ms, BrtEndSst, Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] RichStr(string text)
    {
        var ms = new MemoryStream();
        ms.WriteByte(0); // flags（无富文本）
        WriteWideString(ms, text);
        return ms.ToArray();
    }

    // ── styles.bin ──

    private static byte[] BuildStylesBin(List<(int Ifmt, string? FmtCode)> cellXfs)
    {
        var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginStyleSheet, Array.Empty<byte>());

        // 自定义数字格式
        var customFormats = new List<(int Id, string Code)>();
        var customXfId = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var (_, fmtCode) in cellXfs)
        {
            if (string.IsNullOrEmpty(fmtCode)) continue;
            int ifmt = ResolveFmtId(fmtCode!);
            if (ifmt >= FirstCustomFmtId && !customXfId.ContainsKey(fmtCode!))
            {
                customXfId[fmtCode!] = ifmt;
                customFormats.Add((ifmt, fmtCode!));
            }
        }
        if (customFormats.Count > 0)
        {
            WriteRecord(ms, BrtBeginFmts, UInt32((uint)customFormats.Count));
            foreach (var (id, code) in customFormats)
                WriteRecord(ms, BrtFmt, Fmt(id, code));
            WriteRecord(ms, BrtEndFmts, Array.Empty<byte>());
        }

        // 字体（默认 Calibri）
        WriteRecord(ms, BrtBeginFonts, UInt32(1));
        WriteRecord(ms, BrtFont, Font());
        WriteRecord(ms, BrtEndFonts, Array.Empty<byte>());

        // 填充：none + gray125
        WriteRecord(ms, BrtBeginFills, UInt32(2));
        WriteRecord(ms, BrtFill, Fill("none"));
        WriteRecord(ms, BrtFill, Fill("gray125"));
        WriteRecord(ms, BrtEndFills, Array.Empty<byte>());

        // 边框：1 个空边框
        WriteRecord(ms, BrtBeginBorders, UInt32(1));
        WriteRecord(ms, BrtBorder, Border());
        WriteRecord(ms, BrtEndBorders, Array.Empty<byte>());

        // cellStyleXfs：1 个（默认，ixfeParent=0xFFFF）
        WriteRecord(ms, BrtBeginCellStyleXFs, UInt32(1));
        WriteRecord(ms, BrtXF, Xf(0, 0xFFFF));
        WriteRecord(ms, BrtEndCellStyleXFs, Array.Empty<byte>());

        // cellXfs：索引 0 = General，其余按格式
        WriteRecord(ms, BrtBeginCellXFs, UInt32((uint)cellXfs.Count));
        foreach (var (ifmt, _) in cellXfs)
            WriteRecord(ms, BrtXF, Xf(ifmt, 0));
        WriteRecord(ms, BrtEndCellXFs, Array.Empty<byte>());

        // cellStyles
        WriteRecord(ms, BrtBeginStyles, UInt32(1));
        WriteRecord(ms, BrtStyle, Style());
        WriteRecord(ms, BrtEndStyles, Array.Empty<byte>());

        // dxfs（空）
        WriteRecord(ms, BrtBeginDXFs, UInt32(0));
        WriteRecord(ms, BrtEndDXFs, Array.Empty<byte>());

        // tableStyles（空）
        WriteRecord(ms, BrtBeginTableStyles, TableStylesHead());
        WriteRecord(ms, BrtEndTableStyles, Array.Empty<byte>());

        WriteRecord(ms, BrtEndStyleSheet, Array.Empty<byte>());
        return ms.ToArray();
    }

    private static int ResolveFmtId(string fmtCode)
    {
        // 与 FormatDetector 保持一致：内置日期格式码返回 14，其余内置返回其 ID，未知注册为自定义
        if (fmtCode == "yyyy-MM-dd") return BuiltinDateFmtId;
        for (int id = 1; id < 50; id++)
        {
            var code = FormatDetector.GetBuiltInFormatCode(id);
            if (code == fmtCode) return id;
        }
        return FirstCustomFmtId;
    }

    private static byte[] Fmt(int id, string code)
    {
        var ms = new MemoryStream();
        WriteU16(ms, (ushort)id);
        WriteWideString(ms, code);
        return ms.ToArray();
    }

    private static byte[] Font()
    {
        // 与 Excel 原生输出一致的默认字体（等线，11pt）：直接照抄 29 字节
        // sz(2)=0xDC(220=11pt) grbit(2)=0 weight(2)=0x190 vertAlign(2)=0 underline(1)=0
        // family(1)=2 charset(1)=0x86 pad(1)=0 color(8) scheme(1)=2 name=XLWideString("等线")
        return new byte[]
        {
            0xDC, 0x00, 0x00, 0x00, 0x90, 0x01, 0x00, 0x00, 0x00, 0x02, 0x86, 0x00,
            0x07, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0x02,
            0x02, 0x00, 0x00, 0x00, 0x49, 0x7B, 0xBF, 0x7E,
        };
    }

    private static byte[] Fill(string patternType)
    {
        // BrtFill: fls(4) + fgColor(BrtColor,8) + bgColor(BrtColor,8) + 12×u32(48) = 68 字节
        var fls = patternType == "gray125" ? 0x11 : 0x00;
        var ms = new MemoryStream();
        WriteU32(ms, (uint)fls);
        WriteColorAuto(ms);
        WriteColorAuto(ms);
        for (int j = 0; j < 12; j++) WriteU32(ms, 0);
        return ms.ToArray();
    }

    private static byte[] Border()
    {
        // diagonal(1) + 5 × Blxf(10)
        var ms = new MemoryStream();
        ms.WriteByte(0);
        for (int i = 0; i < 5; i++)
        {
            ms.WriteByte(0); ms.WriteByte(0);
            WriteU32(ms, 0); WriteU32(ms, 0);
        }
        return ms.ToArray();
    }

    private static byte[] Xf(int ifmt, int ixfeParent)
    {
        // 16 字节，对照 Excel 原生输出：
        // ixfeParent(2) ifmt(2) iFont(2) iFill(2) ixBorder(2) trot(1) indent(1) flow(1)=0x08 pad(1)=0x10 pad(1) pad(1)
        // Excel: cellStyleXfs = FF FF 00 00 00 00 00 00 00 00 00 00 08 10 00 00
        //        cellXfs     = 00 00 00 00 00 00 00 00 00 00 00 00 08 10 00 00
        return new byte[]
        {
            (byte)(ixfeParent & 0xFF), (byte)(ixfeParent >> 8),
            (byte)(ifmt & 0xFF), (byte)(ifmt >> 8),
            0, 0, 0, 0, 0, 0,
            0, 0,
            0x08, 0x10, 0x00, 0x00,
        };
    }

    private static byte[] Style()
    {
        // 对照 Excel 原生默认样式：xfId(4)=0 flags(2)=1 builtinId(1)=0 iLevel(1)=0
        // name = XLNullableWideString(等线) → 16 字节
        return new byte[]
        {
            0x00, 0x00, 0x00, 0x00,
            0x01, 0x00,
            0x00, 0x00,
            0x02, 0x00, 0x00, 0x00,
            0x38, 0x5E, 0xC4, 0x89,
        };
    }

    private static byte[] TableStylesHead()
    {
        // cnt(4) + defaultTableStyle(XLNullableWideString) + defaultPivotStyle(XLNullableWideString)
        var ms = new MemoryStream();
        WriteU32(ms, 0);
        WriteNullableWideString(ms, "TableStyleMedium9");
        WriteNullableWideString(ms, "PivotStyleMedium4");
        return ms.ToArray();
    }

    // ── worksheet.bin ──

    private static byte[] BuildWorksheetBin(SheetData sheet, Dictionary<string, int> sstIndex, Func<string?, int> getXf, bool date1904, List<string> extTargets)
    {
        var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginSheet, Array.Empty<byte>());

        var extRelIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < extTargets.Count; i++) extRelIndex[extTargets[i]] = i;

        // 计算范围
        int maxRow = -1, maxCol = -1;
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Count; c++)
            {
                if (row[c].IsEmpty) continue;
                if (r > maxRow) maxRow = r;
                if (c > maxCol) maxCol = c;
            }
        }
        foreach (var m in sheet.MergedRanges)
        {
            if (m.LastRow > maxRow) maxRow = m.LastRow;
            if (m.LastCol > maxCol) maxCol = m.LastCol;
        }
        if (sheet.ColumnWidths is { } widths && widths.Count - 1 > maxCol)
            maxCol = widths.Count - 1;
        int dimR = maxRow < 0 ? 0 : maxRow;
        int dimC = maxCol < 0 ? 0 : maxCol;

        // WsProp（必选；含 sheet CodeName，可空）
        WriteRecord(ms, BrtWsProp, WsProp(sheet.CodeName));

        // WsDim
        WriteRecord(ms, BrtWsDim, RfX(0, dimR, 0, dimC));

        // 视图（冻结）
        WriteRecord(ms, BrtBeginWsViews, Array.Empty<byte>());
        WriteRecord(ms, BrtBeginWsView, WsView());
        int freezeRows = sheet.FreezeRows;
        int freezeCols = sheet.FreezeColumns;
        if (sheet.FreezeHeader) freezeRows = Math.Max(freezeRows, 1);
        if (freezeRows > 0 || freezeCols > 0)
            WriteRecord(ms, BrtPane, Pane(freezeRows, freezeCols));
        WriteRecord(ms, BrtEndWsView, Array.Empty<byte>());
        WriteRecord(ms, BrtEndWsViews, Array.Empty<byte>());

        // 列宽
        if (sheet.ColumnWidths is { } cw && cw.Count > 0)
        {
            var any = false;
            var colInfos = new MemoryStream();
            for (int c = 0; c < cw.Count; c++)
            {
                if (cw[c] <= 0) continue;
                any = true;
                WriteRecord(colInfos, BrtColInfo, ColInfo(c, cw[c]));
            }
            if (any)
            {
                WriteRecord(ms, BrtBeginColInfos, Array.Empty<byte>());
                colInfos.Position = 0;
                colInfos.CopyTo(ms);
                WriteRecord(ms, BrtEndColInfos, Array.Empty<byte>());
            }
        }

        // 单元格数据
        WriteRecord(ms, BrtBeginSheetData, Array.Empty<byte>());
        for (int r = 0; r <= maxRow; r++)
        {
            var row = sheet.Rows.Count > r ? sheet.Rows[r] : null;
            var cells = new List<(int Col, Cell Cell)>();
            if (row is not null)
            {
                for (int c = 0; c <= maxCol && c < row.Count; c++)
                {
                    if (row[c].IsEmpty) continue;
                    cells.Add((c, row[c]));
                }
            }
            if (cells.Count == 0) continue;

            int prevCol = -1;
            WriteRecord(ms, BrtRowHdr, RowHdr(r, cells[0].Col, cells[cells.Count - 1].Col, sheet, r));
            bool firstInRow = true;
            foreach (var (col, cell) in cells)
            {
                WriteCell(ms, col, cell, sstIndex, getXf, ref prevCol, firstInRow, date1904);
                firstInRow = false;
            }
        }
        WriteRecord(ms, BrtEndSheetData, Array.Empty<byte>());

        // 合并单元格
        if (sheet.MergedRanges.Count > 0)
        {
            WriteRecord(ms, BrtBeginMergeCells, UInt32((uint)sheet.MergedRanges.Count));
            foreach (var m in sheet.MergedRanges)
                WriteRecord(ms, BrtMergeCell, RfX(m.FirstRow, m.LastRow, m.FirstCol, m.LastCol));
            WriteRecord(ms, BrtEndMergeCells, Array.Empty<byte>());
        }

        // 超链接（外部经 relId 指向 sheet rels；内部走 location）
        WriteHyperlinks(ms, sheet, extRelIndex);

        WriteRecord(ms, BrtEndSheet, Array.Empty<byte>());
        return ms.ToArray();
    }

    /// <summary>写出 BrtHLink 记录。外部链接 relId = rIdH{n+1}；内部链接 location = 目标去前导 '#' </summary>
    private static void WriteHyperlinks(MemoryStream ms, SheetData sheet, Dictionary<string, int> extRelIndex)
    {
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Count; c++)
            {
                var link = row[c].Hyperlink;
                if (link is null || string.IsNullOrEmpty(link.Target)) continue;

                var inner = new MemoryStream();
                WriteRfX(inner, r, r, c, c);
                if (link.IsInternal)
                {
                    WriteNullableWideString(inner, null);
                    var loc = link.Target.StartsWith("#", StringComparison.Ordinal) ? link.Target.Substring(1) : link.Target;
                    WriteWideString(inner, loc);
                    WriteWideString(inner, link.Tooltip ?? "");
                    WriteWideString(inner, "");
                }
                else
                {
                    int relIdx = extRelIndex.TryGetValue(link.Target, out var i) ? i : 0;
                    WriteNullableWideString(inner, "rIdH" + (relIdx + 1));
                    WriteWideString(inner, "");
                    WriteWideString(inner, link.Tooltip ?? "");
                    WriteWideString(inner, "");
                }
                WriteRecord(ms, BrtHLink, inner.ToArray());
            }
        }
    }

    private static void WriteRfX(MemoryStream ms, int rwFirst, int rwLast, int colFirst, int colLast)
    {
        WriteS32(ms, rwFirst);
        WriteS32(ms, rwLast);
        WriteS32(ms, colFirst);
        WriteS32(ms, colLast);
    }

    private static byte[] WsProp(string? codeName)
    {
        // flags(1)=0xC0 + padding(2) + BrtColor(auto,8) + s32×2(-1) + CodeName(XLWideString)
        var ms = new MemoryStream();
        ms.WriteByte(0xC0);
        ms.WriteByte(0);
        ms.WriteByte(0);
        WriteColorAuto(ms);
        WriteS32(ms, -1);
        WriteS32(ms, -1);
        WriteWideString(ms, codeName ?? "");
        return ms.ToArray();
    }

    private static byte[] WsView()
    {
        // flags(2)=0x39C + xview(4)=0 + rwTop(4)=0 + colLeft(4)=0 + gridlineColor(1)+pad(1)+u16(2)
        // + zoomScale(2)=100 + u16×3(6) + workbookViewId(4)=0  → 26 字节（Excel 为 30，缺 4 字节补 0）
        var ms = new MemoryStream();
        WriteU16(ms, 0x039C);
        WriteU32(ms, 0);
        WriteU32(ms, 0);
        WriteU32(ms, 0);
        ms.WriteByte(0);
        ms.WriteByte(0);
        WriteU16(ms, 0);
        WriteU16(ms, 100);
        WriteU16(ms, 0);
        WriteU16(ms, 0);
        WriteU16(ms, 0);
        WriteU32(ms, 0);
        return ms.ToArray();
    }

    private static byte[] Pane(int freezeRows, int freezeCols)
    {
        // colFrozen(Xnum 8) + rowFrozen(Xnum 8) + topLeftCell 行(4) + 列(4) + activePane(4) + state(1)=frozen
        var ms = new MemoryStream();
        WriteDouble(ms, freezeCols); // colFrozen
        WriteDouble(ms, freezeRows); // rowFrozen
        WriteU32(ms, (uint)freezeRows); // topLeftCell 行
        WriteU32(ms, (uint)freezeCols); // topLeftCell 列
        uint activePane = freezeRows > 0 && freezeCols > 0 ? 0u : freezeRows > 0 ? 2u : 1u;
        WriteU32(ms, activePane);
        ms.WriteByte(0x01); // state = frozen
        return ms.ToArray();
    }

    private static byte[] RowHdr(int rw, int colFirst, int colLast, SheetData sheet, int rowIndex)
    {
        var ms = new MemoryStream();
        WriteS32(ms, rw);
        WriteU32(ms, 0); // ixfe
        int miyRw = 0x0140; // 20pt 默认
        if (sheet.RowHeights is not null && sheet.RowHeights.TryGetValue(rowIndex, out var h))
            miyRw = (int)Math.Round(h * 20);
        WriteU16(ms, (ushort)miyRw);
        ms.WriteByte(0); // top/bot padding
        byte flags = 0;
        if (sheet.RowHeights is not null && sheet.RowHeights.TryGetValue(rowIndex, out _))
            flags |= 0x20; // Excel 原生 BrtRowHdr：b11 & 0x20 标记显式行高（与读取端/XlsbTestFile 一致）
        ms.WriteByte(flags);
        ms.WriteByte(0); // phonetic
        WriteU32(ms, 1); // ncolspan
        WriteS32(ms, colFirst);
        WriteS32(ms, colLast);
        return ms.ToArray();
    }

    private static void WriteCell(MemoryStream ms, int col, Cell cell, Dictionary<string, int> sstIndex,
        Func<string?, int> getXf, ref int prevCol, bool firstInRow, bool date1904)
    {
        int xf = getXf(cell.NumberFormat);
        bool lastSeen = !firstInRow && col == prevCol + 1;

        switch (cell.Type)
        {
            case CellType.Text:
                if (cell.Text is null) { WriteCellBlank(ms, col, xf, lastSeen); return; }
                if (sstIndex.TryGetValue(cell.Text, out var sstIdx))
                {
                    if (lastSeen)
                        WriteRecord(ms, BrtShortIsst, ShortIsst(xf, sstIdx));
                    else
                        WriteRecord(ms, BrtCellIsst, CellIsst(col, xf, sstIdx));
                }
                else
                {
                    if (lastSeen)
                        WriteRecord(ms, BrtShortSt, ShortSt(xf, cell.Text));
                    else
                        WriteRecord(ms, BrtCellSt, CellSt(col, xf, cell.Text));
                }
                break;
            case CellType.Number:
            case CellType.Date:
                double v = cell.Type == CellType.Date ? FormatDetector.DateToSerial(cell.Date, date1904) : cell.Number;
                // 整数小值用 RK，其余用 Real
                if (v == Math.Floor(v) && v > -1000 && v < 1000)
                {
                    if (lastSeen)
                        WriteRecord(ms, BrtShortRk, ShortRk(xf, v));
                    else
                        WriteRecord(ms, BrtCellRk, CellRk(col, xf, v));
                }
                else
                {
                    if (lastSeen)
                        WriteRecord(ms, BrtShortReal, ShortReal(xf, v));
                    else
                        WriteRecord(ms, BrtCellReal, CellReal(col, xf, v));
                }
                break;
            case CellType.Boolean:
                if (lastSeen)
                    WriteRecord(ms, BrtShortBool, ShortBool(xf, cell.Boolean));
                else
                    WriteRecord(ms, BrtCellBool, CellBool(col, xf, cell.Boolean));
                break;
            default:
                WriteCellBlank(ms, col, xf, lastSeen);
                break;
        }
        prevCol = col;
    }

    private static void WriteCellBlank(MemoryStream ms, int col, int xf, bool lastSeen)
    {
        if (lastSeen)
            WriteRecord(ms, BrtShortBlank, ShortCell(xf));
        else
            WriteRecord(ms, BrtCellBlank, Cell(col, xf));
    }

    private static byte[] Cell(int col, int xf)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        return ms.ToArray();
    }

    private static byte[] ShortCell(int xf)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        return ms.ToArray();
    }

    private static byte[] CellIsst(int col, int xf, int sstIdx)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        WriteS32(ms, sstIdx);
        return ms.ToArray();
    }

    private static byte[] ShortIsst(int xf, int sstIdx)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        WriteS32(ms, sstIdx);
        return ms.ToArray();
    }

    private static byte[] CellSt(int col, int xf, string text)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        WriteWideString(ms, text);
        return ms.ToArray();
    }

    private static byte[] ShortSt(int xf, string text)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        WriteWideString(ms, text);
        return ms.ToArray();
    }

    private static byte[] CellRk(int col, int xf, double v)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        WriteU32(ms, RkNumber(v));
        return ms.ToArray();
    }

    private static byte[] ShortRk(int xf, double v)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        WriteU32(ms, RkNumber(v));
        return ms.ToArray();
    }

    private static byte[] CellReal(int col, int xf, double v)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        WriteDouble(ms, v);
        return ms.ToArray();
    }

    private static byte[] ShortReal(int xf, double v)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        WriteDouble(ms, v);
        return ms.ToArray();
    }

    private static byte[] CellBool(int col, int xf, bool b)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteU32(ms, (uint)xf);
        ms.WriteByte(b ? (byte)1 : (byte)0);
        return ms.ToArray();
    }

    private static byte[] ShortBool(int xf, bool b)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)xf);
        ms.WriteByte(b ? (byte)1 : (byte)0);
        return ms.ToArray();
    }

    private static byte[] ColInfo(int col, double width)
    {
        var ms = new MemoryStream();
        WriteS32(ms, col);
        WriteS32(ms, col);
        WriteU32(ms, (uint)Math.Round(width * 256));
        WriteU32(ms, 0); // ixfe
        WriteU16(ms, 0x0002); // flags: fWidth
        return ms.ToArray();
    }

    private static byte[] RfX(int sRow, int eRow, int sCol, int eCol)
    {
        var ms = new MemoryStream();
        WriteS32(ms, sRow);
        WriteS32(ms, eRow);
        WriteS32(ms, sCol);
        WriteS32(ms, eCol);
        return ms.ToArray();
    }

    private static uint RkNumber(double v)
    {
        // 仅整数小值调用。整数编码：fInt=1，val = v << 2
        long iv = (long)v;
        return (uint)((iv << 2) | 0x02);
    }

    // ── 记录与标量写入辅助 ──

    private static void WriteRecord(MemoryStream ms, int rt, byte[] data)
    {
        WriteVarInt(ms, rt);
        WriteVarInt(ms, data.Length);
        ms.Write(data, 0, data.Length);
    }

    private static void WriteVarInt(MemoryStream ms, int v)
    {
        uint value = (uint)v;
        while (value >= 0x80)
        {
            ms.WriteByte((byte)(value | 0x80));
            value >>= 7;
        }
        ms.WriteByte((byte)value);
    }

    private static void WriteWideString(MemoryStream ms, string s)
    {
        WriteU32(ms, (uint)s.Length);
        var bytes = Encoding.Unicode.GetBytes(s);
        ms.Write(bytes, 0, bytes.Length);
    }

    private static void WriteNullableWideString(MemoryStream ms, string s)
    {
        if (s is null || s.Length == 0)
        {
            WriteU32(ms, 0xFFFFFFFF);
            return;
        }
        WriteU32(ms, (uint)s.Length);
        var bytes = Encoding.Unicode.GetBytes(s);
        ms.Write(bytes, 0, bytes.Length);
    }

    private static void WriteU32(MemoryStream ms, uint v) => ms.Write(new[] { (byte)v, (byte)(v >> 8), (byte)(v >> 16), (byte)(v >> 24) }, 0, 4);

    private static void WriteU32To(byte[] b, int off, uint v)
    {
        b[off] = (byte)v;
        b[off + 1] = (byte)(v >> 8);
        b[off + 2] = (byte)(v >> 16);
        b[off + 3] = (byte)(v >> 24);
    }

    private static void WriteS32(MemoryStream ms, int v) => WriteU32(ms, unchecked((uint)v));

    private static void WriteU16(MemoryStream ms, ushort v) => ms.Write(new[] { (byte)v, (byte)(v >> 8) }, 0, 2);

    private static void WriteDouble(MemoryStream ms, double v)
    {
        var b = BitConverter.GetBytes(v);
        ms.Write(b, 0, 8);
    }

    private static void WriteColorAuto(MemoryStream ms)
    {
        // 8 字节全 0 = auto 颜色
        WriteU32(ms, 0);
        WriteU32(ms, 0);
    }

    private static byte[] UInt32(uint v)
    {
        var b = new byte[4];
        WriteU32To(b, 0, v);
        return b;
    }
}
