using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using LiteExcel.Internal.Biff;
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
internal static partial class XlsbWriter
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
    private const int BrtDXF = 0x01FB;
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
    private const int BrtBeginSheetData = 0x01E5;
    private const int BrtRowHdr = 0x0000;
    private const int BrtCellBlank = 0x0001;
    private const int BrtCellRk = 0x0002;
    private const int BrtCellBool = 0x0004;
    private const int BrtCellReal = 0x0005;
    private const int BrtCellSt = 0x0006;
    private const int BrtFmlaString = 0x0008;
    private const int BrtFmlaNum = 0x0009;
    private const int BrtFmlaBool = 0x000A;
    private const int BrtFmlaError = 0x000B;
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
    private const int BrtBeginAFilter = 0x00A1;
    private const int BrtEndAFilter = 0x00A2;
    private const int BrtHLink = 0x01EE;
    private const int BrtEndSheet = 0x0082;

    // 超级表（Table/ListObject）
    private const int BrtBeginTableParts = 0x0294; // cTableParts(4)
    private const int BrtTablePart = 0x0295;       // XLWideString rId
    private const int BrtEndTableParts = 0x0296;
    private const int BrtBeginList = 0x0157;       // 表属性
    private const int BrtEndList = 0x0158;         // 表结束
    private const int BrtBeginListCols = 0x0159;   // cCols(4)
    private const int BrtEndListCols = 0x015A;
    private const int BrtBeginListCol = 0x015B;    // 列定义
    private const int BrtTableStyleClient = 0x0201; // 表样式

    // 绘图（图片/图表/形状）：BrtDrawing 引用 sheet rels 中的 drawing 关系
    private const int BrtDrawing = 0x0226;          // XLWideString rId

    // 数据验证
    private const int BrtBeginDVs = 0x023D;
    private const int BrtEndDVs = 0x023E;
    private const int BrtDVal = 0x0040;

    // 条件格式（conditionalFormatting）：容器 + 规则 + 类型子记录
    private const int BrtBeginCF = 0x01CD;         // cCF(4) + reserved(4) + cSqref(4) + Ref8U×cSqref
    private const int BrtCFRule = 0x01CF;          // 规则体（46 字节头 + 公式段）
    private const int BrtEndCFRule = 0x01D0;
    private const int BrtEndCF = 0x01CE;
    private const int BrtBeginIconSet = 0x01D1;    // iTemplate(4) + flags(2)
    private const int BrtEndIconSet = 0x01D2;
    private const int BrtCFVO = 0x01D7;            // type(4) + value(f64) + f1(4) + f2(4) + 0(4)
    private const int BrtBeginColorScale = 0x01D5; // 0 字节
    private const int BrtEndColorScale = 0x01D6;   // 0 字节
    private const int BrtBeginDataBar = 0x01D3;    // fMinLength(u8) + fMaxLength(u8) + flags(u8)
    private const int BrtEndDataBar = 0x01D4;      // 0 字节
    private const int BrtColor = 0x0234;           // BrtColor(8)：05 FF 00 00 R G B FF

    private const int DefaultFontId = 0;
    private const int DefaultFillId = 0;
    private const int DefaultBorderId = 0;
    private const int DefaultCellStyleXf = 0;
    private const int DefaultCellXf = 0;
    private const int BuiltinDateFmtId = 14;
    private const int FirstCustomFmtId = 164;

    /// <summary>写出 .xlsb 工作簿到流。vbaProject 为源工作簿捕获的宏工程字节（可为 null）；workbookCodeName 为宿主代码名（可为 null）。
    /// <paramref name="preserved"/> 为打开时捕获的未重建 OOXML 部件（图表/透视表/主题/绘图等），保存时透传；
    /// <paramref name="properties"/> 为文档属性，非 null 时写出 docProps。
    /// <paramref name="verbatim"/> = true 时原样保留 workbook.bin / styles.bin / sharedStrings.bin / sheetN.bin 及其 rels，
    /// 不重建（保留透视表/切片器等 BIFF12 宿主记录）。仅当工作簿结构不变且无单元格修改时由调用方启用。</summary>
    public static void Write(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject = null, string? workbookCodeName = null, bool date1904 = false,
        string? fileSharingHash = null, string? fileSharingSalt = null, int? fileSharingSpin = null, bool fileSharingReadOnlyRecommended = false,
        Action<DegradationInfo>? onDegradation = null, ExcelFormat targetFormat = ExcelFormat.Xlsb,
        OoxmlPreservedParts? preserved = null, WorkbookProperties? properties = null,
        IReadOnlyList<NamedRange>? names = null, bool verbatim = false, bool surgical = false, bool allowFeatureLoss = false,
        bool surgicalEdit = false)
    {
        if (sheets is null || sheets.Count == 0)
            throw new ArgumentException("至少需要一张工作表", nameof(sheets));

        // 原样模式需要完整的原始二进制部件。
        verbatim = verbatim && preserved?.VerbatimBinaries is not null
            && preserved.VerbatimBinaries.ContainsKey("xl/workbook.bin")
            && preserved.VerbatimBinaries.ContainsKey("xl/styles.bin");

        // 跨格式兼容（D2）：保留部件仅能与源容器格式匹配时透传。
        // 源为 xlsx/xlsm（XML-OOXML）的保留部件（透视表/切片器/连接等 XML 部件）
        // 若塞入 xlsb（BIFF12 容器）会产生结构性损坏（Excel 报"文件级验证和修复"）。
        if (preserved is not null && preserved.VerbatimBinaries is null)
        {
            ReportCrossFormatPreservedDrop(onDegradation, targetFormat, preserved);
            preserved = null;
        }

        // 手术式删除通道：仅删除若干工作表，其余二进制部件原样保留（含数据模型/透视/连接等全部高级部件）。
        if (surgical && preserved is not null && preserved.VerbatimBinaries is not null
            && preserved.VerbatimBinaries.ContainsKey("xl/workbook.bin"))
        {
            WriteSurgicalXlsb(stream, sheets, vbaProject, preserved, properties);
            return;
        }

        // 手术式编辑通道：表结构不变、仅内容修改，只对被改表做字节级单元格补丁，其余部件逐字节保留。
        if (surgicalEdit && preserved is not null && preserved.VerbatimBinaries is not null
            && preserved.VerbatimBinaries.ContainsKey("xl/workbook.bin"))
        {
            WriteSurgicalEditXlsb(stream, sheets, vbaProject, preserved, properties);
            return;
        }

        if (!verbatim)
        {
            // 手术式通道（verbatim / surgical delete / surgical edit）全部保真透传，已在上方 return，不产生降级；
            // 仅重建/转换路径才上报降级 —— 避免 verbatim 无损保存时误报 NamedRanges/DocumentProperties/Styles 等伪降级。
            ReportDegradations(sheets, names, properties, onDegradation, targetFormat);

            var sst = new List<string>();
            var sstIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            var cellXfs = new List<(int Ifmt, string? FmtCode)>();
            var fmtCodeToXf = new Dictionary<string, int>(StringComparer.Ordinal);
            var customFmtIds = new Dictionary<string, int>(StringComparer.Ordinal);
            cellXfs.Add((0, null)); // 索引 0 = General 默认样式

            int GetXf(string? fmtCode)
            {
                if (string.IsNullOrEmpty(fmtCode)) return DefaultCellXf;
                if (fmtCodeToXf.TryGetValue(fmtCode, out var idx)) return idx;
                idx = cellXfs.Count;
                fmtCodeToXf[fmtCode] = idx;
                int ifmt = ResolveFmtId(fmtCode);
                if (ifmt >= FirstCustomFmtId)
                {
                    if (!customFmtIds.TryGetValue(fmtCode, out var customId))
                    {
                        customId = FirstCustomFmtId + customFmtIds.Count;
                        customFmtIds[fmtCode] = customId;
                    }
                    ifmt = customId;
                }
                cellXfs.Add((ifmt, fmtCode));
                return idx;
            }

            void ScanCell(Cell cell)
            {
                if (cell.IsEmpty) return;
                if (!string.IsNullOrEmpty(cell.NumberFormat))
                    GetXf(cell.NumberFormat);
                if (cell.Type == CellType.Text && cell.Text is not null && !sstIndex.ContainsKey(cell.Text))
                {
                    sstIndex[cell.Text] = sst.Count;
                    sst.Add(cell.Text);
                }
            }

            foreach (var sheet in sheets)
                foreach (var row in sheet.Rows)
                    foreach (var cell in row)
                        ScanCell(cell);

            WriteRebuilt(stream, sheets, vbaProject, workbookCodeName, date1904,
                fileSharingHash, fileSharingSalt, fileSharingSpin, fileSharingReadOnlyRecommended,
                preserved, properties, sst, sstIndex, cellXfs, GetXf);
        }
        else
        {
            WriteVerbatim(stream, sheets, vbaProject, preserved, properties);
        }
    }

    private static void WriteRebuilt(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject, string? workbookCodeName, bool date1904,
        string? fileSharingHash, string? fileSharingSalt, int? fileSharingSpin, bool fileSharingReadOnlyRecommended,
        OoxmlPreservedParts? preserved, WorkbookProperties? properties,
        List<string> sst, Dictionary<string, int> sstIndex, List<(int Ifmt, string? FmtCode)> cellXfs, Func<string?, int> GetXf)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        // 图片规划：分配 media 序号、生成 drawing 部件（xlsb 的 drawing/media 与 xlsx 同为 XML 部件）。
        var imagePlan = XlsxWriter.ImagePlan.Create(sheets, preserved, binary: true);

        // 条件格式样式（DXF）：全局去重，供 sheet bin 的 BrtCFRule.dxfId 与 styles.bin 的 BrtDXF 共用。
        var dxfRegistry = new DxfRegistry();
        foreach (var sheet in sheets)
            if (sheet.ConditionalFormats is { Count: > 0 })
                foreach (var cf in sheet.ConditionalFormats)
                    if (cf.Style is not null) dxfRegistry.GetOrCreate(cf.Style);

        // 源为含数据模型（Power Pivot / Power Query，xl/model/item.data）的 xlsb 时，重建会写出全新的
        // workbook.bin，其中不含绑定数据模型的 BIFF12 宿主记录（数据模型 FRT 记录）；若仍原样透传
        // 数据模型/透视缓存/连接/自定义 XML 等高级部件，会产生「部件存在但 workbook.bin 未声明」的孤儿，
        // Excel 打开时报文件级修复/拒绝。故此类重建丢弃全部高级部件（经降级回调上报）。
        // 其余情形（如图表等宿主记录随 sheet 一并重建的部件）保持原样透传。
        // 写入保留部件，跳过由写入器重建的条目。
        // 跨格式转换（源非 xlsb）时不写入 XML 保留部件——它们无法转为 BIFF12 二进制格式。
        if (preserved is not null && preserved.VerbatimBinaries is not null)
        {
            var rebuilt = OoxmlPreservedParts.BuildRebuiltEntries(sheets.Count, binary: true);
            // 图片规划会重建/合并的 drawing 部件须跳过，避免与下方写出重复（zip 重名）。
            var imageEntries = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (entry, _) in imagePlan.FloatingDrawingParts(preserved))
                imageEntries.Add(entry);
            foreach (var kv in preserved.Parts)
            {
                if (rebuilt.Contains(kv.Key)) continue;
                if (imageEntries.Contains(kv.Key)) continue;
                // 超级表部件由下方按模型重建写出，源里的同名部件必须跳过，否则会重复。
                if (kv.Key.StartsWith("xl/tables/table", StringComparison.Ordinal)
                    && kv.Key.EndsWith(".bin", StringComparison.Ordinal)) continue;
                WriteEntry(zip, kv.Key, kv.Value);
            }
        }

        // 包结构
            var sheetsWithComments = new List<int>();
            for (int i = 0; i < sheets.Count; i++)
            {
                if (sheets[i].Comments is { Count: > 0 })
                    sheetsWithComments.Add(i + 1);
            }

        int totalTables = 0;
        for (int i = 0; i < sheets.Count; i++)
            totalTables += sheets[i].Tables?.Count ?? 0;
        int tableCounter = 0;

WriteEntry(zip, "[Content_Types].xml", ContentTypesXml(sheets.Count, sst.Count > 0, vbaProject is not null, properties is not null, preserved, sheetsWithComments, totalTables, imagePlan));
        WriteEntry(zip, "_rels/.rels", RootRelsXml(properties is not null));
        WriteEntry(zip, "xl/workbook.bin", BuildWorkbookBin(sheets, workbookCodeName, date1904, fileSharingHash, fileSharingSalt, fileSharingSpin, fileSharingReadOnlyRecommended));
        WriteEntry(zip, "xl/_rels/workbook.bin.rels", WorkbookRelsXml(sheets.Count, sst.Count > 0, vbaProject is not null, preserved));
        WriteEntry(zip, "xl/styles.bin", BuildStylesBin(cellXfs, dxfRegistry.Styles));
        if (vbaProject is not null && vbaProject.Length > 0)
            WriteEntry(zip, "xl/vbaProject.bin", vbaProject);
        if (sst.Count > 0)
            WriteEntry(zip, "xl/sharedStrings.bin", BuildSharedStringsBin(sst, sstIndex));
        if (properties is not null)
        {
            WriteEntry(zip, "docProps/core.xml", XlsxWriter.CorePropsXml(properties));
            WriteEntry(zip, "docProps/app.xml", XlsxWriter.AppPropsXml(properties, sheets));
        }

        // 浮动图片：media + drawing（drawing/media 在 xlsb 中与 xlsx 同为 XML/二进制标准部件）
        foreach (var (entry, bytes) in imagePlan.MediaEntries())
            WriteEntry(zip, entry, bytes);
        foreach (var (entry, xml) in imagePlan.FloatingDrawingParts(preserved))
            WriteEntry(zip, entry, xml);

        for (int i = 0; i < sheets.Count; i++)
        {
            var extLinks = CollectExternalHyperlinks(sheets[i]);
            bool hasNewFloating = imagePlan.FloatingBySheet[i].Count > 0;
            bool hasPreservedDrawing = HasPreservedDrawingRel(preserved, i + 1);
            bool hasDrawing = hasNewFloating || hasPreservedDrawing;
            string? drawingRelId = hasDrawing ? imagePlan.DrawingTargetFor(i, preserved).RelId : null;
            WriteEntry(zip, $"xl/worksheets/sheet{i + 1}.bin", BuildWorksheetBin(sheets[i], sstIndex, GetXf, date1904, extLinks, drawingRelId, dxfRegistry));
            bool hasComments = sheets[i].Comments is { Count: > 0 };
            if (hasComments)
            {
                WriteEntry(zip, $"xl/comments{i + 1}.bin", BuildCommentsBin(sheets[i].Comments!));
                WriteEntry(zip, $"xl/drawings/vmlDrawing{i + 1}.vml", XlsxWriter.VmlDrawingXml(sheets[i].Comments!));
            }
            // 超级表部件：xl/tables/tableN.bin（全局编号）
            int tableIdBase = -1;
            if (sheets[i].Tables is { Count: > 0 })
            {
                tableIdBase = tableCounter + 1;
                for (int t = 0; t < sheets[i].Tables.Count; t++)
                {
                    tableCounter++;
                    WriteEntry(zip, $"xl/tables/table{tableCounter}.bin", BuildTableBin(sheets[i].Tables[t], tableCounter));
                }
            }
            var sheetRels = BuildSheetRelsXml(i + 1, extLinks, preserved, hasComments, sheets[i], tableIdBase,
                hasDrawing, hasNewFloating, imagePlan, i);
            if (!string.IsNullOrEmpty(sheetRels))
                WriteEntry(zip, $"xl/worksheets/_rels/sheet{i + 1}.bin.rels", sheetRels);
        }
    }

    /// <summary>源 xlsb 保留的工作表关系里是否已有 drawing 关联（图表/图片），据此在新 sheetN.bin 中写 BrtDrawing 引用。</summary>
    internal static bool HasPreservedDrawingRel(OoxmlPreservedParts? preserved, int sheetNumber)
    {
        if (preserved is null) return false;
        if (!preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetNumber}.bin.rels", out var relsXml)) return false;
        foreach (var rel in XlsxWriter.ParseRels(relsXml))
            if (rel.Type.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    /// <summary>XLSB 原样写出：直接保留 workbook.bin / styles.bin / sheetN.bin 等原始二进制部件。
    /// 保留透视表/切片器等 BIFF12 宿主记录。仅当工作簿结构不变且无修改时调用。</summary>
    private static void WriteVerbatim(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject,
        OoxmlPreservedParts? preserved, WorkbookProperties? properties)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        var vb = preserved!.VerbatimBinaries!;

        // 写入非重建部件，例如透视表、切片器、缓存、绘图和主题。
        var rebuilt = OoxmlPreservedParts.BuildRebuiltEntries(sheets.Count, binary: true);
        if (preserved.Parts.Count > 0)
        {
            foreach (var kv in preserved.Parts)
            {
                if (rebuilt.Contains(kv.Key)) continue;
                if (kv.Key == "xl/vbaProject.bin") continue; // 由下方显式写出，避免重复条目
                WriteEntry(zip, kv.Key, kv.Value);
            }
        }

        // Content_Types / root rels（重建，但已合并保留类型声明）
        bool hasSst = vb.ContainsKey("xl/sharedStrings.bin");
        WriteEntry(zip, "[Content_Types].xml", ContentTypesXml(sheets.Count, hasSst, vbaProject is not null, properties is not null, preserved));
        WriteEntry(zip, "_rels/.rels", RootRelsXml(properties is not null));

        // 原样写出 workbook.bin。
        if (vb.TryGetValue("xl/workbook.bin", out var wbBin))
            WriteEntry(zip, "xl/workbook.bin", wbBin);

        // 原样写出 workbook.bin.rels，保留原关系 ID。
        if (preserved.Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var wbRels))
            WriteEntry(zip, "xl/_rels/workbook.bin.rels", wbRels);

        // 原样写出 styles.bin。
        if (vb.TryGetValue("xl/styles.bin", out var styBin))
            WriteEntry(zip, "xl/styles.bin", styBin);

        // 写出 VBA 部件。
        if (vbaProject is not null && vbaProject.Length > 0)
            WriteEntry(zip, "xl/vbaProject.bin", vbaProject);

        // 原样写出 sharedStrings.bin（若存在）。
        if (vb.TryGetValue("xl/sharedStrings.bin", out var sstBin))
            WriteEntry(zip, "xl/sharedStrings.bin", sstBin);

        // docProps（重建，可能用户修改了属性）
        if (properties is not null)
        {
            WriteEntry(zip, "docProps/core.xml", XlsxWriter.CorePropsXml(properties));
            WriteEntry(zip, "docProps/app.xml", XlsxWriter.AppPropsXml(properties, sheets));
        }

        // 原样写出各工作表部件及其关系。
        for (int i = 0; i < sheets.Count; i++)
        {
            var sheetPath = $"xl/worksheets/sheet{i + 1}.bin";
            if (vb.TryGetValue(sheetPath, out var sheetBin))
                WriteEntry(zip, sheetPath, sheetBin);

            var relsPath = $"xl/worksheets/_rels/sheet{i + 1}.bin.rels";
            if (preserved.Rels.TryGetValue(relsPath, out var sheetRels))
                WriteEntry(zip, relsPath, sheetRels);
        }
    }

    /// <summary>XLSB 保真手术式删除：原样复制整个包，并复刻 Excel「删除工作表后另存」的全部同步改动，
    /// 使输出与 Excel 自身产出等价（保留透视表/缓存/超级表/连接/宏/数据模型等全部高级部件）：
    /// <list type="number">
    /// <item>移除被删表的 BrtBundleSh；BrtExternSheet 中引用被删表的条目 itab 置 -1，其余递减；</item>
    /// <item>sheet-local 定义名移除，其余 itab 递减；引用被删表的定义名 rgce 失效化；
    ///   `_xlcn.LinkedTable_*` 连接名去掉尾部冗余 "1"（workbook.bin 与 connections.bin 同步）；</item>
    /// <item>剩余 sheet/table/binaryIndex 部件**重编号**（消除编号空洞）；workbook.bin.rels 的 rId 全量重编；</item>
    /// <item>[Content_Types].xml 移除被删表 override 并同步重编号；透视表名称长度字段越界规范化。</item>
    /// </list></summary>
    private static void WriteSurgicalXlsb(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject,
        OoxmlPreservedParts preserved, WorkbookProperties? properties)
    {
        // 先写入内存缓冲，做包结构自检通过后再落盘：避免产出「孤儿 rels / 悬空引用」等
        // 会让 Excel 报修复甚至闪退的损坏文件。自检失败直接抛错，不写坏文件。
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteSurgicalXlsbCore(zip, sheets, vbaProject, preserved, properties);
        }
        buffer.Position = 0;
        PackageIntegrity.ValidateOrThrow(buffer);
        buffer.Position = 0;
        buffer.CopyTo(stream);
    }

    /// <summary>
    /// xlsb 手术式编辑：保留源包全部部件（透视表/切片器/连接/数据模型/宏/绘图等），
    /// 只对「被用户修改过」的工作表做字节级单元格补丁写回其原始 sheetN.bin，其余部件逐字节透传。
    /// 输出与 Excel 自身另存等价，避免整表重建丢失高级部件。
    /// 仅用于「表结构不变（表数/顺序/表名一致）且仅内容修改」的场景。
    /// </summary>
    private static void WriteSurgicalEditXlsb(Stream stream, IReadOnlyList<SheetData> sheets, byte[]? vbaProject,
        OoxmlPreservedParts preserved, WorkbookProperties? properties)
    {
        using var buffer = new MemoryStream();
        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
            WriteSurgicalEditXlsbCore(zip, sheets, vbaProject, preserved, properties);
        buffer.Position = 0;
        PackageIntegrity.ValidateOrThrow(buffer, checkDanglingRels: false);
        buffer.Position = 0;
        buffer.CopyTo(stream);
    }

    private static void WriteSurgicalEditXlsbCore(ZipArchive zip, IReadOnlyList<SheetData> sheets, byte[]? vbaProject,
        OoxmlPreservedParts preserved, WorkbookProperties? properties)
    {
        var vb = preserved.VerbatimBinaries!;

        // 源共享字符串（文本单元格 → isst 索引；找不到则用内联字符串）
        var sst = vb.TryGetValue("xl/sharedStrings.bin", out var sstBin)
            ? XlsbBackend.ParseSharedStrings(sstBin)
            : new List<string>();
        var sstIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < sst.Count; i++)
            if (!sstIndex.ContainsKey(sst[i])) sstIndex[sst[i]] = i;

        // 打开时表序号 → 包内 sheetN.bin 路径（按 workbook.bin 的 BrtBundleSh relId + workbook rels 解析）
        var origRelsXml = preserved.Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var wr) ? wr : "";
        var relIdToTarget = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var rel in XlsxWriter.ParseRels(origRelsXml)) relIdToTarget[rel.Id] = rel.Target;
        var bundleRelIds = new List<string>();
        foreach (var rec in Biff12Records.ReadAll(vb["xl/workbook.bin"]))
            if (rec.Rt == BrtBundleSh && rec.Data.Length >= 8)
            {
                int off = 8;
                bundleRelIds.Add(Biff12Records.ReadWideString(rec.Data, ref off));
            }
        string SheetPath(int origIdx)
        {
            if (origIdx < 0 || origIdx >= bundleRelIds.Count) return "";
            if (!relIdToTarget.TryGetValue(bundleRelIds[origIdx], out var t)) return "";
            return t.StartsWith("/", StringComparison.Ordinal) ? t.TrimStart('/')
                : t.StartsWith("xl/", StringComparison.Ordinal) ? t : "xl/" + t;
        }

        // 补丁被修改的工作表
        var patchedPaths = new HashSet<string>(StringComparer.Ordinal);
        var patchedBins = new List<(string Path, byte[] Data)>();
        foreach (var s in sheets)
        {
            if (!s.IsModified || s.OrigIndex < 0 || s.ModifiedCells is null || s.ModifiedCells.Count == 0) continue;
            var path = SheetPath(s.OrigIndex);
            if (path.Length == 0 || !vb.TryGetValue(path, out var origBin)) continue;
            patchedBins.Add((path, PatchSheetBin(origBin, s, sstIndex)));
            patchedPaths.Add(path);
        }

        // 表格表头同步：若修改的单元格落在某个超级表的表头行，需同步更新该表的列名，
        // 否则表定义与表头单元格不一致会让 Excel 报修复/崩溃。
        var patchedTables = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var s in sheets)
        {
            if (!s.IsModified || s.OrigIndex < 0 || s.ModifiedCells is null || s.ModifiedCells.Count == 0) continue;
            var sheetPath = SheetPath(s.OrigIndex);
            if (sheetPath.Length == 0) continue;
            var relsPath = "xl/worksheets/_rels/" + System.IO.Path.GetFileName(sheetPath) + ".rels";
            if (!preserved.Rels.TryGetValue(relsPath, out var srx)) continue;
            foreach (var rel in XlsxWriter.ParseRels(srx))
            {
                if (!rel.Type.EndsWith("/table", StringComparison.OrdinalIgnoreCase)) continue;
                var abs = ResolveRelsTarget("xl/worksheets", rel.Target);
                if (!preserved.Parts.TryGetValue(abs, out var tblBin)) continue;
                var patchedTbl = PatchTableHeaderCells(tblBin, s);
                if (patchedTbl is not null) patchedTables[abs] = patchedTbl;
            }
        }

        // 新增工作表（仅允许追加在末尾）：生成 sheetN.bin，并记录待追加的 rel / tabId / 路径。
        int openedCount = bundleRelIds.Count;
        var newSheetPaths = new List<string>();
        var newRelIds = new List<string>();
        var newTabIds = new List<int>();
        if (sheets.Count > openedCount)
        {
            var emptySst = new Dictionary<string, int>(StringComparer.Ordinal); // 新表文本用内联字符串，避免改动共享字符串表
            Func<string?, int> xf0 = _ => 0;
            int maxRel = 0;
            foreach (var id in bundleRelIds) { int n = ParseRelId(id); if (n > maxRel) maxRel = n; }
            foreach (var rel in XlsxWriter.ParseRels(origRelsXml)) { int n = ParseRelId(rel.Id); if (n > maxRel) maxRel = n; }
            int maxTab = 0;
            foreach (var rec in Biff12Records.ReadAll(vb["xl/workbook.bin"]))
                if (rec.Rt == BrtBundleSh && rec.Data.Length >= 8) { int t = Biff12Records.ReadS32(rec.Data, 4); if (t > maxTab) maxTab = t; }
            for (int i = openedCount; i < sheets.Count; i++)
            {
                var path = $"xl/worksheets/sheet{i + 1}.bin";
                newSheetPaths.Add(path);
                newRelIds.Add("rId" + (++maxRel));
                newTabIds.Add(++maxTab);
                patchedBins.Add((path, BuildWorksheetBin(sheets[i], emptySst, xf0, false, new List<string>(), null, null)));
                patchedPaths.Add(path);
            }
        }

        // 表名变更 / 新增表：改写 workbook.bin 的 BrtBundleSh
        var patchedWb = PatchWorkbookBinForEdit(vb["xl/workbook.bin"], sheets, openedCount, newRelIds, newTabIds);

        // workbook.bin.rels：若有新增表，追加 worksheet rel
        string? patchedWbRels = null;
        if (newSheetPaths.Count > 0 && preserved.Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var wbRelsXml))
        {
            var sb = new System.Text.StringBuilder();
            sb.Append(wbRelsXml.Substring(0, wbRelsXml.LastIndexOf("</Relationships>", StringComparison.Ordinal)));
            for (int k = 0; k < newSheetPaths.Count; k++)
                sb.Append("<Relationship Id=\"").Append(newRelIds[k])
                  .Append("\" Type=\"").Append(OfficeRelNs).Append("/worksheet\" Target=\"worksheets/")
                  .Append(System.IO.Path.GetFileName(newSheetPaths[k])).Append("\"/>");
            sb.Append("</Relationships>");
            patchedWbRels = sb.ToString();
        }

        // 包结构
        if (preserved.Rels.TryGetValue("_rels/.rels", out var rootRels))
            WriteEntry(zip, "_rels/.rels", System.Text.Encoding.UTF8.GetBytes(rootRels));
        else
            WriteEntry(zip, "_rels/.rels", System.Text.Encoding.UTF8.GetBytes(RootRelsXml(properties is not null)));
        WriteEntry(zip, "[Content_Types].xml", System.Text.Encoding.UTF8.GetBytes(
            ContentTypesXmlFromPreserved(preserved, newSheetPaths)));

        // 其余保留 rels（workbook / worksheet）逐字节透传
        foreach (var kv in preserved.Rels)
        {
            if (kv.Key == "_rels/.rels") continue;
            if (kv.Key == "xl/_rels/workbook.bin.rels" && patchedWbRels is not null)
            {
                WriteEntry(zip, kv.Key, System.Text.Encoding.UTF8.GetBytes(patchedWbRels));
                continue;
            }
            WriteEntry(zip, kv.Key, System.Text.Encoding.UTF8.GetBytes(kv.Value));
        }

        // verbatim 二进制部件（跳过被补丁的表、docProps）
        foreach (var kv in vb)
        {
            if (patchedPaths.Contains(kv.Key)) continue;
            if (kv.Key == "xl/workbook.bin" && patchedWb is not null) continue;
            if (kv.Key == "docProps/core.xml" || kv.Key == "docProps/app.xml") continue;
            WriteEntry(zip, kv.Key, kv.Value);
        }
        if (patchedWb is not null) WriteEntry(zip, "xl/workbook.bin", patchedWb);

        // docProps
        if (properties is not null)
        {
            WriteEntry(zip, "docProps/core.xml", System.Text.Encoding.UTF8.GetBytes(XlsxWriter.CorePropsXml(properties)));
            WriteEntry(zip, "docProps/app.xml", System.Text.Encoding.UTF8.GetBytes(XlsxWriter.AppPropsXml(properties, sheets)));
        }
        else
        {
            if (vb.TryGetValue("docProps/core.xml", out var core)) WriteEntry(zip, "docProps/core.xml", core);
            if (vb.TryGetValue("docProps/app.xml", out var app)) WriteEntry(zip, "docProps/app.xml", app);
        }

        // 其余保留部件（高级部件 / media / 各部件 .rels 等）
        foreach (var kv in preserved.Parts)
        {
            if (patchedTables.TryGetValue(kv.Key, out var pt)) WriteEntry(zip, kv.Key, pt);
            else WriteEntry(zip, kv.Key, kv.Value);
        }

        // vbaProject 兜底
        if (vbaProject is not null && vbaProject.Length > 0
            && !preserved.Parts.ContainsKey("xl/vbaProject.bin") && !vb.ContainsKey("xl/vbaProject.bin"))
            WriteEntry(zip, "xl/vbaProject.bin", vbaProject);

        // 补丁后的工作表（最后写，避免与 vb 重复）
        foreach (var (path, data) in patchedBins)
            WriteEntry(zip, path, data);
    }

    private static string ContentTypesXmlFromPreserved(OoxmlPreservedParts preserved, IReadOnlyList<string>? newSheetPaths = null)
    {
        var sb = new System.Text.StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        foreach (var (ext, ct) in preserved.DefaultTypes)
            sb.Append("<Default Extension=\"").Append(ext).Append("\" ContentType=\"").Append(ct).Append("\"/>");
        foreach (var (part, ct) in preserved.OverrideTypes)
            sb.Append("<Override PartName=\"").Append(part).Append("\" ContentType=\"").Append(ct).Append("\"/>");
        if (newSheetPaths is not null)
            foreach (var p in newSheetPaths)
                sb.Append("<Override PartName=\"/").Append(p).Append("\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
        sb.Append("</Types>");
        return sb.ToString();
    }

    private static bool IsCellRecord(int rt) => rt >= BrtCellBlank && rt <= BrtShortIsst;

    /// <summary>改写 workbook.bin：既有 BrtBundleSh 改名；新增表时在 BrtEndBundleSh 前追加 BrtBundleSh 记录。
    /// 返回 null 表示无需改动。</summary>
    private static byte[]? PatchWorkbookBinForEdit(byte[] wbBin, IReadOnlyList<SheetData> sheets, int openedCount,
        IReadOnlyList<string> newRelIds, IReadOnlyList<int> newTabIds)
    {
        var recs = Biff12Records.ReadAll(wbBin);
        var names = new List<string>();
        foreach (var rec in recs)
        {
            if (rec.Rt != BrtBundleSh || rec.Data.Length < 8) continue;
            int o = 8;
            Biff12Records.ReadWideString(rec.Data, ref o);
            names.Add(Biff12Records.ReadWideString(rec.Data, ref o));
        }
        bool renamed = false;
        for (int i = 0; i < names.Count && i < sheets.Count; i++)
            if (!string.Equals(names[i], sheets[i].SheetName, StringComparison.Ordinal)) { renamed = true; break; }
        bool addSheets = sheets.Count > openedCount;
        if (!renamed && !addSheets) return null;

        using var ms = new MemoryStream();
        int idx = 0;
        foreach (var rec in recs)
        {
            if (rec.Rt == BrtEndBundleShs && addSheets)
            {
                for (int k = openedCount; k < sheets.Count; k++)
                {
                    using var bs = new MemoryStream();
                    WriteU32(bs, 0);                          // hsState = visible
                    WriteU32(bs, (uint)newTabIds[k - openedCount]); // iTabID
                    WriteWideString(bs, newRelIds[k - openedCount]);
                    WriteWideString(bs, sheets[k].SheetName);
                    WriteRecord(ms, BrtBundleSh, bs.ToArray());
                }
            }
            if (rec.Rt == BrtBundleSh && rec.Data.Length >= 8)
            {
                int o = 8;
                Biff12Records.ReadWideString(rec.Data, ref o);
                int nameStart = o;
                var oldName = Biff12Records.ReadWideString(rec.Data, ref o);
                int nameEnd = o;
                var newName = idx < sheets.Count ? sheets[idx].SheetName : oldName;
                idx++;
                if (string.Equals(newName, oldName, StringComparison.Ordinal)) { WriteRecord(ms, rec.Rt, rec.Data); continue; }
                using var msr = new MemoryStream();
                msr.Write(rec.Data, 0, nameStart);
                WriteWideString(msr, newName);
                msr.Write(rec.Data, nameEnd, rec.Data.Length - nameEnd);
                WriteRecord(ms, rec.Rt, msr.ToArray());
                continue;
            }
            WriteRecord(ms, rec.Rt, rec.Data);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// 超级表表头同步：源 tableN.bin 的 BrtBeginList 给出表范围，若被修改的单元格落在表头行（rwFirst）内，
    /// 用工作表模型中的对应单元格文本替换该列的 BrtBeginListCol.stCaption。
    /// 返回 null 表示无需改动（无表头单元格被修改）。
    /// </summary>
    private static byte[]? PatchTableHeaderCells(byte[] tableBin, SheetData sheet)
    {
        // 解析记录
        var recs = new List<(int Rt, int Start, int End, byte[] Data)>();
        int pos = 0;
        while (pos < tableBin.Length)
        {
            int start = pos;
            int rt = Biff12Records.ReadVarInt(tableBin, ref pos);
            int cb = Biff12Records.ReadVarInt(tableBin, ref pos);
            if (cb < 0 || pos + cb > tableBin.Length) break;
            var d = new byte[cb];
            Array.Copy(tableBin, pos, d, 0, cb);
            pos += cb;
            recs.Add((rt, start, pos, d));
        }

        // 表范围
        int rwFirst = -1, colFirst = -1, colLast = -1;
        for (int i = 0; i < recs.Count; i++)
        {
            if (recs[i].Rt != BrtBeginList || recs[i].Data.Length < 16) continue;
            rwFirst = Biff12Records.ReadS32(recs[i].Data, 0);
            colFirst = Biff12Records.ReadS32(recs[i].Data, 8);
            colLast = Biff12Records.ReadS32(recs[i].Data, 12);
            break;
        }
        if (rwFirst < 0 || colFirst < 0) return null;

        // 是否存在落在表头行的被修改单元格
        bool anyHeader = false;
        foreach (var key in sheet.ModifiedCells!)
        {
            int r0 = (int)(key >> 32), c0 = (int)(key & 0xFFFFFFFF);
            if (r0 == rwFirst && c0 >= colFirst && c0 <= colLast) { anyHeader = true; break; }
        }
        if (!anyHeader) return null;

        // 重写每个 BrtBeginListCol 的 stCaption（列序号 = colFirst + 该列在表内 0-based 位置）
        using var ms = new MemoryStream(tableBin.Length + 32);
        int colOrdinal = 0;
        foreach (var (rt, start, end, d) in recs)
        {
            if (rt != BrtBeginListCol)
            {
                WriteRecord(ms, rt, d);
                continue;
            }
            int c0 = colFirst + colOrdinal;
            colOrdinal++;
            string? newName = null;
            if (c0 <= colLast && rwFirst < sheet.Rows.Count && c0 < sheet.Rows[rwFirst].Count)
            {
                var cell = sheet.Rows[rwFirst][c0];
                if (cell.Type == CellType.Text && cell.Text is not null) newName = cell.Text;
            }
            if (newName is null) { WriteRecord(ms, rt, d); continue; }

            // 解析 stName 与 stCaption 的偏移
            int off = 0;
            off += 4; off += 4 * 4; off += 4; // idField + ilta + 3×nDxf + idqsif
            var stName = ReadNullableWideAt(d, ref off);
            int capLenOff = off;
            var stCaption = ReadNullableWideAt(d, ref off);
            int afterCap = off;
            if (stCaption is null || newName == stCaption) { WriteRecord(ms, rt, d); continue; }
            var newCap = NullableWideString(newName);
            var buf = new byte[capLenOff + newCap.Length + (d.Length - afterCap)];
            Array.Copy(d, 0, buf, 0, capLenOff);
            Array.Copy(newCap, 0, buf, capLenOff, newCap.Length);
            Array.Copy(d, afterCap, buf, capLenOff + newCap.Length, d.Length - afterCap);
            WriteRecord(ms, rt, buf);
        }
        return ms.ToArray();
    }

    /// <summary>从 data 的 off 处读取 nullable wide string（cch==0xFFFFFFFF 表示 null）。</summary>
    private static string? ReadNullableWideAt(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return null;
        int cch = Biff12Records.ReadS32(d, off); off += 4;
        if (cch < 0) return null;
        if (off + cch * 2 > d.Length) return null;
        var s = Encoding.Unicode.GetString(d, off, cch * 2);
        off += cch * 2;
        return s;
    }

    /// <summary>
    /// 对源 sheetN.bin 做字节级单元格补丁：只改写/插入/删除「被修改」的单元格记录，
    /// 其余记录（行头、表格部件引用、what-if、打印设置、binaryIndex 引用等）逐字节保留。
    /// </summary>
    private static byte[] PatchSheetBin(byte[] orig, SheetData sheet, Dictionary<string, int> sstIndex)
    {
        var records = new List<(int Rt, int Start, int DataStart, int End, byte[] Data)>();
        int pos = 0;
        while (pos < orig.Length)
        {
            int start = pos;
            int rt = Biff12Records.ReadVarInt(orig, ref pos);
            int cb = Biff12Records.ReadVarInt(orig, ref pos);
            if (cb < 0 || pos + cb > orig.Length) break;
            int ds = pos;
            var d = new byte[cb];
            Array.Copy(orig, ds, d, 0, cb);
            pos += cb;
            records.Add((rt, start, ds, pos, d));
        }

        // 建立 (row0,col0) → 记录下标；并记录每行的 cell 记录区间
        var cellIndex = new Dictionary<long, int>();
        var rowCellStart = new Dictionary<int, int>();   // row0 → 首个 cell 记录下标
        var rowCellEnd = new Dictionary<int, int>();      // row0 → 末个 cell 记录下标 + 1
        var rowHdrIndex = new Dictionary<int, int>();     // row0 → BrtRowHdr 记录下标
        int curRow = -1, prevCol = -1;
        for (int i = 0; i < records.Count; i++)
        {
            var (rt, _, _, _, d) = records[i];
            if (rt == BrtRowHdr)
            {
                if (d.Length >= 4) { curRow = Biff12Records.ReadS32(d, 0); rowHdrIndex[curRow] = i; }
                prevCol = -1;
                continue;
            }
            if (rt == BrtEndSheetData) { curRow = -1; continue; }
            if (curRow < 0 || !IsCellRecord(rt)) continue;
            bool isShort = rt >= BrtShortBlank;
            int col = isShort ? prevCol + 1 : (d.Length >= 4 ? Biff12Records.ReadS32(d, 0) : prevCol + 1);
            if (col < 0) col = prevCol + 1;
            prevCol = col;
            cellIndex[((long)curRow << 32) | (uint)col] = i;
            if (!rowCellStart.ContainsKey(curRow)) rowCellStart[curRow] = i;
            rowCellEnd[curRow] = i + 1;
        }

        // 计算编辑（offset, oldLen, newBytes）；newBytes 为 null 表示删除
        var edits = new List<(int Offset, int OldLen, byte[]? New)>();
        // 新增单元格按「插入锚点」聚合：同一锚点的新增必须合并为一次插入，否则多条记录会互相错位。
        // 每个锚点内按行号、列号升序排列，新建行在该行首个单元格前补一条行头。
        var insertByAnchor = new Dictionary<int, List<(int Row, int Col, byte[] Rec, bool NewRow)>>();
        void AddInsert(int anchor, int row, int col, byte[] rec, bool newRow)
        {
            if (!insertByAnchor.TryGetValue(anchor, out var list)) { list = new List<(int, int, byte[], bool)>(); insertByAnchor[anchor] = list; }
            list.Add((row, col, rec, newRow));
        }
        foreach (var key in sheet.ModifiedCells!)
        {
            int r0 = (int)(key >> 32);
            int c0 = (int)(key & 0xFFFFFFFF);
            Cell? model = (r0 >= 0 && r0 < sheet.Rows.Count && c0 >= 0 && c0 < sheet.Rows[r0].Count)
                ? sheet.Rows[r0][c0] : null;
            bool empty = model is null || model.IsEmpty;

            if (cellIndex.TryGetValue(key, out var ri))
            {
                var rec = records[ri];
                int xf = CellRecordXf(rec.Rt, rec.Data);
                if (empty)
                    edits.Add((rec.Start, rec.End - rec.Start, null));
                else
                    edits.Add((rec.Start, rec.End - rec.Start, EncodeCellRecord(c0, xf, model!, sstIndex)));
            }
            else if (!empty)
            {
                var recBytes = EncodeCellRecord(c0, 0, model!, sstIndex);
                if (rowCellStart.TryGetValue(r0, out var cs) && rowCellEnd.TryGetValue(r0, out var ce))
                {
                    // 已有该行：插入到该行 cell 记录中「按列有序」的位置
                    int at = ce;
                    for (int i = cs; i < ce; i++)
                    {
                        int c = CellRecordCol(records, i);
                        if (c > c0) { at = i; break; }
                    }
                    AddInsert(records[at].Start, r0, c0, recBytes, newRow: false);
                }
                else if (rowHdrIndex.TryGetValue(r0, out var h))
                {
                    AddInsert(records[h].End, r0, c0, recBytes, newRow: false);
                }
                else
                {
                    // 新行：插到「按行有序」的位置（该行需先写一条行头）
                    int at = records.Count;
                    for (int i = 0; i < records.Count; i++)
                    {
                        if (records[i].Rt == BrtEndSheetData) { at = i; break; }
                        if (records[i].Rt == BrtRowHdr && records[i].Data.Length >= 4
                            && Biff12Records.ReadS32(records[i].Data, 0) > r0) { at = i; break; }
                    }
                    AddInsert(records[at].Start, r0, c0, recBytes, newRow: true);
                }
            }
        }
        // 汇总插入：每个锚点内按行、列升序拼接；新建行在其首个单元格前补一条行头。
        var inserted = new List<(int Offset, byte[] New)>();
        foreach (var kv in insertByAnchor)
        {
            int anchor = kv.Key;
            var list = kv.Value;
            list.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Col.CompareTo(b.Col));
            using var ms2 = new MemoryStream();
            int i = 0;
            while (i < list.Count)
            {
                int row = list[i].Row;
                int j = i;
                while (j < list.Count && list[j].Row == row) j++;
                if (list[i].NewRow)
                {
                    int minCol = list[i].Col, maxCol = list[j - 1].Col;
                    WriteRecord(ms2, BrtRowHdr, RowHdr(row, minCol, maxCol, sheet, row));
                }
                for (int k = i; k < j; k++) ms2.Write(list[k].Rec, 0, list[k].Rec.Length);
                i = j;
            }
            inserted.Add((anchor, ms2.ToArray()));
        }

        if (edits.Count == 0 && inserted.Count == 0) return orig;

        // 替换与插入合并为单趟、按 offset 降序应用：低位的替换可能改变记录长度，
        // 若插入另起一趟仍按原始坐标落点，就会提前切断记录。
        var ops = new List<(int Offset, int OldLen, byte[]? New)>(edits.Count + inserted.Count);
        ops.AddRange(edits);
        foreach (var (offset, newBytes) in inserted) ops.Add((offset, 0, newBytes));

        var result = orig;
        foreach (var (offset, oldLen, newBytes) in ops.OrderByDescending(o => o.Offset))
        {
            int newLen = newBytes?.Length ?? 0;
            var buf = new byte[offset + newLen + (result.Length - offset - oldLen)];
            Array.Copy(result, 0, buf, 0, offset);
            if (newBytes is not null) Array.Copy(newBytes, 0, buf, offset, newLen);
            Array.Copy(result, offset + oldLen, buf, offset + newLen, result.Length - offset - oldLen);
            result = buf;
        }
        return result;
    }

    private static int CellRecordCol(List<(int Rt, int Start, int DataStart, int End, byte[] Data)> records, int i)
    {
        var (rt, _, _, _, d) = records[i];
        bool isShort = rt >= BrtShortBlank;
        if (isShort) return -1; // 短记录列号需顺序推导，此处不用于排序判定
        return d.Length >= 4 ? Biff12Records.ReadS32(d, 0) : -1;
    }

    private static int CellRecordXf(int rt, byte[] d)
    {
        bool isShort = rt >= BrtShortBlank;
        int off = isShort ? 0 : 4;
        if (off + 3 > d.Length) return 0;
        return d[off] | (d[off + 1] << 8) | (d[off + 2] << 16);
    }

    /// <summary>把模型单元格编码为一条完整（含记录头）的非短 BIFF12 单元格记录。</summary>
    private static byte[] EncodeCellRecord(int col, int xf, Cell cell, Dictionary<string, int> sstIndex)
    {
        int rt;
        byte[] data;
        switch (cell.Type)
        {
            case CellType.Text:
                if (cell.Text is null) { rt = BrtCellBlank; data = Cell(col, xf); }
                else if (sstIndex.TryGetValue(cell.Text, out var idx)) { rt = BrtCellIsst; data = CellIsst(col, xf, idx); }
                else { rt = BrtCellSt; data = CellSt(col, xf, cell.Text); }
                break;
            case CellType.Boolean:
                rt = BrtCellBool; data = CellBool(col, xf, cell.Boolean);
                break;
            case CellType.Date:
                rt = BrtCellReal; data = CellReal(col, xf, FormatDetector.DateToSerial(cell.Date, false));
                break;
            case CellType.Number:
                double v = cell.Number;
                if (v == Math.Floor(v) && v > -1000 && v < 1000) { rt = BrtCellRk; data = CellRk(col, xf, v); }
                else { rt = BrtCellReal; data = CellReal(col, xf, v); }
                break;
            default:
                rt = BrtCellBlank; data = Cell(col, xf);
                break;
        }
        using var ms = new MemoryStream();
        WriteRecord(ms, rt, data);
        return ms.ToArray();
    }

    private static void WriteSurgicalXlsbCore(ZipArchive zip, IReadOnlyList<SheetData> sheets, byte[]? vbaProject,
        OoxmlPreservedParts preserved, WorkbookProperties? properties)
    {
        var vb = preserved.VerbatimBinaries!;

        // 打开时表数（VerbatimBinaries 中的 sheetN.bin 条目数）
        int openedCount = 0;
        foreach (var key in vb.Keys)
            if (key.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && key.EndsWith(".bin", StringComparison.Ordinal))
                openedCount++;

        var keptOrig = new HashSet<int>();
        foreach (var s in sheets)
            if (s.OrigIndex >= 0) keptOrig.Add(s.OrigIndex);
        var deletedOrig = new List<int>();
        for (int i = 0; i < openedCount; i++)
            if (!keptOrig.Contains(i)) deletedOrig.Add(i);
        if (deletedOrig.Count == 0)
            throw new LiteExcelException("xlsb 手术式删除未检测到被删除的工作表");

        var deletedSheetNums = new HashSet<int>(deletedOrig.Select(i => i + 1));

        var origRelsXml = preserved.Rels.TryGetValue("xl/_rels/workbook.bin.rels", out var r) ? r : "";
        var relIdToTarget = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var rel in XlsxWriter.ParseRels(origRelsXml))
            relIdToTarget[rel.Id] = rel.Target;

        var origWbBin = vb["xl/workbook.bin"];
        var origRecords = Biff12Records.ReadAll(origWbBin);

        // 删除工作表后，存活表公式中指向被删表的 3D 引用必须失效化为 #REF!（Excel 行为）。
        // 否则残留的悬空 3D 引用会让 Excel 打开时崩溃（访问违例）。此处预计算「指向被删表的 XTI 条目下标」。
        var deletedXtiIndexes = ComputeDeletedXtiIndexes(origRecords, deletedOrig);

        // 删除工作表后，workbook.bin 定义名表中「挂被删表」的局部名会被移除（见 ModifyWorkbookBin）。
        // 存活表的单元格公式、数据验证公式（BrtDVal）、表格公式（BrtTableFormula）以 PtgName(0x23)
        // 按「名表 0 基下标」引用这些名；名表缩短后这些下标必须同步重映射，否则残留越界下标会让
        // Excel 打开时报「已修复的记录: ... 部分的 公式」。此处预计算被移除名在原名表中的下标集合。
        var deletedDefinedNameIndexes = ComputeDeletedDefinedNameIndexes(origRecords, deletedSheetNums);

        // workbook.bin 中透视缓存引用记录（0x0182 / 0x046D）按出现顺序 → 缓存部件路径的 0 基索引。
        // 透视表 BrtBeginPivotTable 的 cacheId 即其 rels 指向的缓存在该列表中的位置（真实 Excel 行为）。
        var cacheIndexByTarget = new Dictionary<string, int>(StringComparer.Ordinal);
        int cacheOrder = 0;
        foreach (var rec in origRecords)
        {
            if (rec.Rt != 0x0182 && rec.Rt != 0x046D) continue;
            string relId;
            if (rec.Rt == 0x0182) // flags(u32) + cch(u32) + rId(UTF16)
            {
                int o = 4;
                relId = Biff12Records.ReadWideString(rec.Data, ref o);
            }
            else // 0x046D：flags(u32) + cch(u16) + rId(UTF16) + 尾部(u32)
            {
                if (rec.Data.Length < 6) continue;
                int cch = Biff12Records.ReadU16(rec.Data, 4);
                if (6 + cch * 2 > rec.Data.Length) continue;
                relId = Encoding.Unicode.GetString(rec.Data, 6, cch * 2);
            }
            if (relId.Length == 0 || !relIdToTarget.TryGetValue(relId, out var cacheTarget)) continue;
            var cacheAbs = cacheTarget.StartsWith("/", StringComparison.Ordinal) ? cacheTarget.TrimStart('/')
                : cacheTarget.StartsWith("xl/", StringComparison.Ordinal) ? cacheTarget
                : "xl/" + cacheTarget;
            if (!cacheIndexByTarget.ContainsKey(cacheAbs))
                cacheIndexByTarget[cacheAbs] = cacheOrder++;
        }

        var bundleShList = new List<(string RelId, string Name)>();
        foreach (var rec in origRecords)
        {
            if (rec.Rt == BrtBundleSh && rec.Data.Length >= 8)
            {
                int off = 8;
                var relId = Biff12Records.ReadWideString(rec.Data, ref off);
                var name = Biff12Records.ReadWideString(rec.Data, ref off);
                bundleShList.Add((relId, name));
            }
        }
        int RelIdNum(string id) => ParseRelId(id);
        var deletedRelIdNums = new HashSet<int>();
        foreach (var d in deletedOrig)
            if (d < bundleShList.Count) deletedRelIdNums.Add(RelIdNum(bundleShList[d].RelId));

        // 被删表 rels 引用的部件（tableN.bin 等）+ 被删表 table 编号
        var deletedTableNums = new HashSet<int>();
        var deletedPivotNums = new HashSet<int>();
        var deletedPaths = new HashSet<string>(StringComparer.Ordinal);
        foreach (var d in deletedOrig)
        {
            if (d >= bundleShList.Count) continue;
            if (!relIdToTarget.TryGetValue(bundleShList[d].RelId, out var target)) continue;
            var abs = target.StartsWith("/") ? target.TrimStart('/') : target.StartsWith("xl/") ? target : "xl/" + target;
            deletedPaths.Add(abs);
            var relsPath = "xl/worksheets/_rels/" + System.IO.Path.GetFileName(abs) + ".rels";
            deletedPaths.Add(relsPath);
            if (preserved.Rels.TryGetValue(relsPath, out var srx))
                foreach (var dep in XlsxWriter.ParseRels(srx))
                {
                    var mm = System.Text.RegularExpressions.Regex.Match(dep.Target, @"table(\d+)\.bin");
                    if (mm.Success) deletedTableNums.Add(int.Parse(mm.Groups[1].Value));
                    var mp = System.Text.RegularExpressions.Regex.Match(dep.Target, @"pivotTable(\d+)\.bin");
                    if (mp.Success) deletedPivotNums.Add(int.Parse(mp.Groups[1].Value));
                    var depAbs = ResolveRelsTarget("xl/worksheets", dep.Target);
                    deletedPaths.Add(depAbs);
                    // 关键：依赖部件本体被删时，其自身的 .rels 也必须一并删除。
                    // 否则 .rels 会成为孤儿（父部件已不存在），Excel 打开时报“已修复的记录”或直接崩溃。
                    // 例如被删表带图片/图表 drawing → 删 drawingN.xml 却漏删 xl/drawings/_rels/drawingN.xml.rels。
                    deletedPaths.Add(RelsPathFor(depAbs));
                }
        }

        // 被删透视表引用的透视缓存（pivotCacheDefinitionN.bin）也一并删除（Excel 行为）。
        // 注意：pivotTables/_rels/*.rels 不在 preserved.Rels（仅合并 root/workbook/worksheet rels），而在 preserved.Parts。
        string? PivotRels(string relsPath)
        {
            if (preserved.Rels.TryGetValue(relsPath, out var r1)) return r1;
            if (preserved.Parts.TryGetValue(relsPath, out var b1)) return System.Text.Encoding.UTF8.GetString(b1);
            return null;
        }
        var deletedPivotCacheNums = new HashSet<int>();
        foreach (var pn in deletedPivotNums)
        {
            var ptRels = PivotRels($"xl/pivotTables/_rels/pivotTable{pn}.bin.rels");
            if (ptRels is null) continue;
            foreach (var dep in XlsxWriter.ParseRels(ptRels))
            {
                var depAbs = ResolveRelsTarget("xl/pivotTables", dep.Target);
                var mc = System.Text.RegularExpressions.Regex.Match(depAbs, @"pivotCacheDefinition(\d+)\.bin$");
                if (mc.Success) deletedPivotCacheNums.Add(int.Parse(mc.Groups[1].Value));
                deletedPaths.Add(depAbs);
                deletedPaths.Add(RelsPathFor(depAbs));
            }
        }

        // 逐表出现的编号部件（drawingN / vmlDrawingN / activeXN / printerSettingsN）随表删除而重编号。
        // 这些编号按「源表出现顺序」1 基递增；被删表占用的编号进入各自 deleted 集合，其余编号前移。
        var deletedDrawingNums = new HashSet<int>();
        var deletedVmlNums = new HashSet<int>();
        var deletedActiveXNums = new HashSet<int>();
        var deletedPrinterNums = new HashSet<int>();
        bool disablePartRenumber = Environment.GetEnvironmentVariable("LITEXCEL_DISABLE_PART_RENUMBER") == "1";
        if (!disablePartRenumber)
        {
            int dIdx = 0, vIdx = 0, aIdx = 0, pIdx = 0;
            for (int orig = 0; orig < bundleShList.Count; orig++)
            {
                bool isDeleted = deletedOrig.Contains(orig);
                if (orig >= bundleShList.Count) break;
                if (!relIdToTarget.TryGetValue(bundleShList[orig].RelId, out var tgt)) continue;
                var absSheet = tgt.StartsWith("/") ? tgt.TrimStart('/') : tgt.StartsWith("xl/") ? tgt : "xl/" + tgt;
                var rp = "xl/worksheets/_rels/" + System.IO.Path.GetFileName(absSheet) + ".rels";
                if (!preserved.Rels.TryGetValue(rp, out var sx)) continue;
                foreach (var dep in XlsxWriter.ParseRels(sx))
                {
                    var md = System.Text.RegularExpressions.Regex.Match(dep.Target, @"drawings/drawing(\d+)\.xml$");
                    if (md.Success) { dIdx++; if (isDeleted) deletedDrawingNums.Add(int.Parse(md.Groups[1].Value)); continue; }
                    var mv = System.Text.RegularExpressions.Regex.Match(dep.Target, @"drawings/vmlDrawing(\d+)\.vml$");
                    if (mv.Success) { vIdx++; if (isDeleted) deletedVmlNums.Add(int.Parse(mv.Groups[1].Value)); continue; }
                    var ma = System.Text.RegularExpressions.Regex.Match(dep.Target, @"activeX/activeX(\d+)\.xml$");
                    if (ma.Success) { aIdx++; if (isDeleted) deletedActiveXNums.Add(int.Parse(ma.Groups[1].Value)); continue; }
                    var mpr = System.Text.RegularExpressions.Regex.Match(dep.Target, @"printerSettings/printerSettings(\d+)\.bin$");
                    if (mpr.Success) { pIdx++; if (isDeleted) deletedPrinterNums.Add(int.Parse(mpr.Groups[1].Value)); continue; }
                }
            }
        }

        // 级联删除：被删部件（如 drawingN.xml）的 rels 所引用的部件（如 media/imageN.png）
        // 若不再被任何存活部件引用，也必须一并删除，否则会留下“孤儿部件”。
        // Excel 自身删表时会做此级联；缺失会导致包结构异常。
        if (!disablePartRenumber)
        {
            var allRels = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            void AddRels(string relsPath, string text)
            {
                var src = SourcePartOfRels(relsPath);
                if (src == null) return;
                if (!allRels.TryGetValue(src, out var list)) { list = new List<string>(); allRels[src] = list; }
                var srcDir = src.Contains('/') ? src.Substring(0, src.LastIndexOf('/')) : "";
                foreach (var dep in XlsxWriter.ParseRels(text))
                {
                    if (dep.Target.Contains("://")) continue;
                    list.Add(ResolveRelsTarget(srcDir, dep.Target));
                }
            }
            foreach (var kv in preserved.Rels) AddRels(kv.Key, kv.Value);
            foreach (var kv in preserved.Parts)
                if (kv.Key.EndsWith(".rels", StringComparison.Ordinal))
                    AddRels(kv.Key, System.Text.Encoding.UTF8.GetString(kv.Value));

            bool changed = true;
            while (changed)
            {
                changed = false;
                var referencedByKept = new HashSet<string>(StringComparer.Ordinal);
                foreach (var kv in allRels)
                {
                    if (deletedPaths.Contains(kv.Key)) continue;
                    foreach (var t in kv.Value) referencedByKept.Add(t);
                }
                foreach (var part in deletedPaths.ToList())
                {
                    if (!allRels.TryGetValue(part, out var targets)) continue;
                    foreach (var t in targets)
                    {
                        if (referencedByKept.Contains(t)) continue;
                        if (deletedPaths.Add(t)) { deletedPaths.Add(RelsPathFor(t)); changed = true; }
                    }
                }
            }
        }

        // 被删表上的切片器（slicerN.bin）按名字引用的 slicerCacheN.bin 也随表删除（Excel 行为）。
        // slicer → slicerCache 的关联不在 .rels 里（xl/slicers/_rels 不存在），靠 slicerCache 的
        // Name 字段（如 "切片器_PACKAGEGROUP"）与 slicer 的 Cache 字段匹配。
        // 仅当某 slicerCache 被「被删切片器」引用、且不再被任何存活切片器引用时才删除。
        {
            var deletedSlicerCacheNames = new HashSet<string>(StringComparer.Ordinal);
            var liveSlicerCacheNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (var kv in preserved.Parts)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(kv.Key, @"^xl/slicers/slicer\d+\.bin$")) continue;
                bool isDeleted = deletedPaths.Contains(kv.Key);
                foreach (var sl in Biff12.XlsbSlicerTranscoder.ParseSlicers(kv.Value))
                {
                    if (sl.Cache.Length == 0) continue;
                    if (isDeleted) deletedSlicerCacheNames.Add(sl.Cache);
                    else liveSlicerCacheNames.Add(sl.Cache);
                }
            }
            if (deletedSlicerCacheNames.Count > 0)
            {
                foreach (var kv in preserved.Parts)
                {
                    if (!System.Text.RegularExpressions.Regex.IsMatch(kv.Key, @"^xl/slicerCaches/slicerCache\d+\.bin$")) continue;
                    var info = Biff12.XlsbSlicerTranscoder.ParseCache(kv.Value);
                    if (info.Name.Length == 0) continue;
                    if (deletedSlicerCacheNames.Contains(info.Name) && !liveSlicerCacheNames.Contains(info.Name))
                    {
                        deletedPaths.Add(kv.Key);
                        deletedPaths.Add(RelsPathFor(kv.Key));
                    }
                }
            }
        }

        // 指向被删部件（被删表的依赖、或上面识别出的孤儿切片器缓存）的 workbook.bin.rels 关系：
        // 其对应的缓存引用记录（0x0182/0x046D/0x0430）必须整体移除，rel 本身也从 workbook.bin.rels
        // 剔除、且不参与 rId 重编号。Excel 会连带删除这些“孤儿”缓存；残留会让 Excel 打开时报
        // 「已删除的记录: /xl/slicerCaches/slicerCacheN.bin」并重写 workbook.bin（「已修复的记录: 工作簿属性」）。
        var deletedCacheRelIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rel in XlsxWriter.ParseRels(origRelsXml))
        {
            var tgtAbs = ResolveRelsTarget("xl", rel.Target);
            if (deletedPaths.Contains(tgtAbs)) deletedCacheRelIds.Add(rel.Id);
        }
        foreach (var id in deletedCacheRelIds)
        {
            int n = ParseRelId(id);
            if (n >= 0) deletedRelIdNums.Add(n);
        }

        int NewSheet(int n) => n - deletedSheetNums.Count(x => x < n);
        int NewRel(int n) => n - deletedRelIdNums.Count(x => x < n);
        int NewTable(int n) => n - deletedTableNums.Count(x => x < n);
        int NewPivot(int n) => n - deletedPivotNums.Count(x => x < n);
        int NewPivotCache(int n) => n - deletedPivotCacheNums.Count(x => x < n);
        int NewDrawing(int n) => n - deletedDrawingNums.Count(x => x < n);
        int NewVml(int n) => n - deletedVmlNums.Count(x => x < n);
        int NewActiveX(int n) => n - deletedActiveXNums.Count(x => x < n);
        int NewPrinter(int n) => n - deletedPrinterNums.Count(x => x < n);

        string RenamePartName(string name)
        {
            var m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)sheet(\d+)(\.bin(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}sheet{NewSheet(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)binaryIndex(\d+)(\.bin)$");
            if (m.Success) return $"{m.Groups[1].Value}binaryIndex{NewSheet(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)table(\d+)(\.bin(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}table{NewTable(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)pivotTable(\d+)(\.bin(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}pivotTable{NewPivot(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)pivotCacheDefinition(\d+)(\.bin(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}pivotCacheDefinition{NewPivotCache(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)drawing(\d+)(\.xml(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}drawing{NewDrawing(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)vmlDrawing(\d+)(\.vml(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}vmlDrawing{NewVml(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)activeX(\d+)(\.(xml|bin)(\.rels)?)$");
            if (m.Success) return $"{m.Groups[1].Value}activeX{NewActiveX(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            m = System.Text.RegularExpressions.Regex.Match(name, @"^(.*/)printerSettings(\d+)(\.bin)$");
            if (m.Success) return $"{m.Groups[1].Value}printerSettings{NewPrinter(int.Parse(m.Groups[2].Value))}{m.Groups[3].Value}";
            return name;
        }
        string RenameRefs(string text)
        {
            text = System.Text.RegularExpressions.Regex.Replace(text, @"sheet(\d+)\.bin", m => $"sheet{NewSheet(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"binaryIndex(\d+)\.bin", m => $"binaryIndex{NewSheet(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"table(\d+)\.bin", m => $"table{NewTable(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"pivotTable(\d+)\.bin", m => $"pivotTable{NewPivot(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"pivotCacheDefinition(\d+)\.bin", m => $"pivotCacheDefinition{NewPivotCache(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"drawings/drawing(\d+)\.xml", m => $"drawings/drawing{NewDrawing(int.Parse(m.Groups[1].Value))}.xml");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"drawings/vmlDrawing(\d+)\.vml", m => $"drawings/vmlDrawing{NewVml(int.Parse(m.Groups[1].Value))}.vml");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"activeX/activeX(\d+)\.xml", m => $"activeX/activeX{NewActiveX(int.Parse(m.Groups[1].Value))}.xml");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"activeX(\d+)\.bin", m => $"activeX{NewActiveX(int.Parse(m.Groups[1].Value))}.bin");
            text = System.Text.RegularExpressions.Regex.Replace(text, @"printerSettings/printerSettings(\d+)\.bin", m => $"printerSettings/printerSettings{NewPrinter(int.Parse(m.Groups[1].Value))}.bin");
            return text;
        }

        if (preserved.Rels.TryGetValue("_rels/.rels", out var rootRels))
            WriteEntry(zip, "_rels/.rels", System.Text.Encoding.UTF8.GetBytes(rootRels));
        else
            WriteEntry(zip, "_rels/.rels", System.Text.Encoding.UTF8.GetBytes(RootRelsXml(properties is not null)));

        WriteEntry(zip, "[Content_Types].xml", System.Text.Encoding.UTF8.GetBytes(
            ContentTypesAfterDeleteBinary(sheets.Count, vbaProject is not null, properties is not null, preserved, deletedPaths, deletedTableNums, RenameRefs)));

        WriteEntry(zip, "xl/workbook.bin", ModifyWorkbookBin(origWbBin, deletedOrig, deletedSheetNums, deletedRelIdNums, deletedCacheRelIds, NewSheet, NewRel));

        WriteEntry(zip, "xl/_rels/workbook.bin.rels", System.Text.Encoding.UTF8.GetBytes(
            WorkbookRelsAfterDeleteBinary(origRelsXml, deletedRelIdNums, NewRel, RenameRefs)));

        if (vb.TryGetValue("xl/styles.bin", out var styBin)) WriteEntry(zip, "xl/styles.bin", styBin);
        if (vb.TryGetValue("xl/sharedStrings.bin", out var sstBin)) WriteEntry(zip, "xl/sharedStrings.bin", sstBin);

        if (vbaProject is not null && vbaProject.Length > 0 && !preserved.Parts.ContainsKey("xl/vbaProject.bin"))
            WriteEntry(zip, "xl/vbaProject.bin", vbaProject);

        if (properties is not null)
        {
            WriteEntry(zip, "docProps/core.xml", System.Text.Encoding.UTF8.GetBytes(XlsxWriter.CorePropsXml(properties)));
            WriteEntry(zip, "docProps/app.xml", System.Text.Encoding.UTF8.GetBytes(XlsxWriter.AppPropsXml(properties, sheets)));
        }
        else
        {
            if (vb.TryGetValue("docProps/core.xml", out var core)) WriteEntry(zip, "docProps/core.xml", core);
            if (vb.TryGetValue("docProps/app.xml", out var app)) WriteEntry(zip, "docProps/app.xml", app);
        }

        foreach (var s in sheets)
        {
            if (s.OrigIndex < 0 || s.OrigIndex >= bundleShList.Count) continue;
            if (!relIdToTarget.TryGetValue(bundleShList[s.OrigIndex].RelId, out var target)) continue;
            var abs = target.StartsWith("/") ? target.TrimStart('/') : target.StartsWith("xl/") ? target : "xl/" + target;
            if (vb.TryGetValue(abs, out var sheetBin))
                WriteEntry(zip, RenamePartName(abs), RewriteSheetRefsToDeleted(sheetBin, deletedXtiIndexes, deletedDefinedNameIndexes));
            var relsPath = "xl/worksheets/_rels/" + System.IO.Path.GetFileName(abs) + ".rels";
            if (preserved.Rels.TryGetValue(relsPath, out var sheetRels))
                WriteEntry(zip, RenamePartName(relsPath), System.Text.Encoding.UTF8.GetBytes(RenameRefs(sheetRels)));
        }

        // 其余保留部件重编号后写出；仅 pivotTable 做 cacheId 保真微调，connections 原样透传
        // （`_xlcn.LinkedTable_*` 名是纯元数据、无部件按字符串引用它；Excel 另存时会自行追加 "1"，
        //   库不应改写源名——旧「去尾部 1」会截断合法表名后缀如 Table1→Table，且多次另存逐次降级）。
        foreach (var kv in preserved.Parts)
        {
            if (deletedPaths.Contains(kv.Key)) continue;
            if (kv.Key == "xl/calcChain.bin") continue;
            byte[] data = kv.Value;
            if (kv.Key.StartsWith("xl/pivotTables/pivotTable", StringComparison.Ordinal) && kv.Key.EndsWith(".bin", StringComparison.Ordinal))
            {
                var cacheAbs = PivotTableCacheTarget(kv.Key, preserved);
                int cacheId = cacheAbs is not null && cacheIndexByTarget.TryGetValue(cacheAbs, out var ci) ? ci : -1;
                if (cacheId >= 0) data = NormalizePivotTableBin(data, cacheId);
            }
            // 表格（ListObject）定义内含结构化公式（BrtTableFormula 0x015F），其 3D 引用同样可能在删除表后悬空，
            // 必须与工作表公式一并失效化为 #REF!，否则 Excel 打开崩溃。
            else if (kv.Key.StartsWith("xl/tables/", StringComparison.Ordinal) && kv.Key.EndsWith(".bin", StringComparison.Ordinal))
            {
                data = RewriteTableRefsToDeleted(data, deletedXtiIndexes);
            }
            // 其它 .rels 部件的 Target 中若含被重编号的部件名（如 activeX/activeX2.bin、drawings/drawing3.xml），
            // 也需同步重编号，否则产生悬空引用。
            else if (kv.Key.EndsWith(".rels", StringComparison.Ordinal))
            {
                var txt = System.Text.Encoding.UTF8.GetString(data);
                var renamed = RenameRefs(txt);
                if (!ReferenceEquals(renamed, txt) && renamed != txt)
                    data = System.Text.Encoding.UTF8.GetBytes(renamed);
            }
            WriteEntry(zip, RenamePartName(kv.Key), data);
        }
    }

    /// <summary>改写表格定义二进制：BrtTableFormula(0x015F) 的 rgce 布局为 flags(1)+cce(u32)+rgce；
    /// 把其中指向被删表的 3D 引用令牌（`18 19 <ixti>`）body 失效化为 `10 FF FF FF FF`（Excel 行为）。</summary>
    internal static byte[] RewriteTableRefsToDeleted(byte[] tableBin, HashSet<int> deletedXti)
    {
        if (deletedXti.Count == 0) return tableBin;
        var records = Biff12Records.ReadAll(tableBin);
        bool any = false;
        foreach (var rec in records)
            if (rec.Rt == 0x015F && rec.Data.Length >= 5 && RgceHasDeletedRef(rec.Data, 5, deletedXti)) { any = true; break; }
        if (!any) return tableBin;

        using var ms = new MemoryStream(tableBin.Length);
        foreach (var rec in records)
        {
            if (rec.Rt == 0x015F && rec.Data.Length >= 5)
                WriteRecord(ms, rec.Rt, RewriteRgceAt(rec.Data, 5, deletedXti));
            else
                WriteRecord(ms, rec.Rt, rec.Data);
        }
        return ms.ToArray();
    }

    /// <summary>rgce 位于 data 的 rgceOff 处（此前的 cce(u32) 在 rgceOff-4）。判断是否含指向被删表的 3D 引用。</summary>
    private static bool RgceHasDeletedRef(byte[] d, int rgceOff, HashSet<int> deletedXti)
    {
        int end = d.Length;
        for (int i = rgceOff; i + 6 <= end; i++)
        {
            if (d[i] == 0x18 && d[i + 1] == 0x19)
            {
                if (d[i + 4] == 0x01 && deletedXti.Contains(d[i + 2] | (d[i + 3] << 8))) return true;
                i += 9;
            }
        }
        return false;
    }

    private static byte[] RewriteRgceAt(byte[] d, int rgceOff, HashSet<int> deletedXti)
    {
        var work = (byte[])d.Clone();
        for (int i = rgceOff; i + 6 <= work.Length; i++)
        {
            if (work[i] == 0x18 && work[i + 1] == 0x19)
            {
                if (work[i + 4] == 0x01 && i + 10 <= work.Length
                    && deletedXti.Contains(work[i + 2] | (work[i + 3] << 8)))
                {
                    work[i + 5] = 0x10; work[i + 6] = 0xFF; work[i + 7] = 0xFF;
                    work[i + 8] = 0xFF; work[i + 9] = 0xFF;
                }
                i += 9;
            }
        }
        return work;
    }

    /// <summary>解析 "rIdN" → N（非 rId 前缀或非数字返回 -1）。</summary>
    private static int ParseRelId(string id)
    {
        if (id is null || id.Length < 4 || !id.StartsWith("rId", StringComparison.Ordinal)) return -1;
        return int.TryParse(id.Substring(3), out var n) ? n : -1;
    }

    /// <summary>复刻 Excel「删除工作表后另存」对 workbook.bin 的全部改动：
    /// 移除被删表的 BrtBundleSh 并重编号剩余 BundleSh 的 relId；BrtExternSheet 引用被删表的条目 itab 置 -1、其余递减；
    /// 定义名：sheet-local（挂被删表）移除、其余 itab 递减、引用被删表的 rgce 失效化、`_xlcn.LinkedTable_*` 去尾部 "1"；
    /// 0x0182（BookView 内 rId 引用）重编号；BrtBookView 的 activeTab/firstSheet 调整。</summary>
    private static byte[] ModifyWorkbookBin(byte[] source, List<int> deletedOrigIdx, HashSet<int> deletedSheetNums,
        HashSet<int> deletedRelIdNums, HashSet<string> deletedCacheRelIds, Func<int, int> newSheet, Func<int, int> newRel)
    {
        var records = Biff12Records.ReadAll(source);
        var deletedSet = new HashSet<int>(deletedOrigIdx);

        // 预先计算「引用被删表」的 XTI 条目下标（定义名 rgce 的 ixti 指向这些条目时需失效化）
        var deletedXti = new HashSet<int>();
        foreach (var rec in records)
        {
            if (rec.Rt != 0x016A || rec.Data.Length < 4) continue;
            int cXti = (int)Biff12Records.ReadU32(rec.Data, 0);
            for (int i = 0; i < cXti; i++)
            {
                int off = 4 + i * 12;
                if (off + 12 > rec.Data.Length) break;
                int first = (int)Biff12Records.ReadU32(rec.Data, off + 4);
                int last = (int)Biff12Records.ReadU32(rec.Data, off + 8);
                if ((first >= 0 && deletedSet.Contains(first)) || (last >= 0 && deletedSet.Contains(last)))
                    deletedXti.Add(i);
            }
        }

        using var ms = new MemoryStream(source.Length);
        int bundleIdx = 0;
        int rIdIdx = 0; // 0x0182 序号（flags 规范化）

        for (int ri = 0; ri < records.Count; ri++)
        {
            var rec = records[ri];
            // 缓存引用的 BEGIN 记录（0x0182/0x046D/0x0430）被移除时，其紧邻的 END 伙伴
            // （0x0183/0x046E/0x0431）也必须一并移除，否则残留孤立的 END 记录会让 Excel 拒开。
            void SkipEndIfNext(int endRt)
            {
                if (ri + 1 < records.Count && records[ri + 1].Rt == endRt) ri++;
            }
            if (rec.Rt == BrtBundleSh)
            {
                int cur = bundleIdx++;
                if (deletedSet.Contains(cur)) continue; // 移除被删表
                var d = rec.Data;
                int off = 8;
                var relId = Biff12Records.ReadWideString(d, ref off);
                int num = ParseRelId(relId);
                var newRelStr = "rId" + (num >= 0 ? newRel(num) : num);
                using var b = new MemoryStream();
                b.Write(d, 0, 8);
                var nb = Encoding.Unicode.GetBytes(newRelStr);
                b.Write(BitConverter.GetBytes((uint)newRelStr.Length), 0, 4);
                b.Write(nb, 0, nb.Length);
                b.Write(d, off, d.Length - off);
                WriteRecord(ms, rec.Rt, b.ToArray());
                continue;
            }
            if (rec.Rt == 0x0027 && rec.Data.Length >= 9) // BrtDefinedName
            {
                int itab = (int)Biff12Records.ReadU32(rec.Data, 5);
                if (itab >= 0 && deletedSheetNums.Contains(itab + 1)) continue; // 挂被删表的局部名 → 移除
                var nd = (byte[])rec.Data.Clone();
                if (itab >= 0)
                {
                    int sh = deletedSheetNums.Count(x => x < itab + 1);
                    if (sh > 0) WriteU32To(nd, 5, (uint)(itab - sh));
                }
                nd = RewriteDefinedName(nd, deletedXti);
                WriteRecord(ms, rec.Rt, nd);
                continue;
            }
            if (rec.Rt == 0x009E && rec.Data.Length >= 28) // BrtBookView
            {
                var nd = (byte[])rec.Data.Clone();
                AdjustSheetIndex(nd, 20, deletedOrigIdx);
                AdjustSheetIndex(nd, 24, deletedOrigIdx);
                WriteRecord(ms, rec.Rt, nd);
                continue;
            }
            if (rec.Rt == 0x016A) // BrtExternSheet
            {
                WriteRecord(ms, rec.Rt, ModifyExternSheet(rec.Data, deletedOrigIdx));
                continue;
            }
            if (rec.Rt == 0x0182 && rec.Data.Length >= 8) // BookView 内 rId 引用
            {
                int o0 = 4;
                if (deletedCacheRelIds.Contains(Biff12Records.ReadWideString(rec.Data, ref o0))) { SkipEndIfNext(0x0183); continue; } // 指向被删缓存 → 移除记录
                var nd = (byte[])rec.Data.Clone();
                WriteU32To(nd, 0, (uint)rIdIdx); // flags 规范化为序号（源可能含异常值，如 0x1E）
                int o = 4;
                var id = Biff12Records.ReadWideString(nd, ref o);
                int num = ParseRelId(id);
                if (num >= 0)
                {
                    var newId = "rId" + newRel(num);
                    using var b = new MemoryStream();
                    b.Write(nd, 0, 4);
                    var nb = Encoding.Unicode.GetBytes(newId);
                    b.Write(BitConverter.GetBytes((uint)newId.Length), 0, 4);
                    b.Write(nb, 0, nb.Length);
                    b.Write(nd, o, nd.Length - o);
                    nd = b.ToArray();
                }
                WriteRecord(ms, rec.Rt, nd);
                rIdIdx++;
                continue;
            }
            if ((rec.Rt == 0x046D || rec.Rt == 0x0430) && rec.Data.Length >= 6) // 透视/切片缓存 rId 引用
            {
                // 布局：flags(u32) + cch(u16) + rId(UTF16) [+ 尾部字节]。仅重编号 rId，其余原样保留。
                int cch = Biff12Records.ReadU16(rec.Data, 4);
                int strEnd = 6 + cch * 2;
                if (strEnd <= rec.Data.Length)
                {
                    var rid0 = Encoding.Unicode.GetString(rec.Data, 6, cch * 2);
                    if (deletedCacheRelIds.Contains(rid0)) { SkipEndIfNext(rec.Rt == 0x046D ? 0x046E : 0x0431); continue; } // 指向被删缓存 → 移除记录
                    int num = ParseRelId(rid0);
                    if (num >= 0)
                    {
                        var newId = "rId" + newRel(num);
                        using var b = new MemoryStream();
                        b.Write(rec.Data, 0, 4);
                        var nb = Encoding.Unicode.GetBytes(newId);
                        b.Write(BitConverter.GetBytes((ushort)newId.Length), 0, 2);
                        b.Write(nb, 0, nb.Length);
                        b.Write(rec.Data, strEnd, rec.Data.Length - strEnd);
                        WriteRecord(ms, rec.Rt, b.ToArray());
                        continue;
                    }
                }
                WriteRecord(ms, rec.Rt, rec.Data);
                continue;
            }
            WriteRecord(ms, rec.Rt, rec.Data);
        }
        return ms.ToArray();
    }

    /// <summary>计算指向被删工作表的 XTI 条目下标（BrtExternSheet 中 itabFirst/itabLast ∈ deletedOrig）。
    /// 这些下标对应的 3D 引用在定义名与存活表公式中都必须失效化为 #REF!。</summary>
    private static HashSet<int> ComputeDeletedXtiIndexes(IReadOnlyList<Biff12Records.Record> records, IReadOnlyCollection<int> deletedOrig)
    {
        var deleted = new HashSet<int>();
        var deletedSet = deletedOrig as HashSet<int> ?? new HashSet<int>(deletedOrig);
        foreach (var rec in records)
        {
            if (rec.Rt != 0x016A || rec.Data.Length < 4) continue;
            int cXti = (int)Biff12Records.ReadU32(rec.Data, 0);
            for (int i = 0; i < cXti; i++)
            {
                int off = 4 + i * 12;
                if (off + 12 > rec.Data.Length) break;
                int first = (int)Biff12Records.ReadU32(rec.Data, off + 4);
                int last = (int)Biff12Records.ReadU32(rec.Data, off + 8);
                if ((first >= 0 && deletedSet.Contains(first)) || (last >= 0 && deletedSet.Contains(last)))
                    deleted.Add(i);
            }
            break; // 仅首个 BrtExternSheet
        }
        return deleted;
    }

    /// <summary>计算被删表所挂局部定义名（BrtDefinedName itab 指向被删表）在 workbook.bin 名表中的 0 基下标集合。
    /// 删除工作表后这些名会被移除，名表随之缩短；存活表公式/数据验证/表格公式以 PtgName 按名表下标引用名，
    /// 必须据此集合重映射，否则残留越界下标会让 Excel 打开时报「已修复的记录: ... 部分的 公式」。</summary>
    private static HashSet<int> ComputeDeletedDefinedNameIndexes(IReadOnlyList<Biff12Records.Record> records,
        HashSet<int> deletedSheetNums)
    {
        var deleted = new HashSet<int>();
        int idx = 0;
        foreach (var rec in records)
        {
            if (rec.Rt != 0x0027 || rec.Data.Length < 9) continue;
            int itab = (int)Biff12Records.ReadU32(rec.Data, 5);
            if (itab >= 0 && deletedSheetNums.Contains(itab + 1)) deleted.Add(idx);
            idx++;
        }
        return deleted;
    }

    /// <summary>改写存活工作表二进制：
    /// (1) 把公式 rgce 中指向被删表的 3D 引用令牌失效化为 #REF!（令牌布局：`18 19` + ixti(u16) + 5 字节 body；
    ///     当 ixti 指向被删表的 XTI 条目时把 body 改为 `10 FF FF FF FF`，与 Excel 产出逐字节一致）。
    /// (2) 把单元格公式与数据验证公式（BrtDVal）rgce 中以 PtgName(0x23) 按名表下标引用的名重映射，
    ///     以补偿被删表局部定义名移除导致的名表缩短。
    /// 不做 (1) 会残留悬空 3D 引用（Excel 崩溃 0xc0000005）；不做 (2) 会残留越界名下标
    /// （Excel 打开报「已修复的记录: ... 部分的 公式」）。</summary>
    internal static byte[] RewriteSheetRefsToDeleted(byte[] sheetBin, HashSet<int> deletedXti,
        HashSet<int>? deletedNameIndexes = null)
    {
        bool remapNames = deletedNameIndexes is { Count: > 0 };
        if (deletedXti.Count == 0 && !remapNames) return sheetBin;
        var records = Biff12Records.ReadAll(sheetBin);
        bool any = false;
        foreach (var rec in records)
        {
            if (IsFormulaRecord(rec.Rt) && rec.Data.Length >= 12
                && (deletedXti.Count > 0 && FormulaRefsDeleted(rec.Rt, rec.Data, deletedXti)
                    || remapNames && FormulaHasNameRef(rec.Rt, rec.Data)))
            { any = true; break; }
            if (remapNames && rec.Rt == BrtDVal && DValHasNameRef(rec.Data)) { any = true; break; }
        }
        if (!any) return sheetBin;

        using var ms = new MemoryStream(sheetBin.Length);
        foreach (var rec in records)
        {
            if (IsFormulaRecord(rec.Rt))
            {
                var nd = rec.Data;
                if (deletedXti.Count > 0) nd = RewriteFormulaRefs(rec.Rt, nd, deletedXti);
                if (remapNames) nd = RewriteFormulaNames(rec.Rt, nd, deletedNameIndexes!);
                WriteRecord(ms, rec.Rt, nd);
            }
            else if (remapNames && rec.Rt == BrtDVal)
                WriteRecord(ms, rec.Rt, RewriteDValNames(rec.Data, deletedNameIndexes!));
            else
                WriteRecord(ms, rec.Rt, rec.Data);
        }
        return ms.ToArray();
    }

    private static bool IsFormulaRecord(int rt) => rt == 0x0008 || rt == 0x0009 || rt == 0x000A || rt == 0x000B;

    /// <summary>定位公式记录中 rgce 的偏移（值区之后 2 字节 flags + cce(u32)）。返回 -1 表示布局不符。</summary>
    private static int FormulaRgceOffset(int rt, byte[] d)
    {
        int valueOff = 8;
        int valueLen;
        switch (rt)
        {
            case 0x0008: // BrtFmlaString: cch(u32) + UTF-16
                if (d.Length < 12) return -1;
                uint cch = Biff12Records.ReadU32(d, valueOff);
                valueLen = 4 + (int)cch * 2;
                break;
            case 0x0009: valueLen = 8; break;  // BrtFmlaNum
            case 0x000A: valueLen = 1; break;  // BrtFmlaBool
            case 0x000B: valueLen = 1; break;  // BrtFmlaError
            default: return -1;
        }
        int fOff = valueOff + valueLen;
        if (fOff + 6 > d.Length) return -1;
        int cce = Biff12Records.ReadS32(d, fOff + 2);
        if (cce <= 0 || fOff + 6 + cce > d.Length) return -1;
        return fOff + 6;
    }

    private static bool FormulaRefsDeleted(int rt, byte[] d, HashSet<int> deletedXti)
    {
        int off = FormulaRgceOffset(rt, d);
        if (off < 0) return false;
        int end = d.Length;
        for (int i = off; i + 6 <= end; i++)
        {
            if (d[i] == 0x18 && d[i + 1] == 0x19)
            {
                // 仅当形如 `18 19 <ixti> 01 <4 bytes>`（3D 引用令牌，第 5 字节为 0x01）才算数。
                // 否则 `18 19` 只是恰好出现的普通令牌（如 PtgElfLel），误改会破坏工作表。
                if (d[i + 4] == 0x01)
                {
                    int ixti = d[i + 2] | (d[i + 3] << 8);
                    if (deletedXti.Contains(ixti)) return true;
                }
                i += 9; // 跳过整枚令牌（10 字节）
            }
        }
        return false;
    }

    private static byte[] RewriteFormulaRefs(int rt, byte[] d, HashSet<int> deletedXti)
    {
        int off = FormulaRgceOffset(rt, d);
        if (off < 0) return d;
        var work = (byte[])d.Clone();
        for (int i = off; i + 6 <= work.Length; i++)
        {
            if (work[i] == 0x18 && work[i + 1] == 0x19)
            {
                if (work[i + 4] == 0x01 && i + 10 <= work.Length)
                {
                    int ixti = work[i + 2] | (work[i + 3] << 8);
                    if (deletedXti.Contains(ixti))
                    {
                        // 令牌：18 19 <ixti u16> 01 <body>，body 5 字节改为 10 FF FF FF FF
                        work[i + 5] = 0x10; work[i + 6] = 0xFF; work[i + 7] = 0xFF;
                        work[i + 8] = 0xFF; work[i + 9] = 0xFF;
                    }
                }
                i += 9;
            }
        }
        return work;
    }

    // ===== 定义名下标重映射（PtgName 0x23）=====
    // 删除工作表后，挂被删表的局部定义名从 workbook.bin 名表中移除，名表缩短；
    // 单元格公式、数据验证公式（BrtDVal）以 PtgName 按「名表 0 基下标」引用名，须同步重映射。

    /// <summary>把 0 基名下标 oldIdx 映射为新名表下标（减去其前方被移除的名数量）。</summary>
    private static int RemapNameIndex(int oldIdx, HashSet<int> deletedNameIndexes)
    {
        if (oldIdx < 0) return oldIdx;
        int removed = 0;
        foreach (var d in deletedNameIndexes)
            if (d < oldIdx) removed++;
        return oldIdx - removed;
    }

    /// <summary>按令牌长度遍历 rgce，对每枚 PtgName(0x23) 调用回调（参数为令牌起始偏移）。</summary>
    private static void ForEachPtgName(byte[] d, int start, int end, Action<int> onName)
    {
        int i = start;
        while (i < end)
        {
            byte ptg = d[i];
            int body;
            switch (ptg)
            {
                case 0x17: // PtgStr: cch(u16) + UTF-16
                    if (i + 3 > end) return;
                    body = 2 + (d[i + 1] | (d[i + 2] << 8)) * 2;
                    break;
                case 0x19: body = 3; break;                 // PtgAttr: grbit(u8)+data(u16)
                case 0x1C: case 0x1D: body = 1; break;      // PtgErr / PtgBool
                case 0x1E: body = 2; break;                 // PtgInt
                case 0x1F: body = 8; break;                 // PtgNum
                case 0x20: body = 7; break;                 // PtgArray
                case 0x21: body = 2; break;                 // PtgFunc
                case 0x22: body = 3; break;                 // PtgFuncVar
                case 0x23: body = 4; break;                 // PtgName: index(u32)
                case 0x24: body = 6; break;                 // PtgRef
                case 0x25: body = 12; break;                // PtgArea
                case 0x26: case 0x27: case 0x28: body = 6; break; // MemArea/MemErr/MemNoMem
                case 0x29: body = 4; break;                 // PtgMemFunc
                case 0x2A: body = 6; break;                 // PtgRefErr
                case 0x2B: body = 12; break;                // PtgAreaErr
                case 0x2C: body = 6; break;                 // PtgRefN
                case 0x2D: body = 12; break;                // PtgAreaN
                case 0x2E: case 0x2F: body = 6; break;      // MemAreaN / MemNoMemN
                case 0x39: body = 6; break;                 // PtgNameX: ixti(u16)+index(u32)
                case 0x3A: body = 8; break;                 // PtgRef3d
                case 0x3B: body = 14; break;                // PtgArea3d
                case 0x3C: body = 8; break;                 // PtgRefErr3d
                case 0x3D: body = 14; break;                // PtgAreaErr3d
                default:
                    // ptg > 0x3D 为带类标记的变体（+0x20 / +0x40 等），归一到基类长度。
                    int basePtg = ptg > 0x3D ? (ptg >= 0x60 && ptg <= 0x6F ? ptg - 0x40 : ptg - 0x20) : ptg;
                    switch (basePtg)
                    {
                        case 0x1F: body = 8; break;
                        case 0x22: body = 3; break;
                        case 0x23: body = 4; break;
                        case 0x24: body = 6; break;
                        case 0x25: body = 12; break;
                        case 0x2C: body = 6; break;
                        case 0x2D: body = 12; break;
                        case 0x39: body = 6; break;
                        case 0x3A: body = 8; break;
                        case 0x3B: body = 14; break;
                        case 0x3C: body = 8; break;
                        case 0x3D: body = 14; break;
                        default: body = 0; break;
                    }
                    break;
            }
            if (ptg == 0x23 && i + 5 <= end) onName(i);
            i += 1 + body;
        }
    }

    private static bool FormulaHasNameRef(int rt, byte[] d)
    {
        int off = FormulaRgceOffset(rt, d);
        if (off < 0) return false;
        int cce = Biff12Records.ReadS32(d, off - 4);
        bool found = false;
        ForEachPtgName(d, off, off + cce, _ => found = true);
        return found;
    }

    private static byte[] RewriteFormulaNames(int rt, byte[] d, HashSet<int> deletedNameIndexes)
    {
        int off = FormulaRgceOffset(rt, d);
        if (off < 0) return d;
        int cce = Biff12Records.ReadS32(d, off - 4);
        var work = (byte[])d.Clone();
        ForEachPtgName(work, off, off + cce, i =>
        {
            int old = Biff12Records.ReadS32(work, i + 1);
            int nu = RemapNameIndex(old, deletedNameIndexes);
            if (nu != old) WriteU32To(work, i + 1, (uint)nu);
        });
        return work;
    }

    /// <summary>BrtDVal 是否含 PtgName(0x23) 引用（用于快速跳过无需改写的记录）。</summary>
    private static bool DValHasNameRef(byte[] d)
    {
        foreach (var (off, len) in DValRgceRanges(d))
        {
            bool found = false;
            ForEachPtgName(d, off, off + len, _ => found = true);
            if (found) return true;
        }
        return false;
    }

    /// <summary>重映射 BrtDVal 中 rgce1/rgce2 的 PtgName 下标。</summary>
    private static byte[] RewriteDValNames(byte[] d, HashSet<int> deletedNameIndexes)
    {
        var work = (byte[])d.Clone();
        foreach (var (off, len) in DValRgceRanges(work))
        {
            ForEachPtgName(work, off, off + len, i =>
            {
                int old = Biff12Records.ReadS32(work, i + 1);
                int nu = RemapNameIndex(old, deletedNameIndexes);
                if (nu != old) WriteU32To(work, i + 1, (uint)nu);
            });
        }
        return work;
    }

    /// <summary>定位 BrtDVal 记录内 rgce1/rgce2 的 (偏移, 长度)。布局：
    /// flags(4)+cRefs(4)+refs(16×cRefs)+promptTitle/prompt/errorTitle/errorMessage(可空宽串)
    /// +cce1(4)+rgce1+reserved(4)+cce2(4)+rgce2+reserved(4)。布局不符返回空。</summary>
    private static List<(int Off, int Len)> DValRgceRanges(byte[] d)
    {
        var res = new List<(int, int)>();
        if (d.Length < 8) return res;
        int cRefs = (int)Biff12Records.ReadU32(d, 4);
        int o = 8 + 16 * cRefs;
        if (cRefs < 0 || o > d.Length) return res;
        for (int k = 0; k < 4; k++) // 4 个可空宽串
        {
            if (o + 4 > d.Length) return res;
            int cch = Biff12Records.ReadS32(d, o); o += 4;
            if (cch >= 0) o += cch * 2;
            if (o > d.Length) return res;
        }
        if (o + 4 > d.Length) return res;
        int cce1 = Biff12Records.ReadS32(d, o); o += 4;
        if (cce1 < 0 || o + cce1 > d.Length) return res;
        res.Add((o, cce1)); o += cce1;
        o += 4; // reserved
        if (o + 4 > d.Length) return res;
        int cce2 = Biff12Records.ReadS32(d, o); o += 4;
        if (cce2 < 0 || o + cce2 > d.Length) return res;
        res.Add((o, cce2));
        return res;
    }

    /// <summary>改写定义名：rgce 引用被删表时失效化（名称本身原样保留）。
    /// 库不应改写源名（旧「去尾部 1」会截断合法表名后缀 Table1→Table，且多次另存逐次降级）。
    /// rgce 布局（真实样本标定）：`18 19` + ixti(u16) + body(10)；当 ixti 指向已失效的 XTI 条目时，
    /// 把 body 的 `00 xx 00 00 00` 改为失效标记 `10 FF FF FF FF`（与 Excel 产出逐字节一致）。</summary>
    private static byte[] RewriteDefinedName(byte[] data, HashSet<int> deletedXti)
    {
        var work = data;
        int nameEnd = 9;
        int cch = (int)Biff12Records.ReadU32(work, nameEnd);
        if (cch <= 0 || cch == -1) return work;
        nameEnd += 4 + cch * 2;
        if (nameEnd + 4 > work.Length) return work;
        int cce = (int)Biff12Records.ReadU32(work, nameEnd);
        int rgceOff = nameEnd + 4;
        if (cce >= 14 && rgceOff + 14 <= work.Length && work[rgceOff] == 0x18 && work[rgceOff + 1] == 0x19)
        {
            int ixti = work[rgceOff + 2] | (work[rgceOff + 3] << 8);
            if (deletedXti.Contains(ixti))
            {
                work[rgceOff + 5] = 0x10; work[rgceOff + 6] = 0xFF; work[rgceOff + 7] = 0xFF;
                work[rgceOff + 8] = 0xFF; work[rgceOff + 9] = 0xFF;
            }
        }
        return work;
    }

    /// <summary>解析 rels Target 为包内规范化绝对路径（处理 "../" 与 "xl/" 前缀）。
    /// baseDir 为 rels 所在部件目录（如 "xl/worksheets"）。</summary>
    internal static string ResolveRelsTarget(string baseDir, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal))
            return target.TrimStart('/');
        var combined = baseDir.Length == 0 ? target : baseDir + "/" + target;
        var segs = combined.Split('/');
        var stack = new List<string>(segs.Length);
        foreach (var seg in segs)
        {
            if (seg.Length == 0 || seg == ".") continue;
            if (seg == "..")
            {
                if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                continue;
            }
            stack.Add(seg);
        }
        return string.Join("/", stack);
    }

    /// <summary>返回包内某部件的 .rels 路径（如 "xl/drawings/drawing3.xml" → "xl/drawings/_rels/drawing3.xml.rels"）。
    /// 部件路径为包内规范化绝对路径（不含前导 '/'）。</summary>
    internal static string RelsPathFor(string partPath)
    {
        int slash = partPath.LastIndexOf('/');
        var dir = slash >= 0 ? partPath.Substring(0, slash) : "";
        var file = slash >= 0 ? partPath.Substring(slash + 1) : partPath;
        return dir.Length == 0 ? "_rels/" + file + ".rels" : dir + "/_rels/" + file + ".rels";
    }

    /// <summary>RelsPathFor 的逆运算：由 ".rels" 路径求其描述的源部件路径。
    /// 根 rels（"_rels/.rels"）无源部件，返回 null。</summary>
    internal static string? SourcePartOfRels(string relsPath)
    {
        var segs = relsPath.Split('/');
        for (int i = segs.Length - 1; i >= 0; i--)
        {
            if (segs[i] == "_rels")
            {
                if (i == 0) return null;
                if (i != segs.Length - 2) return null;
                var dir = string.Join("/", segs, 0, i);
                var file = segs[segs.Length - 1];
                if (!file.EndsWith(".rels", StringComparison.Ordinal)) return null;
                var src = file.Substring(0, file.Length - 5);
                return dir.Length == 0 ? src : dir + "/" + src;
            }
        }
        return null;
    }

    /// <summary>修正 BrtExternSheet：**条目数量与顺序保持不变**。对每个 XTI 条目：
    /// <list type="bullet">
    /// <item>itabFirst/itabLast 引用被删表 → 置为 -1（0xFFFFFFFF）；</item>
    /// <item>itab 大于被删表索引 → 递减（减去其前方被删表数量）；</item>
    /// <item>其余（含 -1/-2 特殊值）保持不变。</item>
    /// </list>
    /// 这是 Excel 自身的真实行为（2026-09-15 用 Excel COM 标定确认）：删除被命名区域引用的工作表后，
    /// XTI 条目**不移除**，被删表对应的 itab 变为 -1；命名区域的 rgce ixti **不重映射**，
    /// 而是通过该 XTI 的 -1 状态在 Excel 中呈现为 `#REF!`。
    /// 早期实现"整条移除 + 左移 + 重映射 rgce"会破坏 ixti 与 XTI 的对应关系，导致 Excel 崩溃。</summary>
    private static byte[] ModifyExternSheet(byte[] data, List<int> deletedOrigIdx)
    {
        if (data.Length < 4) return data;
        int cXti = (int)Biff12Records.ReadU32(data, 0);
        var deletedSet = new HashSet<int>(deletedOrigIdx);
        var result = (byte[])data.Clone();
        for (int i = 0; i < cXti; i++)
        {
            int off = 4 + i * 12;
            if (off + 12 > data.Length) break;
            int first = (int)Biff12Records.ReadU32(data, off + 4);
            int last = (int)Biff12Records.ReadU32(data, off + 8);

            if ((first >= 0 && deletedSet.Contains(first)) || (last >= 0 && deletedSet.Contains(last)))
            {
                // 引用被删表 → itab 置 -1（保留条目位置，不左移）
                WriteU32To(result, off + 4, 0xFFFFFFFF);
                WriteU32To(result, off + 8, 0xFFFFFFFF);
                continue;
            }

            // 非被删条目：itab > 被删索引者递减（保持与原表的对应关系）
            if (first >= 0)
            {
                int shiftFirst = deletedOrigIdx.Count(d => d < first);
                if (shiftFirst > 0) WriteU32To(result, off + 4, (uint)(first - shiftFirst));
            }
            if (last >= 0)
            {
                int shiftLast = deletedOrigIdx.Count(d => d < last);
                if (shiftLast > 0) WriteU32To(result, off + 8, (uint)(last - shiftLast));
            }
        }
        return result;
    }

    /// <summary>调整 u32 字段中的 sheet 索引：==deleted→0, >deleted→递减, <deleted→不变</summary>
    private static void AdjustSheetIndex(byte[] data, int offset, List<int> deletedOrigIdx)
    {
        uint val = Biff12Records.ReadU32(data, offset);
        if (val == 0xFFFFFFFF) return; // -1 = 无效/未设置
        int ival = (int)val;
        if (ival < 0) return;
        if (deletedOrigIdx.Contains(ival))
        {
            // active tab 被删 → 指向第一张保留表
            var u32 = 0u;
            data[offset] = (byte)(u32 & 0xFF);
            data[offset + 1] = (byte)((u32 >> 8) & 0xFF);
            data[offset + 2] = (byte)((u32 >> 16) & 0xFF);
            data[offset + 3] = (byte)((u32 >> 24) & 0xFF);
        }
        else
        {
            int shift = deletedOrigIdx.Count(d => d < ival);
            if (shift > 0)
            {
                var u32 = (uint)(ival - shift);
                data[offset] = (byte)(u32 & 0xFF);
                data[offset + 1] = (byte)((u32 >> 8) & 0xFF);
                data[offset + 2] = (byte)((u32 >> 16) & 0xFF);
                data[offset + 3] = (byte)((u32 >> 24) & 0xFF);
            }
        }
    }

    /// <summary>重建 [Content_Types].xml：剔除被删表 override + 其独占部件 override，其余保留并对剩余部件重编号。</summary>
    private static string ContentTypesAfterDeleteBinary(int sheetCount, bool hasVba, bool hasProps,
        OoxmlPreservedParts preserved, HashSet<string> deletedPaths, HashSet<int> deletedTableNums, Func<string, string> renameRefs)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");

        var seenExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "bin", "rels", "vml", "xml" };
        if (preserved.VerbatimBinaries is not null)
        {
            foreach (var (ext, ct) in preserved.DefaultTypes)
                if (!seenExt.Contains(ext) && seenExt.Add(ext))
                    sb.Append($"<Default Extension=\"{ext}\" ContentType=\"{ct}\"/>");
        }

        var deletedOverrides = new HashSet<string>(StringComparer.Ordinal);
        foreach (var p in deletedPaths)
            deletedOverrides.Add("/" + p);
        foreach (var t in deletedTableNums)
            deletedOverrides.Add($"/xl/tables/table{t}.bin");

        // 从 preserved.OverrideTypes 中保留未被删的条目（含保留表的 sheet/binaryIndex override），并重编号。
        // 剔除：被删表及其依赖部件、陈旧 calcChain。
        var seenPart = new HashSet<string>(StringComparer.Ordinal);
        if (preserved.VerbatimBinaries is not null)
        {
            foreach (var (part, ct) in preserved.OverrideTypes)
            {
                if (deletedOverrides.Contains(part)) continue;
                if (part == "/xl/calcChain.bin") continue;
                var renamed = "/" + renameRefs(part.TrimStart('/'));
                if (seenPart.Add(renamed))
                    sb.Append($"<Override PartName=\"{renamed}\" ContentType=\"{ct}\"/>");
            }
        }

        // 确保关键公共部件 override 存在（workbook / styles / sharedStrings / vba / docProps）
        seenPart.Add("/xl/workbook.bin");
        sb.Append("<Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
        if (seenPart.Add("/xl/styles.bin"))
            sb.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
        if (preserved.VerbatimBinaries is not null && preserved.VerbatimBinaries.ContainsKey("xl/sharedStrings.bin"))
        {
            if (seenPart.Add("/xl/sharedStrings.bin"))
                sb.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
        }
        if (hasVba)
        {
            if (seenPart.Add("/xl/vbaProject.bin"))
                sb.Append("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
        }
        if (hasProps)
        {
            if (seenPart.Add("/docProps/core.xml"))
                sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            if (seenPart.Add("/docProps/app.xml"))
                sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }

        sb.Append("</Types>");
        return sb.ToString();
    }

    /// <summary>重建 workbook.bin.rels：剔除被删表/陈旧 calcChain 关系，剩余 rId 全量重编号、Target 同步重编号。</summary>
    private static string WorkbookRelsAfterDeleteBinary(string origRels, HashSet<int> deletedRelIdNums,
        Func<int, int> newRel, Func<string, string> renameRefs)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        foreach (var rel in XlsxWriter.ParseRels(origRels))
        {
            int num = ParseRelId(rel.Id);
            if (num >= 0 && deletedRelIdNums.Contains(num)) continue;
            if (rel.Target.IndexOf("calcChain.bin", StringComparison.OrdinalIgnoreCase) >= 0) continue;
            var newId = num >= 0 ? "rId" + newRel(num) : rel.Id;
            var newTarget = renameRefs(rel.Target);
            sb.Append($"<Relationship Id=\"{newId}\" Type=\"{rel.Type}\" Target=\"{newTarget}\"");
            if (rel.TargetMode.Length > 0) sb.Append($" TargetMode=\"{rel.TargetMode}\"");
            sb.Append("/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    /// <summary>保真微调 pivotTableN.bin：把 BrtBeginPivotTable(0x0118) 的 cacheId(off28) 归位为其 rels
    /// 指向的透视缓存在 workbook.bin 缓存引用序列中的 0 基索引。源文件该字段常为脏值（Excel 打开时按 rels 重算），
    /// 保持原值会被 Excel 判为断链并删除透视表。布局：... + off24(u32) + off28(u32 cacheId) + off32(u32 cch) + name。</summary>
    private static byte[] NormalizePivotTableBin(byte[] data, int cacheId)
    {
        var recs = Biff12Records.ReadAll(data);
        using var ms = new MemoryStream(data.Length);
        foreach (var rec in recs)
        {
            if (rec.Rt == 0x0118 && rec.Data.Length >= 36)
            {
                var nd = (byte[])rec.Data.Clone();
                if (Biff12Records.ReadU32(nd, 28) != (uint)cacheId)
                    WriteU32To(nd, 28, (uint)cacheId);
                WriteRecord(ms, rec.Rt, nd);
                continue;
            }
            WriteRecord(ms, rec.Rt, rec.Data);
        }
        return ms.ToArray();
    }

    /// <summary>解析透视表部件 rels，返回其引用的透视缓存部件包内绝对路径（如 "xl/pivotCache/pivotCacheDefinition3.bin"）；无则 null。</summary>
    private static string? PivotTableCacheTarget(string pivotPartPath, OoxmlPreservedParts preserved)
    {
        var dir = pivotPartPath.Substring(0, pivotPartPath.LastIndexOf('/'));
        var fileName = pivotPartPath.Substring(pivotPartPath.LastIndexOf('/') + 1);
        var relsPath = dir + "/_rels/" + fileName + ".rels";
        string? relsXml = null;
        if (preserved.Rels.TryGetValue(relsPath, out var r)) relsXml = r;
        else if (preserved.Parts.TryGetValue(relsPath, out var bytes)) relsXml = Encoding.UTF8.GetString(bytes);
        if (relsXml is null) return null;
        foreach (var rel in XlsxWriter.ParseRels(relsXml))
            if (rel.Type.EndsWith("/pivotCacheDefinition", StringComparison.OrdinalIgnoreCase))
                return ResolveRelsTarget(dir, rel.Target);
        return null;
    }

    /// <summary>合并工作表级保留 rels（图表/透视表等）与重建的超链接 rels </summary>
    private static string? BuildSheetRelsXml(int sheetNumber, List<string> extLinks, OoxmlPreservedParts? preserved, bool hasComments = false, SheetData? sheet = null, int tableIdBase = -1,
        bool hasDrawing = false, bool hasNewFloating = false, XlsxWriter.ImagePlan? imagePlan = null, int sheetIndex = 0)
    {
        var relParts = new List<XlsxWriter.RelInfo>();
        int nextRid = 1;
        for (int k = 0; k < extLinks.Count; k++)
        {
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rIdH{k + 1}",
                Type = $"{OfficeRelNs}/hyperlink",
                Target = extLinks[k],
                TargetMode = "External",
            });
            nextRid = k + 2;
        }
        if (hasComments)
        {
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rIdC1",
                Type = $"{OfficeRelNs}/comments",
                Target = $"../comments{sheetNumber}.bin",
            });
            relParts.Add(new XlsxWriter.RelInfo
            {
                Id = $"rIdV1",
                Type = $"{OfficeRelNs}/vmlDrawing",
                Target = $"../drawings/vmlDrawing{sheetNumber}.vml",
            });
        }
        // 超级表 rels（rIdT1.. → ../tables/tableN.bin，N 为全局表编号）
        if (sheet?.Tables is { Count: > 0 } && tableIdBase >= 0)
        {
            for (int t = 0; t < sheet.Tables.Count; t++)
            {
                relParts.Add(new XlsxWriter.RelInfo
                {
                    Id = "rIdT" + (t + 1),
                    Type = $"{OfficeRelNs}/table",
                    Target = $"../tables/table{tableIdBase + t}.bin",
                });
            }
        }
        // 浮动图片 drawing rel（新建 drawing 用 rIdD1；既有保留 drawing 的 rel 由 MergeRelsXml 保留）
        if (hasDrawing && imagePlan is not null)
        {
            var (drawingEntry, drawingRelId) = imagePlan.DrawingTargetFor(sheetIndex, preserved);
            bool hasExistingDrawingRel = false;
            if (preserved is not null
                && preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetNumber}.bin.rels", out var origRels))
            {
                foreach (var rel in XlsxWriter.ParseRels(origRels))
                    if (rel.Type.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase)
                        && XlsxWriter.ResolveRelsTarget("xl/worksheets", rel.Target) == drawingEntry)
                    { hasExistingDrawingRel = true; break; }
            }
            if (!hasExistingDrawingRel)
            {
                relParts.Add(new XlsxWriter.RelInfo
                {
                    Id = drawingRelId,
                    Type = $"{OfficeRelNs}/drawing",
                    Target = $"../drawings/{drawingEntry.Substring(drawingEntry.LastIndexOf('/') + 1)}",
                });
            }
        }
        string rebuilt = XlsxWriter.RelsXml(relParts);
        string original = "";
        if (preserved is not null && preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetNumber}.bin.rels", out var r))
            original = r;
        // 源里的 table 关系需被重建产物替换：否则旧的 rId 关系会与重建的并存，
        // 且与工作表记录中的 BrtTablePart 引用不一致。
        var rebuiltTargets = new HashSet<string>(StringComparer.Ordinal);
        if (sheet?.Tables is { Count: > 0 } && !string.IsNullOrEmpty(original))
        {
            foreach (var rel in XlsxWriter.ParseRels(original))
                if (rel.Type.EndsWith("/table", StringComparison.OrdinalIgnoreCase))
                    rebuiltTargets.Add(XlsxWriter.ResolveRelsTarget("xl/worksheets", rel.Target));
        }
        return XlsxWriter.MergeRelsXml(original, "xl/worksheets", rebuiltTargets, rebuilt);
    }

    /// <summary>xlsb 目标写出时，源为 XML-OOXML（xlsx/xlsm）的保留部件（透视表/切片器/连接/表格/查询表等）
    /// 其 XML 部件结构与 BIFF12 不兼容，无法透传。经降级回调显式上报（避免静默丢失）。 </summary>
    private static void ReportCrossFormatPreservedDrop(Action<DegradationInfo>? onDegradation, ExcelFormat targetFormat,
        OoxmlPreservedParts preserved)
    {
        if (onDegradation is null) return;
        int count = preserved.Parts.Count(kv => kv.Key.StartsWith("xl/", StringComparison.Ordinal));
        if (count == 0) return;
        onDegradation(new DegradationInfo
        {
            Capability = DegradationCapability.PivotTables,
            TargetFormat = targetFormat,
            Message = $"源为 XML-OOXML（xlsx/xlsm）格式，其包含的 {count} 个高级保留部件（透视表/切片器/查询表/外部连接等 XML 部件）" +
                      "与 xlsb（BIFF12 二进制容器）不兼容，已丢弃。转换后这些高级功能不可用；如需保留请在原格式下保存。"
        });
    }

    /// <summary>写出 xlsb 时逐项上报不支持的能力。
    /// namedRanges / documentProperties 为工作簿级，SheetName 为 null。</summary>
    private static void ReportDegradations(IReadOnlyList<SheetData> sheets,
        IReadOnlyList<NamedRange>? names, WorkbookProperties? properties,
        Action<DegradationInfo>? onDegradation, ExcelFormat targetFormat)
    {
        if (onDegradation is null) return;
        ReportWorkbook(names, properties, onDegradation, targetFormat);
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
            if (sheet.Filter is not null && sheet.Filter.Columns.Count > 0)
                Report(DegradationCapability.AutoFilter, $"xlsb 自动筛选仅支持范围写出，工作表 '{sheet.SheetName}' 的筛选条件已丢弃。");
            // 浮动图片已支持写出（drawing/media 与 xlsx 同为 XML 部件，经 BrtDrawing 引用）。
            // 仅 InCell 图片（richData 体系）在 xlsb 不支持，显式上报。
            if (sheet.Images is { Count: > 0 } && sheet.Images.Any(i => i.Placement == ImagePlacement.InCell))
                Report(DegradationCapability.Images, $"xlsb 不支持单元格内嵌图片（InCell richData），工作表 '{sheet.SheetName}' 的内嵌图片已丢弃。");
            if (DegradationDetector.HasNonNumberFormatStyles(sheet))
                Report(DegradationCapability.Styles, $"xlsb 仅支持数字格式，工作表 '{sheet.SheetName}' 的完整样式（字体/颜色/边框/对齐/换行）已降级。");
            // 条件格式：全部 18 种 OOXML 规则类型均经 WriteConditionalFormats 写出，无丢弃。
            if (sheet.ConditionalFormats is { Count: > 0 })
            {
                int unsupported = 0;
                foreach (var cf in sheet.ConditionalFormats)
                    if (!IsCfTypeSupported(cf.Type)) unsupported++;
                if (unsupported > 0)
                    Report(DegradationCapability.ConditionalFormatting, $"xlsb 条件格式存在 {unsupported} 条未知类型规则，工作表 '{sheet.SheetName}' 的此类规则已丢弃。");
            }
            // 超级表不再上报：三条写出路径（重建 / verbatim / 手术式）均保留超级表
            // （重建经 BuildTableBin 写出，verbatim/手术式原样透传 xl/tables/*.bin），不存在丢弃路径。
            if (!string.IsNullOrEmpty(sheet.TabColor))
                Report(DegradationCapability.SheetVisibility, $"xlsb 不支持工作表标签颜色（tabColor），工作表 '{sheet.SheetName}' 的标签颜色已丢弃。");
        }
    }

    /// <summary>工作簿级静默丢失：命名区域仅读取不上报（root），文档属性已写为 OOXML 但也従工具的范围限制来看需要提醒。</summary>
    private static void ReportWorkbook(IReadOnlyList<NamedRange>? names, WorkbookProperties? properties,
        Action<DegradationInfo> onDegradation, ExcelFormat targetFormat)
    {
        void Report(DegradationCapability cap, string msg)
            => onDegradation(new DegradationInfo
            {
                Capability = cap,
                TargetFormat = targetFormat,
                Message = msg,
            });
        if (names is { Count: > 0 })
            Report(DegradationCapability.NamedRanges, $"xlsb 不支持命名区域（definedNames），工作簿的 {names.Count} 个命名区域已静默丢弃。");
        if (HasMeaningfulProperties(properties))
            Report(DegradationCapability.DocumentProperties, "xlsb 不支持文档属性（Title/Subject/Creator/...）,工作簿属性已静默丢弃。");
    }

    /// <summary>判断文档属性是否被填充。 </summary>
    private static bool HasMeaningfulProperties(WorkbookProperties? p)
        => p is not null && (p.Creator is not null || p.Title is not null || p.Subject is not null
            || p.LastModifiedBy is not null || p.Created is not null || p.Modified is not null || p.Application is not null);

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

    private static string ContentTypesXml(int sheetCount, bool hasSst, bool hasVba, bool hasProps, OoxmlPreservedParts? preserved, IReadOnlyList<int>? sheetsWithComments = null, int tableCount = 0, XlsxWriter.ImagePlan? imagePlan = null)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        for (int i = 1; i <= sheetCount; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>");
        sb.Append("<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>");
        for (int t = 1; t <= tableCount; t++)
            sb.Append($"<Override PartName=\"/xl/tables/table{t}.bin\" ContentType=\"application/vnd.ms-excel.table\"/>");
        if (hasSst)
            sb.Append("<Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>");
        if (sheetsWithComments is not null)
        {
            foreach (var idx in sheetsWithComments)
                sb.Append($"<Override PartName=\"/xl/comments{idx}.bin\" ContentType=\"application/vnd.ms-excel.comments\"/>");
        }
        if (hasVba)
            sb.Append("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
        if (hasProps)
        {
            sb.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            sb.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
        }
        // 图片：media 类型 Default + drawing Override（与 xlsx 一致；drawing/media 在 xlsb 中同为 XML 部件）
        if (imagePlan is { Any: true })
        {
            var imgSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "rels", "xml", "bin", "vml" };
            foreach (var img in imagePlan.All)
            {
                if (imgSeen.Add(img.EffectiveExtension))
                {
                    string ct = img.EffectiveExtension switch
                    {
                        "png" => "image/png",
                        "jpg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        _ => "application/octet-stream",
                    };
                    sb.Append($"<Default Extension=\"{img.EffectiveExtension}\" ContentType=\"{ct}\"/>");
                }
            }
            for (int i = 0; i < imagePlan.FloatingBySheet.Count; i++)
            {
                if (imagePlan.FloatingBySheet[i].Count > 0)
                    sb.Append($"<Override PartName=\"/{imagePlan.DrawingTargetFor(i, preserved).Entry}\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            }
        }
        // 合并保留部件的 content types 声明（仅源为 xlsb 时）。
        if (preserved is not null && preserved.VerbatimBinaries is not null)
        {
            foreach (var (ext, ct) in preserved.DefaultTypes)
                if (ext != "bin" && ext != "rels" && ext != "xml" && !sb.ToString().Contains($"Extension=\"{ext}\""))
                    sb.Append($"<Default Extension=\"{ext}\" ContentType=\"{ct}\"/>");
            foreach (var (part, ct) in preserved.OverrideTypes)
                if (!part.StartsWith("/xl/worksheets/") && !part.StartsWith("/xl/workbook") && !part.StartsWith("/xl/styles")
                    && !part.StartsWith("/xl/sharedStrings") && part != "/xl/vbaProject.bin"
                    && !part.Equals("/xl/calcChain.bin")
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
        rebuiltTargets.Add("xl/calcChain.bin"); // 陈旧计算链不透传，其关系一并移除。

        // 删除表后：原 workbook rels 中所有 worksheet 类型的 rel（含已被删除而未重建的表，如 sheet10.bin）
        // 必须剔除，否则产生孤儿 rel，Excel 打开时会触发「文件级验证和修复」。
        if (!string.IsNullOrEmpty(original))
        {
            foreach (var rel in XlsxWriter.ParseRels(original))
            {
                if (rel.Type.EndsWith("/worksheet", StringComparison.Ordinal))
                {
                    var abs = XlsxWriter.ResolveRelsTarget("xl", rel.Target);
                    rebuiltTargets.Add(abs);
                }
            }
        }
        return XlsxWriter.MergeRelsXml(original, "xl", rebuiltTargets, rebuilt) ?? rebuilt;
    }

    private static void WriteEntry(ZipArchive zip, string name, byte[] data)
    {
        // Excel 对数据模型部件 xl/model/item.data 约定为 STORED（不压缩）：它是内存映射数据库，
        // 压缩无收益且影响随机访问。与 Excel 自身产出保持一致。
        var level = name == "xl/model/item.data" ? CompressionLevel.NoCompression : CompressionLevel.Optimal;
        var entry = zip.CreateEntry(name, level);
        using var s = entry.Open();
        s.Write(data, 0, data.Length);
    }

    private static void WriteEntry(ZipArchive zip, string name, string xml)
    {
        WriteEntry(zip, name, new UTF8Encoding(false).GetBytes(xml));
    }

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
            WriteRecord(ms, BrtBundleSh, BundleSh(i, sheets[i].SheetName,
                SheetVisibilityMap.FromOoxml(sheets[i].SheetState)));
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

    private static byte[] BundleSh(int index, string name, SheetVisibility visibility = SheetVisibility.Visible)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)SheetVisibilityMap.ToBiff(visibility)); // hsState: 0=visible/1=hidden/2=veryHidden
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

    private static byte[] BuildStylesBin(List<(int Ifmt, string? FmtCode)> cellXfs, IReadOnlyList<CellStyle>? dxfs = null)
    {
        var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginStyleSheet, Array.Empty<byte>());

        // 自定义数字格式（使用 cellXfs 中已分配的唯一 ifmt）
        var customFormats = new List<(int Id, string Code)>();
        var seenFmtIds = new HashSet<int>();
        foreach (var (ifmt, fmtCode) in cellXfs)
        {
            if (string.IsNullOrEmpty(fmtCode)) continue;
            if (ifmt >= FirstCustomFmtId && seenFmtIds.Add(ifmt))
                customFormats.Add((ifmt, fmtCode!));
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

        // dxfs（条件格式样式 / 超级表列格式）
        WriteRecord(ms, BrtBeginDXFs, UInt32((uint)(dxfs?.Count ?? 0)));
        if (dxfs is not null)
            foreach (var style in dxfs)
                WriteRecord(ms, BrtDXF, BuildDxf(style));
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

    private static byte[] BuildWorksheetBin(SheetData sheet, Dictionary<string, int> sstIndex, Func<string?, int> getXf, bool date1904, List<string> extTargets, string? drawingRelId = null, DxfRegistry? dxfRegistry = null)
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
        WriteRecord(ms, BrtBeginSheetData, new byte[] { 0xFF, 0xFF, 0xFF, 0xFF, 0x08, 0x00, 0x1D, 0x01, 0x00, 0x00, 0x00, 0x00 });
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

        // 自动筛选（范围写出；复杂条件降级上报）
        if (sheet.Filter is not null && !string.IsNullOrEmpty(sheet.Filter.Range))
        {
            var (afRwFirst, afColFirst, afRwLast, afColLast) = CellRef.ParseRange(sheet.Filter.Range);
            WriteRecord(ms, BrtBeginAFilter, RfX(afRwFirst, afRwLast, afColFirst, afColLast));
            WriteRecord(ms, BrtEndAFilter, Array.Empty<byte>());
        }

        // 数据验证（BrtBeginDVs / BrtDVal×N / BrtEndDVs）
        if (sheet.Validations is { Count: > 0 })
        {
            WriteRecord(ms, BrtBeginDVs, BeginDVs(sheet.Validations.Count));
            foreach (var dv in sheet.Validations)
                WriteRecord(ms, BrtDVal, BuildDVal(dv));
            WriteRecord(ms, BrtEndDVs, Array.Empty<byte>());
        }

        // 合并单元格
        if (sheet.MergedRanges.Count > 0)
        {
            WriteRecord(ms, BrtBeginMergeCells, UInt32((uint)sheet.MergedRanges.Count));
            foreach (var m in sheet.MergedRanges)
                WriteRecord(ms, BrtMergeCell, RfX(m.FirstRow, m.LastRow, m.FirstCol, m.LastCol));
            WriteRecord(ms, BrtEndMergeCells, Array.Empty<byte>());
        }

        // 条件格式（conditionalFormatting）：schema 位于 mergeCells 之后、hyperlinks 之前
        if (dxfRegistry is not null && sheet.ConditionalFormats is { Count: > 0 })
            WriteConditionalFormats(ms, sheet, dxfRegistry.GetOrCreate);

        // 超链接（外部经 relId 指向 sheet rels；内部走 location）
        WriteHyperlinks(ms, sheet, extRelIndex);

        // 绘图引用（图片/图表/形状）——BrtDrawing 引用 sheet rels 中的 rId
        if (!string.IsNullOrEmpty(drawingRelId))
            WriteRecord(ms, BrtDrawing, WideString(drawingRelId));

        // 超级表部件引用（BrtTablePart 引用 sheet rels 中的 rId）
        if (sheet.Tables is { Count: > 0 })
        {
            var tableRids = TableRelIds(sheet);
            WriteRecord(ms, BrtBeginTableParts, UInt32((uint)tableRids.Count));
            foreach (var rid in tableRids)
                WriteRecord(ms, BrtTablePart, WideString(rid));
            WriteRecord(ms, BrtEndTableParts, Array.Empty<byte>());
        }

        WriteRecord(ms, BrtEndSheet, Array.Empty<byte>());
        return ms.ToArray();
    }

    /// <summary>表部件在 sheet rels 中使用的 rId（确定性：rIdT1, rIdT2, ...）。</summary>
    internal static List<string> TableRelIds(SheetData sheet)
    {
        var list = new List<string>();
        if (sheet.Tables is null) return list;
        for (int i = 0; i < sheet.Tables.Count; i++)
            list.Add("rIdT" + (i + 1));
        return list;
    }

    /// <summary>构建 tableN.bin（BIFF12 超级表部件）。</summary>
    private static byte[] BuildTableBin(XlTable table, int id)
    {
        var (rwF, rwL, colF, colL) = ParseRef(table.Ref);
        var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginList, BeginList(rwF, rwL, colF, colL, id, table));
        // 自动筛选范围（与表范围一致）
        WriteRecord(ms, BrtBeginAFilter, RfX(rwF, rwL, colF, colL));
        WriteRecord(ms, BrtEndAFilter, Array.Empty<byte>());
        WriteRecord(ms, BrtBeginListCols, UInt32((uint)table.Columns.Count));
        for (int c = 0; c < table.Columns.Count; c++)
        {
            WriteRecord(ms, BrtBeginListCol, BeginListCol(table.Columns[c], c + 1));
            WriteRecord(ms, 0x015C, Array.Empty<byte>()); // BrtEndListCol
        }
        WriteRecord(ms, BrtEndListCols, Array.Empty<byte>());
        WriteRecord(ms, BrtTableStyleClient, TableStyleClient(table));
        WriteRecord(ms, BrtEndList, Array.Empty<byte>());
        return ms.ToArray();
    }

    /// <summary>A1 范围 → (rwFirst, rwLast, colFirst, colLast)，0-based。</summary>
    private static (int rwF, int rwL, int colF, int colL) ParseRef(string refA1)
    {
        var parts = refA1.Split(':');
        var (r1, c1) = ParseA1(parts[0]);
        var (r2, c2) = parts.Length > 1 ? ParseA1(parts[1]) : (r1, c1);
        return (r1, r2, c1, c2);
    }

    private static (int row, int col) ParseA1(string a1)
    {
        int i = 0, col = 0;
        while (i < a1.Length && char.IsLetter(a1[i])) { col = col * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1); i++; }
        int row = 0;
        while (i < a1.Length && char.IsDigit(a1[i])) { row = row * 10 + (a1[i] - '0'); i++; }
        return (row - 1, col - 1);
    }

    private static byte[] BeginList(int rwF, int rwL, int colF, int colL, int id, XlTable table)
    {
        var ms = new MemoryStream();
        WriteS32(ms, rwF);
        WriteS32(ms, rwL);
        WriteS32(ms, colF);
        WriteS32(ms, colL);
        WriteS32(ms, 0);              // lt
        WriteS32(ms, id);             // idList
        WriteS32(ms, table.TotalsRowShown ? 0 : 1); // crwHeader
        WriteS32(ms, table.TotalsRowShown ? 1 : 0); // crwTotals
        WriteS32(ms, 0);              // flags
        for (int i = 0; i < 6; i++) WriteS32(ms, -1); // nDxfHeader/Data/Agg/Border/HeaderBorder/AggBorder
        WriteS32(ms, 0);              // dwConnID
        ms.Write(NullableWideString(table.Name), 0, NullableWideString(table.Name).Length); // stName
        ms.Write(NullableWideString(table.Name), 0, NullableWideString(table.Name).Length); // stDisplayName
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);             // stComment
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);             // stStyleHeader
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);             // stStyleData
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);             // stStyleAgg
        return ms.ToArray();
    }

    private static byte[] BeginListCol(XlTableColumn col, int idField)
    {
        var ms = new MemoryStream();
        WriteS32(ms, idField); // idField（1-based 列序号）
        WriteS32(ms, 0);   // ilta
        WriteS32(ms, -1);  // nDxfHdr
        WriteS32(ms, -1);  // nDxfInsertRow
        WriteS32(ms, -1);  // nDxfAgg
        WriteS32(ms, 0);   // idqsif
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);            // stName
        ms.Write(NullableWideString(col.Name), 0, NullableWideString(col.Name).Length);    // stCaption
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);            // stTotal
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);            // stStyleHeader
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);            // stStyleInsertRow
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length);            // stStyleAgg
        return ms.ToArray();
    }

    private static byte[] TableStyleClient(XlTable table)
    {
        var ms = new MemoryStream();
        int flags = 0;
        if (table.ShowFirstColumn) flags |= 0x01;
        if (table.ShowLastColumn) flags |= 0x02;
        if (table.ShowRowStripes) flags |= 0x04;
        if (table.ShowColumnStripes) flags |= 0x08;
        WriteU16(ms, (ushort)flags);
        ms.Write(NullableWideString(table.StyleName), 0, NullableWideString(table.StyleName).Length);
        return ms.ToArray();
    }

    /// <summary>BrtBeginDVs：reserved(4) + ... + cDVs(4)（对齐 Excel 样本 18 字节：前 12 字节 0，第 14 字节起 cDVs）。</summary>
    private static byte[] BeginDVs(int count)
    {
        var ms = new MemoryStream();
        WriteS32(ms, 0);       // reserved
        WriteS32(ms, 0);       // reserved
        WriteS32(ms, 0);       // reserved
        WriteU16(ms, 0);       // reserved
        WriteS32(ms, count);   // cDVs
        return ms.ToArray();
    }

    /// <summary>
    /// BrtDVal：flags(4) + cRefs(4) + refs + promptTitle/prompt/errorTitle/errorMessage（nullable wide）
    /// + cce(4)+rgce[f1] + cce(4)+rgce[f2] + reserved(4)。
    /// flags 位0-3=类型(1=whole/2=decimal/3=list/4=date/6=textLength)，位16=allowBlank，
    /// 位17=InCellDropdown，位18=ShowInputMessage，位19=ShowErrorMessage。
    /// </summary>
    private static byte[] BuildDVal(DataValidation dv)
    {
        int typeCode = dv.Type switch
        {
            DataValidationType.WholeNumber => 1,
            DataValidationType.Decimal => 2,
            DataValidationType.List => 3,
            DataValidationType.Date => 4,
            _ => 3,
        };
        int flags = typeCode;
        if (dv.Type == DataValidationType.List) flags |= 1 << 7; // InCellDropdown（列表）
        if (dv.AllowBlank) flags |= 1 << 8;                      // allowBlank
        flags |= 1 << 18; // ShowInputMessage
        flags |= 1 << 19; // ShowErrorMessage

        var (rwF, colF, rwL, colL) = CellRef.ParseRange(dv.Sqref);

        var ms = new MemoryStream();
        WriteS32(ms, flags);
        WriteS32(ms, 1);        // cRefs
        WriteS32(ms, rwF); WriteS32(ms, rwL); WriteS32(ms, colF); WriteS32(ms, colL);
        ms.Write(NullableWideString(dv.PromptTitle), 0, NullableWideString(dv.PromptTitle).Length);
        ms.Write(NullableWideString(dv.Prompt), 0, NullableWideString(dv.Prompt).Length);
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length); // errorTitle
        ms.Write(NullableWideString(null), 0, NullableWideString(null).Length); // errorMessage

        var f1 = EncodeDvFormula(dv.Formula1, dv.Type);
        WriteS32(ms, f1.Length); ms.Write(f1, 0, f1.Length);
        WriteS32(ms, 0); // reserved（rgce1 与 rgce2 之间）
        var f2 = EncodeDvFormula(dv.Formula2, dv.Type);
        WriteS32(ms, f2.Length); ms.Write(f2, 0, f2.Length); // rgce2（无 f2 时 cce=0）
        WriteS32(ms, 0); // reserved（记录尾部）
        return ms.ToArray();
    }

    /// <summary>数据验证公式 → rgce。列表为字符串常量；数值/日期为 PtgInt/PtgNum。</summary>
    private static byte[] EncodeDvFormula(string formula, DataValidationType type)
    {
        if (string.IsNullOrEmpty(formula)) return Array.Empty<byte>();
        // 列表：去除两端引号后按 PtgStr 编码
        if (type == DataValidationType.List)
        {
            var s = formula;
            if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"') s = s.Substring(1, s.Length - 2);
            var ms = new MemoryStream();
            ms.WriteByte(0x17); // PtgStr
            WriteU16(ms, (ushort)s.Length);
            var b = Encoding.Unicode.GetBytes(s);
            ms.Write(b, 0, b.Length);
            return ms.ToArray();
        }
        // 数值/日期：PtgInt 为 u16（0..65535），超出用 PtgNum
        if (double.TryParse(formula, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num))
        {
            var ms = new MemoryStream();
            if (num == Math.Floor(num) && num >= 0 && num <= 65535)
            {
                ms.WriteByte(0x1E); // PtgInt
                WriteU16(ms, (ushort)(int)num);
            }
            else
            {
                ms.WriteByte(0x1F); // PtgNum
                var b = BitConverter.GetBytes(num);
                ms.Write(b, 0, b.Length);
            }
            return ms.ToArray();
        }
        return Array.Empty<byte>();
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

        var formulaText = cell.Formula ?? (cell.IsFormula ? cell.Text : null);
        if (!string.IsNullOrEmpty(formulaText))
        {
            var rpn = FormulaEncoder.TryEncode(formulaText, biff12: true);
            if (rpn is not null)
            {
                WriteFormulaCell(ms, col, cell, xf, rpn, lastSeen);
                prevCol = col;
                return;
            }
        }

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

    private static void WriteFormulaCell(MemoryStream ms, int col, Cell cell, int xf, byte[] rpn, bool lastSeen)
    {
        // BIFF12 BrtFmla* records: col(4) + ixfe(3) + padding(1) + value + reserved(2) + cce(4) + RPN + ctrlExp(4)
        var data = new MemoryStream();
        WriteS32(data, col);
        data.WriteByte((byte)(xf & 0xFF));
        data.WriteByte((byte)((xf >> 8) & 0xFF));
        data.WriteByte((byte)((xf >> 16) & 0xFF));
        data.WriteByte(0x00); // padding to align value at offset 8 (same as non-formula cell records)

        switch (cell.Type)
        {
            case CellType.Number:
                WriteDouble(data, cell.Number);
                WriteU16(data, 0); // reserved
                WriteU32(data, (uint)rpn.Length);
                data.Write(rpn, 0, rpn.Length);
                WriteU32(data, 0); // ctrlExp (fField=0, reserved=0)
                WriteRecord(ms, BrtFmlaNum, data.ToArray());
                break;
            case CellType.Date:
                WriteDouble(data, FormatDetector.DateToSerial(cell.Date, false));
                WriteU16(data, 0);
                WriteU32(data, (uint)rpn.Length);
                data.Write(rpn, 0, rpn.Length);
                WriteU32(data, 0); // ctrlExp
                WriteRecord(ms, BrtFmlaNum, data.ToArray());
                break;
            case CellType.Boolean:
                data.WriteByte((byte)(cell.Boolean ? 1 : 0));
                WriteU16(data, 0);
                WriteU32(data, (uint)rpn.Length);
                data.Write(rpn, 0, rpn.Length);
                WriteU32(data, 0); // ctrlExp
                WriteRecord(ms, BrtFmlaBool, data.ToArray());
                break;
            default:
                // BrtFmlaString: value = XLWideString (cch(4) + chars)
                WriteU32(data, 0); // cch = 0 (empty string result)
                WriteU16(data, 0); // reserved
                WriteU32(data, (uint)rpn.Length);
                data.Write(rpn, 0, rpn.Length);
                WriteU32(data, 0); // ctrlExp
                WriteRecord(ms, BrtFmlaString, data.ToArray());
                break;
        }
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

    /// <summary>XLWideString 字节（cch(4) + UTF-16LE），用于记录 payload。</summary>
    private static byte[] WideString(string s)
    {
        var ms = new MemoryStream();
        WriteWideString(ms, s);
        return ms.ToArray();
    }

    /// <summary>XLNullableWideString 字节（cch(4)，空串写 0xFFFFFFFF）。</summary>
    private static byte[] NullableWideString(string? s)
    {
        var ms = new MemoryStream();
        if (string.IsNullOrEmpty(s)) WriteU32(ms, 0xFFFFFFFF);
        else WriteWideString(ms, s);
        return ms.ToArray();
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

    private const int BrtBeginComments = 0x0274;
    private const int BrtEndComments = 0x0275;
    private const int BrtBeginCommentAuthors = 0x0276;
    private const int BrtEndCommentAuthors = 0x0277;
    private const int BrtCommentAuthor = 0x0278;
    private const int BrtBeginCommentList = 0x0279;
    private const int BrtBeginComment = 0x0025;
    private const int BrtUid = 0x0C00;
    private const int BrtBeginCommentText = 0x0026;
    private const int BrtCommentText = 0x027D;
    private const int BrtEndCommentText = 0x027C;
    private const int BrtEndComment = 0x027A;
    private const int BrtCommentRichValue = 0x027B;

    /// <summary>
    /// 构建 commentsN.bin 的 BIFF12 记录流。
    /// 结构：BeginComments > BeginCommentAuthors > Author > EndCommentAuthors >
    /// BeginCommentList > [BeginComment > Uid > BeginCommentText > RichValue > CommentText > EndCommentText > EndComment]* > EndComments
    /// </summary>
    private static byte[] BuildCommentsBin(IReadOnlyDictionary<string, string> comments)
    {
        var ms = new MemoryStream();

        WriteRecord(ms, BrtBeginComments, Array.Empty<byte>());
        WriteRecord(ms, BrtBeginCommentAuthors, Array.Empty<byte>());

        var authorBytes = new MemoryStream();
        WriteWideString(authorBytes, "LiteExcel");
        WriteRecord(ms, BrtCommentAuthor, authorBytes.ToArray());

        WriteRecord(ms, BrtEndCommentAuthors, Array.Empty<byte>());
        WriteRecord(ms, BrtBeginCommentList, Array.Empty<byte>());

        int shapeId = 1025;
        foreach (var kv in comments)
        {
            var (row, col) = CellRef.Parse(kv.Key);

            // BrtBeginComment: row(4) + col(4) + authorId(4) + shapeId(4) + flags(2)
            var bcData = new MemoryStream();
            WriteS32(bcData, row);
            WriteS32(bcData, col);
            WriteU32(bcData, 0); // authorId
            WriteU32(bcData, (uint)shapeId);
            WriteU16(bcData, 0x0080); // flags
            WriteRecord(ms, BrtBeginComment, bcData.ToArray());

            // BrtUid: 16-byte GUID
            var guid = Guid.NewGuid().ToByteArray();
            WriteRecord(ms, BrtUid, guid);

            // BrtBeginCommentText
            WriteRecord(ms, BrtBeginCommentText, Array.Empty<byte>());

            // BrtCommentRichValue: 36 bytes (from real sample)
            var rvData = new byte[36];
            rvData[0] = 0x00;
            WriteRecord(ms, BrtCommentRichValue, rvData);

            // BrtCommentText: flags(1) + XLWideString
            var ctData = new MemoryStream();
            ctData.WriteByte(0x01); // flags
            WriteWideString(ctData, kv.Value);
            WriteRecord(ms, BrtCommentText, ctData.ToArray());

            // BrtEndCommentText
            WriteRecord(ms, BrtEndCommentText, Array.Empty<byte>());

            // BrtEndComment
            WriteRecord(ms, BrtEndComment, Array.Empty<byte>());

            shapeId++;
        }

        WriteRecord(ms, BrtEndComments, Array.Empty<byte>());
        return ms.ToArray();
    }
}
