using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using LiteExcel.Internal;
using LiteExcel.Internal.Cfb;
namespace LiteExcel.Internal.Biff;

/// <summary>
/// 传统 .xls（BIFF8）写入后端。
/// 从对象模型 SheetData 生成 Workbook 流（全局子流 + 各工作表子流），再包进 OLE2/CFB 容器。
/// 公式单元格降级为静态值（按缓存值写出）。
/// </summary>
internal static class XlsWriter
{
    private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);

    // 全局
    private const ushort OpBof = 0x0809;
    private const ushort OpEof = 0x000A;
    private const ushort OpInterfaceHdr = 0x00E1;
    private const ushort OpMms = 0x00C1;
    private const ushort OpInterfaceEnd = 0x00E2;
    private const ushort OpWriteAccess = 0x005C;
    private const ushort OpCodePage = 0x0042;
    private const ushort OpDsf = 0x0161;
    private const ushort OpTabId = 0x013D;
    private const ushort OpFnGroupCount = 0x009C;
    private const ushort OpWindowProtect = 0x0019;
    private const ushort OpProtect = 0x0012;
    private const ushort OpPassword = 0x0013;
    private const ushort OpProt4Rev = 0x01AF;
    private const ushort OpBoundSheet = 0x0085;
    private const ushort OpSst = 0x00FC;
    private const ushort OpContinue = 0x003C;
    private const ushort OpFormat = 0x041E;
    private const ushort OpFont = 0x0031;
    private const ushort OpXf = 0x00E0;
    private const ushort OpCountry = 0x008C;
    private const ushort OpWindow1 = 0x01C0;

    // 工作表
    private const ushort OpDimensions = 0x0200;
    private const ushort OpWindow2 = 0x023E;
    private const ushort OpPane = 0x0041;
    private const ushort OpColInfo = 0x007D;
    private const ushort OpRow = 0x0208;
    private const ushort OpMergedCells = 0x00E5;
    private const ushort OpLabelSst = 0x00FD;
    private const ushort OpNumber = 0x0203;
    private const ushort OpRk = 0x027E;
    private const ushort OpBoolErr = 0x0205;
    private const ushort OpHlink = 0x01B8;
    private const ushort OpHlinkTooltip = 0x0800;

    private const int MaxRecordData = 8192; // 保守的记录数据上限（规范为 8224 总长）

    /// <summary>Excel 必需的内置数字/日期格式（ifmt &lt; 164）。无自定义格式时也须写出，否则 Excel 拒开。</summary>
    private static readonly (int Id, string Code)[] BuiltInFormats =
    {
        (5, "0"),
        (6, "0.00"),
        (7, "#,##0"),
        (8, "#,##0.00"),
        (23, "m/d/yy"),
        (24, "d-mmm-yy"),
        (25, "d-mmm"),
        (26, "mmm-yy"),
        (27, "h:mm AM/PM"),
        (28, "h:mm:ss AM/PM"),
        (29, "h:mm"),
        (30, "h:mm:ss"),
        (31, "m/d/yy h:mm"),
    };

    public static void Write(Stream stream, IReadOnlyList<SheetData> sheets, bool date1904 = false)
    {
        var workbook = BuildWorkbookStream(sheets, date1904);
        var cfb = CfbWriter.Build("Workbook", workbook);
        stream.Write(cfb, 0, cfb.Length);
    }

    private static byte[] BuildWorkbookStream(IReadOnlyList<SheetData> sheets, bool date1904)
    {
        // ── 预扫描：SST 唯一字符串、格式→XF 表 ──
        var sst = new List<string>();
        var sstIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        var formats = new List<string>(); // formats[k] -> ifmt = 164 + k，cell XF 索引 = 16 + k
        var xfIndex = new Dictionary<string, int>(StringComparer.Ordinal); // 格式码 -> cell XF 索引

        int GetXf(string? fmtCode)
        {
            if (string.IsNullOrEmpty(fmtCode)) return 16; // XF 16 = General 单元格样式
            if (xfIndex.TryGetValue(fmtCode, out var idx)) return idx;
            idx = 16 + formats.Count + 1;
            formats.Add(fmtCode);
            xfIndex[fmtCode] = idx;
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
                case CellType.Date:
                case CellType.Number:
                    GetXf(cell.NumberFormat);
                    break;
            }
        }

        foreach (var sheet in sheets)
            foreach (var row in sheet.Rows)
                foreach (var cell in row)
                    ScanCell(cell);

        // ── 全局子流（BOUNDSHEET 位置先占位 0，随后原地打补丁） ──
        // 记录顺序与内容对齐 Excel/SheetJS 写出
        var global = new MemoryStream();
        var boundFieldOffsets = new int[sheets.Count];
        WriteRecord(global, OpBof, Bof(0x0005)); // workbook globals
        WriteRecord(global, OpInterfaceHdr, new byte[] { 0xB0, 0x04 });
        WriteRecord(global, OpMms, new byte[] { 0, 0 });
        WriteRecord(global, OpInterfaceEnd, Array.Empty<byte>());
        WriteRecord(global, OpWriteAccess, WriteAccess());
        WriteRecord(global, OpCodePage, new byte[] { 0xB0, 0x04 }); // 1200
        WriteRecord(global, OpDsf, new byte[] { 0, 0 });
        WriteRecord(global, OpWindow1, Array.Empty<byte>());
        // TabId：每个工作表 2 字节
        var tabId = new byte[2 * sheets.Count];
        for (int i = 0; i < sheets.Count; i++) { tabId[i * 2] = 1; tabId[i * 2 + 1] = 0; }
        WriteRecord(global, OpTabId, tabId);
        WriteRecord(global, OpFnGroupCount, new byte[] { 0x11, 0 }); // 17
        WriteRecord(global, OpWindowProtect, new byte[] { 0, 0 });
        WriteRecord(global, OpProtect, new byte[] { 0, 0 });
        WriteRecord(global, OpPassword, new byte[] { 0, 0 });
        WriteRecord(global, 0x01AF, new byte[] { 0, 0 });             // Prot4Rev
        WriteRecord(global, 0x01BC, new byte[] { 0, 0 });             // Prot4RevPassword
        WriteRecord(global, 0x003D, WindowPalette());                 // WindowPalette
        WriteRecord(global, 0x0040, new byte[] { 0, 0 });             // Backup
        WriteRecord(global, 0x008D, new byte[] { 0, 0 });             // HideObj
        WriteRecord(global, 0x0022, new byte[] { (byte)(date1904 ? 1 : 0), 0 }); // Date1904
        WriteRecord(global, 0x000E, new byte[] { 1, 0 });             // CalcPrecision
        WriteRecord(global, 0x01B7, new byte[] { 0, 0 });             // RefreshAll
        WriteRecord(global, 0x00DA, new byte[] { 0, 0 });             // BookBool

        WriteRecord(global, OpFont, Font());
        foreach (var (id, code) in BuiltInFormats)
            WriteRecord(global, OpFormat, FormatRecord(id, code));
        for (int k = 0; k < formats.Count; k++)
            WriteRecord(global, OpFormat, FormatRecord(164 + k, formats[k]));
        // BIFF8：前 16 个 XF 为内置样式 XF（Excel 必需），单元格 XF 从索引 16 起
        for (int i = 0; i < 16; i++)
            WriteRecord(global, OpXf, Xf(0, isStyle: true));
        WriteRecord(global, OpXf, Xf(0, isStyle: false)); // XF 16 = General 单元格样式
        for (int k = 0; k < formats.Count; k++)
            WriteRecord(global, OpXf, Xf(164 + k, isStyle: false));
        WriteRecord(global, 0x0160, new byte[] { 0, 0 });             // UsesELFs

        for (int i = 0; i < sheets.Count; i++)
        {
            boundFieldOffsets[i] = (int)global.Position + 4; // 记录头(4) 之后即 lbPlyPos
            WriteRecord(global, OpBoundSheet, BoundSheet(sheets[i].SheetName, 0));
        }

        WriteRecord(global, OpCountry, new byte[] { 1, 0, 1, 0 });
        WriteSst(global, sst);
        WriteRecord(global, OpEof, Array.Empty<byte>());

        var globalBytes = global.ToArray();

        // ── 各工作表子流 ──
        var sheetBytes = new byte[sheets.Count][];
        int sheetStart = globalBytes.Length;
        var positions = new int[sheets.Count];
        for (int i = 0; i < sheets.Count; i++)
        {
            positions[i] = sheetStart;
            sheetBytes[i] = BuildSheet(sheets[i], sst, sstIndex, GetXf, date1904);
            sheetStart += sheetBytes[i].Length;
        }

        // ── 原地打补丁：BOUNDSHEET 位置 ──
        for (int i = 0; i < sheets.Count; i++)
            WriteU32(globalBytes, boundFieldOffsets[i], (uint)positions[i]);

        using var outMs = new MemoryStream();
        outMs.Write(globalBytes, 0, globalBytes.Length);
        for (int i = 0; i < sheets.Count; i++)
            outMs.Write(sheetBytes[i], 0, sheetBytes[i].Length);
        return outMs.ToArray();
    }

    private static byte[] BuildSheet(SheetData sheet, List<string> sst, Dictionary<string, int> sstIndex,
        Func<string?, int> getXf, bool date1904)
    {
        using var ms = new MemoryStream();
        WriteRecord(ms, OpBof, Bof(0x0010)); // worksheet

        // 工作表设置记录（Excel 必需）
        WriteRecord(ms, 0x000D, new byte[] { 1, 0 });                      // CalcMode
        WriteRecord(ms, 0x000C, new byte[] { 0x64, 0 });                   // CalcCount = 100
        WriteRecord(ms, 0x000F, new byte[] { 1, 0 });                      // CalcRefMode
        WriteRecord(ms, 0x0011, new byte[] { 0, 0 });                      // CalcIter
        WriteRecord(ms, 0x0010, BitConverter.GetBytes(0.001));             // CalcDelta
        WriteRecord(ms, 0x005F, new byte[] { 1, 0 });                      // CalcSaveRecalc
        WriteRecord(ms, 0x002A, new byte[] { 0, 0 });                      // PrintRowCol
        WriteRecord(ms, 0x002B, new byte[] { 0, 0 });                      // PrintGrid
        WriteRecord(ms, 0x0082, new byte[] { 1, 0 });                      // GridSet
        WriteRecord(ms, 0x0080, new byte[8]);                              // Guts
        WriteRecord(ms, 0x0083, new byte[] { 0, 0 });                      // HCenter
        WriteRecord(ms, 0x0084, new byte[] { 0, 0 });                      // VCenter

        // 计算使用范围
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

        if (maxRow < 0)
        {
            WriteRecord(ms, OpDimensions, Dimensions(0, 1, 0, 1));
            WriteRecord(ms, OpWindow2, Window2(false));
            WriteRecord(ms, OpEof, Array.Empty<byte>());
            return ms.ToArray();
        }
        if (sheet.ColumnWidths is { } cw)
        {
            for (int c = 0; c < cw.Count; c++)
            {
                double w = cw[c];
                if (w <= 0) continue;
                WriteRecord(ms, OpColInfo, ColInfo(c, c, w));
            }
        }

        WriteRecord(ms, OpDimensions, Dimensions(0, maxRow + 1, 0, maxCol + 1));

        for (int r = 0; r <= maxRow; r++)
        {
            var row = sheet.Rows.Count > r ? sheet.Rows[r] : null;
            // 行内非空单元格（按列升序）
            var cells = new List<(int Col, Cell Cell)>();
            int rowMaxCol = -1;
            if (row is not null)
            {
                for (int c = 0; c < row.Count && c <= maxCol; c++)
                {
                    if (row[c].IsEmpty) continue;
                    cells.Add((c, row[c]));
                    if (c > rowMaxCol) rowMaxCol = c;
                }
            }

            double? height = sheet.RowHeights is not null && sheet.RowHeights.TryGetValue(r, out var h) ? h : null;
            if (cells.Count == 0 && height is null) continue;

            if (height is { } rh)
                WriteRecord(ms, OpRow, Row(r, 0, rowMaxCol + 1, rh));

            foreach (var (col, cell) in cells)
                WriteCell(ms, r, col, cell, sst, sstIndex, getXf, date1904);
        }

        // 尾部：WINDOW2 → PANE → MERGEDCELLS → CodeName → FeatHdr → Feat → EOF（对齐 Excel/SheetJS）
        int freezeRows = sheet.FreezeRows;
        int freezeCols = sheet.FreezeColumns;
        if (sheet.FreezeHeader) freezeRows = Math.Max(freezeRows, 1);
        bool hasFreeze = freezeRows > 0 || freezeCols > 0;
        WriteRecord(ms, OpWindow2, Window2(hasFreeze));
        if (hasFreeze)
            WriteRecord(ms, OpPane, Pane(freezeRows, freezeCols));
        if (sheet.MergedRanges.Count > 0)
            WriteRecord(ms, OpMergedCells, MergedCells(sheet.MergedRanges));
        WriteHyperlinks(ms, sheet);
        WriteRecord(ms, 0x01BA, CodeName(sheet.SheetName));
        WriteRecord(ms, 0x0867, new byte[]
        {
            0x67, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
        });
        WriteRecord(ms, 0x0868, new byte[]
        {
            0x68, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x01, 0x00, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x04, 0x00, 0x00, 0x00,
        });

        WriteRecord(ms, OpEof, Array.Empty<byte>());
        return ms.ToArray();
    }

    /// <summary>写出 BIFF8 HLINK（0x01B8）+ HLinkTooltip（0x0800）记录（对齐 Excel / SheetJS 布局）</summary>
    private static void WriteHyperlinks(MemoryStream ms, SheetData sheet)
    {
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            var row = sheet.Rows[r];
            for (int c = 0; c < row.Count; c++)
            {
                var cell = row[c];
                var link = cell.Hyperlink;
                if (link is null || string.IsNullOrEmpty(link.Target)) continue;

                string display = cell.Type == CellType.Text ? cell.Text ?? "" : "";
                var hlink = BuildHlink(r, c, link.Target, link.IsInternal, display);
                if (hlink.Length <= MaxRecordData)
                    WriteRecord(ms, OpHlink, hlink);

                if (!string.IsNullOrEmpty(link.Tooltip))
                {
                    var tt = BuildHlinkTooltip(r, c, link.Tooltip!);
                    if (tt.Length <= MaxRecordData)
                        WriteRecord(ms, OpHlinkTooltip, tt);
                }
            }
        }
    }

    /// <summary>
    /// BIFF8 HLINK 数据（对齐 Excel 字节）：Ref(8) + HyperlinkCLSID(16) + sVer(4)=2 + flags(4)
    /// + displayName(HyperlinkString) + [内部: loc  |  外部: URL Moniker CLSID(16) + len(4) + URL(UTF-16LE,null)]
    /// </summary>
    private static byte[] BuildHlink(int rw, int col, string target, bool isInternal, string displayText)
    {
        string url = isInternal ? "" : (target.StartsWith("file://", StringComparison.OrdinalIgnoreCase) ? target.Substring(7) : target);
        string loc = isInternal ? target.Substring(1) : "";

        int displayBytes = 4 + 2 * (displayText.Length + 1);
        int locBytes = isInternal ? 4 + 2 * (loc.Length + 1) : 0;
        int urlBytes = isInternal ? 0 : 16 + 4 + 2 * (url.Length + 1);

        int size = 8 + 16 + 4 + 4 + displayBytes + locBytes + urlBytes;
        var d = new byte[size];
        int off = 0;

        // Ref8U：rwFirst, rwLast, colFirst, colLast
        WriteU16(d, off, (ushort)rw); WriteU16(d, off + 2, (ushort)rw);
        WriteU16(d, off + 4, (ushort)col); WriteU16(d, off + 6, (ushort)col);
        off += 8;

        // Hyperlink CLSID（标准 OLE）
        byte[] clsid = { 0xD0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B };
        Array.Copy(clsid, 0, d, off, 16); off += 16;

        // sVer = 2
        WriteU32(d, off, 2u); off += 4;

        // flags：外部 0x0017（moniker + displayName），内部 0x001C（loc）
        uint flags = isInternal ? 0x001Cu : 0x0017u;
        WriteU32(d, off, flags); off += 4;

        WriteHlinkString(d, ref off, displayText);

        if (isInternal)
        {
            WriteHlinkString(d, ref off, loc);
        }
        else
        {
            // URL Moniker CLSID + 长度(字节数, 含 null) + UTF-16LE
            byte[] urlClsid = { 0xE0, 0xC9, 0xEA, 0x79, 0xF9, 0xBA, 0xCE, 0x11, 0x8C, 0x82, 0x00, 0xAA, 0x00, 0x4B, 0xA9, 0x0B };
            Array.Copy(urlClsid, 0, d, off, 16); off += 16;
            WriteU32(d, off, (uint)(2 * (url.Length + 1))); off += 4;
            for (int i = 0; i < url.Length; i++)
            {
                WriteU16(d, off, (ushort)url[i]); off += 2;
            }
            WriteU16(d, off, 0); off += 2;
        }

        return d;
    }

    private static void WriteHlinkString(byte[] d, ref int off, string s)
    {
        WriteU32(d, off, (uint)(s.Length + 1)); off += 4;
        for (int i = 0; i < s.Length; i++)
        {
            WriteU16(d, off, (ushort)s[i]); off += 2;
        }
        WriteU16(d, off, 0); off += 2;
    }

    /// <summary>HLinkTooltip（0x0800）：0x0800 + Ref(8) + UTF-16LE(含 null)</summary>
    private static byte[] BuildHlinkTooltip(int rw, int col, string tooltip)
    {
        var d = new byte[2 + 8 + 2 * (tooltip.Length + 1)];
        int off = 0;
        WriteU16(d, off, 0x0800); off += 2;
        WriteU16(d, off, (ushort)rw); WriteU16(d, off + 2, (ushort)rw);
        WriteU16(d, off + 4, (ushort)col); WriteU16(d, off + 6, (ushort)col);
        off += 8;
        for (int i = 0; i < tooltip.Length; i++)
        {
            WriteU16(d, off, (ushort)tooltip[i]); off += 2;
        }
        WriteU16(d, off, 0); off += 2;
        return d;
    }

    private static void WriteU32(byte[] d, int offset, uint v)
    {
        d[offset] = (byte)v;
        d[offset + 1] = (byte)(v >> 8);
        d[offset + 2] = (byte)(v >> 16);
        d[offset + 3] = (byte)(v >> 24);
    }

    /// <summary>CodeName（0x01BA）：Unicode 字符串形式的 VBA 工作表代号</summary>
    private static byte[] CodeName(string name)
    {
        var nameData = Encoding.Unicode.GetBytes(name);
        var d = new byte[3 + nameData.Length];
        WriteU16(d, 0, (ushort)name.Length);
        d[2] = 0x01; // grbit：高字节
        Array.Copy(nameData, 0, d, 3, nameData.Length);
        return d;
    }

    private static void WriteCell(MemoryStream ms, int rw, int col, Cell cell,
        List<string> sst, Dictionary<string, int> sstIndex, Func<string?, int> getXf, bool date1904)
    {
        int ixfe = getXf(cell.Type == CellType.Date || cell.Type == CellType.Number ? cell.NumberFormat : null);
        switch (cell.Type)
        {
            case CellType.Text:
            {
                var text = cell.Text ?? "";
                if (!sstIndex.TryGetValue(text, out var isst))
                {
                    isst = sst.Count;
                    sst.Add(text);
                    sstIndex[text] = isst;
                }
                var d = new byte[10]; // rw(2) col(2) ixfe(2) isst(4)
                WriteU16(d, 0, (ushort)rw);
                WriteU16(d, 2, (ushort)col);
                WriteU16(d, 4, (ushort)ixfe);
                WriteU32(d, 6, (uint)isst);
                WriteRecord(ms, OpLabelSst, d);
                break;
            }
            case CellType.Number:
            {
                double v = cell.Number;
                var d = new byte[14]; // rw(2) col(2) ixfe(2) num(8)
                WriteU16(d, 0, (ushort)rw);
                WriteU16(d, 2, (ushort)col);
                WriteU16(d, 4, (ushort)ixfe);
                Array.Copy(BitConverter.GetBytes(v), 0, d, 6, 8);
                WriteRecord(ms, OpNumber, d);
                break;
            }
            case CellType.Date:
            {
                var d = new byte[14]; // rw(2) col(2) ixfe(2) num(8)
                WriteU16(d, 0, (ushort)rw);
                WriteU16(d, 2, (ushort)col);
                WriteU16(d, 4, (ushort)ixfe);
                Array.Copy(BitConverter.GetBytes(FormatDetector.DateToSerial(cell.Date, date1904)), 0, d, 6, 8);
                WriteRecord(ms, OpNumber, d);
                break;
            }
            case CellType.Boolean:
            {
                var d = new byte[8];
                WriteU16(d, 0, (ushort)rw);
                WriteU16(d, 2, (ushort)col);
                WriteU16(d, 4, (ushort)ixfe);
                d[6] = cell.Boolean ? (byte)1 : (byte)0;
                d[7] = 0; // 非错误
                WriteRecord(ms, OpBoolErr, d);
                break;
            }
        }
    }

    // ── 记录体构造 ──

    private static byte[] Bof(int type)
    {
        // 对齐 Excel/SheetJS：version(2) type(2) build(2) buildYear(2) fileHistoryFlags(4) lowest(2) lowestBuild(2) lowestBuildYear(2)
        var d = new byte[16];
        WriteU16(d, 0, 0x0600);
        WriteU16(d, 2, (ushort)type);
        WriteU16(d, 4, 0x7262);
        WriteU16(d, 6, 0x07CD);
        WriteU16(d, 8, 0xC009);
        WriteU16(d, 10, 0x0001);
        WriteU16(d, 12, 0x0706); // 最低兼容版本（Excel 校验此字段非 0）
        return d;
    }

    private static byte[] WriteAccess()
    {
        var d = new byte[112];
        const string name = "LiteExcel";
        d[0] = (byte)name.Length; // cch
        d[1] = (byte)(name.Length >> 8);
        d[2] = 0;                 // grbit: 高字节关闭（压缩 Latin1）
        var bytes = Latin1.GetBytes(name);
        Array.Copy(bytes, 0, d, 3, bytes.Length);
        return d;
    }

    private static byte[] BoundSheet(string name, int position)
    {
        // BIFF8 下表名恒以 Unicode 写出（对齐 Excel/SheetJS）
        var nameData = Encoding.Unicode.GetBytes(name);
        var d = new byte[8 + nameData.Length];
        WriteU32(d, 0, (uint)position);          // lbPlyPos
        WriteU16(d, 4, 0);                       // grbit（可见）
        d[6] = (byte)name.Length;                // cch
        d[7] = 0x01;                             // grbit（高字节 = Unicode）
        Array.Copy(nameData, 0, d, 8, nameData.Length);
        return d;
    }

    /// <summary>写入 SST 记录（含 CONTINUE 续接）。字符串压缩仅限纯 ASCII，其余用 UTF-16。
    /// 无字符串时也写一条空 SST（count=0），与 Excel/SheetJS 一致。</summary>
    private static void WriteSst(MemoryStream ms, List<string> strings)
    {
        if (strings.Count == 0)
        {
            var empty = new byte[8];
            WriteRecord(ms, OpSst, empty);
            return;
        }

        var segments = new List<byte[]>();
        var cur = new MemoryStream();
        int curCap = MaxRecordData - 8; // 首条记录减 SST 头

        void Append(byte b) => cur.WriteByte(b);

        void StartSegment()
        {
            segments.Add(cur.ToArray());
            cur = new MemoryStream();
            curCap = MaxRecordData;
        }

        // SST 头
        Append((byte)strings.Count);
        Append((byte)(strings.Count >> 8));
        Append((byte)(strings.Count >> 16));
        Append((byte)(strings.Count >> 24));
        Append((byte)strings.Count);
        Append((byte)(strings.Count >> 8));
        Append((byte)(strings.Count >> 16));
        Append((byte)(strings.Count >> 24));

        foreach (var s in strings)
        {
            bool highByte = s.Any(c => c >= 0x80); // 纯 ASCII 压缩，其余 UTF-16（规避 cp1252 差异）
            byte[] chars = highByte ? Encoding.Unicode.GetBytes(s) : Latin1.GetBytes(s);
            int cch = s.Length;

            // 3 字节头必须完整落在同一段
            if (cur.Length + 3 > curCap) StartSegment();
            Append((byte)(cch & 0xFF));
            Append((byte)((cch >> 8) & 0xFF));
            Append(highByte ? (byte)0x01 : (byte)0x00);

            int pos = 0;
            while (pos < chars.Length)
            {
                int space = curCap - (int)cur.Length;
                if (space <= 0)
                {
                    // 跨段：段首写续接 grbit
                    StartSegment();
                    Append(highByte ? (byte)0x01 : (byte)0x00);
                    continue;
                }
                int take = Math.Min(space, chars.Length - pos);
                if (highByte) take -= take & 1; // UTF-16 不得切裂字符
                if (take <= 0)
                {
                    StartSegment();
                    Append(highByte ? (byte)0x01 : (byte)0x00);
                    continue;
                }
                cur.Write(chars, pos, take);
                pos += take;
            }
        }
        if (cur.Length > 0) segments.Add(cur.ToArray());

        for (int i = 0; i < segments.Count; i++)
            WriteRecord(ms, i == 0 ? OpSst : OpContinue, segments[i]);
    }

    private static byte[] FormatRecord(int ifmt, string code)
    {
        bool highByte = code.Any(c => c >= 0x80);
        var codeData = highByte ? Encoding.Unicode.GetBytes(code) : Latin1.GetBytes(code);
        var d = new byte[5 + codeData.Length];
        WriteU16(d, 0, (ushort)ifmt);
        WriteU16(d, 2, (ushort)code.Length);
        d[4] = highByte ? (byte)0x01 : (byte)0x00;
        Array.Copy(codeData, 0, d, 5, codeData.Length);
        return d;
    }

    /// <summary>BIFF8 XF（20 字节，与 Excel 写出一致）；ifmt 在 offset 2。
    /// 内置样式 XF（索引 0-15）带样式标志 0xFFF4，单元格 XF 从索引 16 起。</summary>
    private static byte[] Xf(int ifmt, bool isStyle)
    {
        var d = new byte[20];
        WriteU16(d, 0, 0);      // ifnt
        WriteU16(d, 2, (ushort)ifmt);
        if (isStyle) WriteU16(d, 4, 0xFFF4);
        return d;
    }

    /// <summary>默认字体（索引 0）：Arial 12pt 常规，与 SheetJS 写出字节一致</summary>
    private static byte[] Font()
    {
        var d = new byte[26];
        WriteU16(d, 0, 240);       // dyHeight 12pt
        WriteU16(d, 6, 0x0190);    // bls = 常规
        d[14] = 5;                 // bFamily
        d[15] = 1;                 // bCharSet
        var name = Encoding.Unicode.GetBytes("Arial");
        Array.Copy(name, 0, d, 16, name.Length);
        return d;
    }

    /// <summary>WindowPalette（0x003D）18 字节，对齐 SheetJS 写出</summary>
    private static byte[] WindowPalette()
    {
        var d = new byte[18];
        WriteU16(d, 4, 0x7260);
        WriteU16(d, 6, 0x44C0);
        WriteU16(d, 8, 0x0038);
        WriteU16(d, 14, 1);
        WriteU16(d, 16, 0x01F4);
        return d;
    }

    private static byte[] Dimensions(int rwMic, int rwMac, int colMic, int colMac)
    {
        var d = new byte[14]; // BIFF8 DIMENSIONS = rwMic(4) rwMac(4) colMic(2) colMac(2) reserved(2)
        WriteU32(d, 0, (uint)rwMic);
        WriteU32(d, 4, (uint)rwMac);
        WriteU16(d, 8, (ushort)colMic);
        WriteU16(d, 10, (ushort)colMac);
        return d;
    }

    private static byte[] Window2(bool frozen)
    {
        var d = new byte[18];
        ushort flags = (ushort)(0x06B6 | (frozen ? 0x08 : 0));
        WriteU16(d, 0, flags);
        WriteU32(d, 8, 0x00000040); // 默认网格线颜色（对齐 Excel）
        return d;
    }

    private static byte[] Pane(int freezeRows, int freezeCols)
    {
        var d = new byte[10];
        WriteU16(d, 0, (ushort)freezeCols); // xSplit
        WriteU16(d, 2, (ushort)freezeRows); // ySplit
        WriteU16(d, 4, (ushort)freezeRows); // topRow
        WriteU16(d, 6, (ushort)freezeCols); // leftCol
        d[8] = freezeRows > 0 && freezeCols > 0 ? (byte)0 : (byte)(freezeRows > 0 ? 2 : 1); // activePane: 双=topLeft, 行=bottomLeft, 列=topRight
        d[9] = 0; // fNoSplit = false（有分隔）
        return d;
    }

    private static byte[] MergedCells(List<CellRange> ranges)
    {
        var d = new byte[2 + ranges.Count * 8];
        WriteU16(d, 0, (ushort)ranges.Count);
        for (int i = 0; i < ranges.Count; i++)
        {
            int o = 2 + i * 8;
            WriteU16(d, o, (ushort)ranges[i].FirstRow);
            WriteU16(d, o + 2, (ushort)ranges[i].LastRow);
            WriteU16(d, o + 4, (ushort)ranges[i].FirstCol);
            WriteU16(d, o + 6, (ushort)ranges[i].LastCol);
        }
        return d;
    }

    private static byte[] ColInfo(int colFirst, int colLast, double width)
    {
        var d = new byte[12]; // colFirst(2) colLast(2) colw(2) ixfe(2) grbit(2) reserved(2)
        WriteU16(d, 0, (ushort)colFirst);
        WriteU16(d, 2, (ushort)colLast);
        WriteU16(d, 4, (ushort)(width * 256));
        WriteU16(d, 6, 0);       // ixfe
        WriteU16(d, 8, 0x0002);  // grbit: 显式宽度
        return d;
    }

    private static byte[] Row(int rw, int colMic, int colMac, double? height)
    {
        var d = new byte[16];
        WriteU16(d, 0, (ushort)rw);
        WriteU16(d, 2, (ushort)colMic);
        WriteU16(d, 4, (ushort)colMac);
        WriteU16(d, 6, height is { } h ? (ushort)(h * 20) : (ushort)0x011D); // 默认行高 285（14.25pt，对齐 Excel）
        WriteU16(d, 8, 0);       // ixfe
        WriteU16(d, 10, height is null ? (ushort)0 : (ushort)0x01); // grbit: fExSet
        return d;
    }

    // ── 基础写入 ──

    private static void WriteRecord(MemoryStream ms, ushort opcode, byte[] data)
    {
        ms.WriteByte((byte)(opcode & 0xFF));
        ms.WriteByte((byte)((opcode >> 8) & 0xFF));
        ms.WriteByte((byte)(data.Length & 0xFF));
        ms.WriteByte((byte)((data.Length >> 8) & 0xFF));
        ms.Write(data, 0, data.Length);
    }

    private static void WriteU16(byte[] d, int offset, ushort v)
    {
        d[offset] = (byte)v;
        d[offset + 1] = (byte)(v >> 8);
    }
}
