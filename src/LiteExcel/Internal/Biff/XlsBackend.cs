using LiteExcel.Internal.Cfb;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// 传统 .xls（OLE2 + BIFF8）读取后端。
/// 仅读取：数据单元格、日期识别、合并单元格、列宽、行高、冻结表头。
/// </summary>
internal static class XlsBackend
{
    /// <summary>net48 无 Encoding.Latin1，用 ISO-8859-1 代码页等价 </summary>
    private static readonly Encoding Latin1 = Encoding.GetEncoding(28591);

    /// <summary>最近一次 ReadAll 捕获的工作簿命名区域（DefinedName）快照，供 Excel.cs 挂到 Workbook.Names。 </summary>
    [ThreadStatic]
    private static List<NamedRange>? s_definedNames;

    /// <summary>取最近一次 .xls 读取解析出的命名区域列表（可为 null/空）。 </summary>
    public static IReadOnlyList<NamedRange>? DefinedNamesSnapshot => s_definedNames;

    public static List<SheetData> ReadAll(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadAll(fs);
    }

    /// <summary>从流读取 .xls 工作簿（含命名区域快照捕获）。 </summary>
    public static List<SheetData> ReadAll(Stream stream)
    {
        var cfb = CfbFile.Open(stream);
        var workbook = cfb.GetStream("Workbook") ?? cfb.GetStream("Book");
        if (workbook is null)
            throw new LiteExcelException(".xls 文件中缺少 Workbook/Book 流");
        return ParseWorkbook(workbook);
    }

    /// <summary>读取 .xls 工作簿的 1904 日期系统标志（DATE1904 记录，0x0022） </summary>
    public static bool ReadDate1904(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadDate1904(fs);
    }

    /// <summary>从流读取 .xls 工作簿的 1904 日期系统标志 </summary>
    public static bool ReadDate1904(Stream stream)
    {
        var cfb = CfbFile.Open(stream);
        var workbook = cfb.GetStream("Workbook") ?? cfb.GetStream("Book");
        if (workbook is null) return false;
        var records = BiffRecords.ReadAll(workbook);
        foreach (var rec in records)
        {
            if (rec.Opcode == BiffRecords.OpDateMode && rec.Data.Length >= 2)
                return BiffRecords.ReadU16(rec.Data, 0) == 1;
        }
        return false;
    }

    private static List<SheetData> ParseWorkbook(byte[] wb)
    {
        var records = BiffRecords.ReadAll(wb);
        if (records.Count == 0 || records[0].Opcode != BiffRecords.OpBof)
            throw new LiteExcelException("这不是有效的 .xls 文件（缺少全局 BOF 记录）");

        var sst = new List<string>();
        var formats = new Dictionary<int, string>();
        var xfIfmt = new List<int>();
        bool date1904 = false;
        var boundSheets = new List<(string Name, byte Type)>();
        var externSheets = new List<int>();          // EXTERNSHEET itab 列表（ixti → itab）
        var definedNameRecords = new List<(byte[] Data, ushort Opcode)>();

        int i = 1;
        for (; i < records.Count; i++)
        {
            var rec = records[i];
            if (rec.Opcode == BiffRecords.OpEof) { i++; break; }

            switch (rec.Opcode)
            {
                case BiffRecords.OpFilePass:
                    // BIFF8 加密：BOF 后紧跟明文 FILEPASS 记录（标准/增强加密标记），后续记录全部密文
                    throw new LiteExcelException("该 .xls 文件已加密（带打开密码）。当前版本暂不支持读取加密工作簿，" +
                        "请在 Excel 中另存为无密码文件后再打开（密码支持规划在后续版本）。");
                case BiffRecords.OpBoundSheet:
                    boundSheets.Add(ParseBoundSheet(rec.Data));
                    break;
                case BiffRecords.OpSst:
                    ParseSst(records, ref i, sst);
                    break;
                case BiffRecords.OpFormat:
                    ParseFormat(rec.Data, formats);
                    break;
                case BiffRecords.OpXf:
                    if (rec.Data.Length >= 4)
                        xfIfmt.Add(BiffRecords.ReadU16(rec.Data, 2));
                    break;
                case BiffRecords.OpDateMode:
                    if (rec.Data.Length >= 2)
                        date1904 = BiffRecords.ReadU16(rec.Data, 0) == 1;
                    break;
                case BiffRecords.OpExternSheet:
                    ParseExternSheet(rec.Data, externSheets);
                    break;
                case BiffRecords.OpDefinedName:
                    ParseDefinedName(rec.Data, definedNameRecords);
                    break;
            }
        }

        // 命名区域：先按 EXTERNSHEET itab 映射到 sheet 名，再解析公式
        var names = new List<NamedRange>();
        foreach (var dn in definedNameRecords)
        {
            var nr = BuildNamedRange(dn, externSheets, boundSheets);
            if (nr is not null) names.Add(nr);
        }
        s_definedNames = names;
        var result = new List<SheetData>();
        while (i < records.Count)
        {
            if (records[i].Opcode != BiffRecords.OpBof) { i++; continue; }
            i++;
            var sheet = ParseSheet(records, ref i, sst, formats, xfIfmt, date1904);
            if (result.Count < boundSheets.Count)
                sheet.SheetName = boundSheets[result.Count].Name;
            result.Add(sheet);
        }

        if (result.Count == 0)
            throw new LiteExcelException("这不是有效的 .xls 文件（未找到任何工作表）");
        return result;
    }

    // ── 全局记录 ──

    private static (string Name, byte Type) ParseBoundSheet(byte[] d)
    {
        // lbPlyPos(4) grbit(2) cch(1) grbit(1) name
        if (d.Length < 8) return ("", 0);
        int cch = d[6];
        bool highByte = (d[7] & 0x01) != 0;
        string name;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, d.Length - 8);
            name = Encoding.Unicode.GetString(d, 8, bytes);
        }
        else
        {
            int bytes = Math.Min(cch, d.Length - 8);
            name = Latin1.GetString(d, 8, bytes);
        }
        return (name, 0);
    }

    // ── 命名区域（DEFINEDNAME / EXTERNSHEET）──

    /// <summary>EXTERNSHEET 记录：cXTI(2) + 每项 iSupBook(2) itabFirst(2) itabLast(2)。收集 itabFirst 序列（ixti → itab）。</summary>
    private static void ParseExternSheet(byte[] d, List<int> externSheets)
    {
        if (d.Length < 2) return;
        int count = BiffRecords.ReadU16(d, 0);
        int pos = 2;
        for (int i = 0; i < count && pos + 6 <= d.Length; i++)
        {
            int itab = BiffRecords.ReadU16(d, pos + 2);
            externSheets.Add(itab);
            pos += 6;
        }
    }

    /// <summary>DEFINEDNAME 记录原始数据暂存（后续与 EXTERNSHEET 一起解析）。</summary>
    private static void ParseDefinedName(byte[] d, List<(byte[] Data, ushort Opcode)> sink)
    {
        if (d.Length >= 14) sink.Add((d, 0));
    }

    /// <summary>把 DEFINEDNAME + EXTERNSHEET 解析为 NamedRange；无法解析（复杂公式/内置名）返回 null 跳过。</summary>
    private static NamedRange? BuildNamedRange((byte[] Data, ushort Opcode) dn,
        List<int> externSheets, List<(string Name, byte Type)> boundSheets)
    {
        var d = dn.Data;
        // option_flag(2) keyboard_shortcut(1) cch(1) cce(2) extern(2) sheet(2)
        //   customMenuLen(1) descLen(1) helpLen(1) statusLen(1) nameIsMultibyte(1) [built_in?] name formula
        if (d.Length < 15) return null;
        int optionFlag = BiffRecords.ReadU16(d, 0);
        int cch = d[3];
        int cce = BiffRecords.ReadU16(d, 4);
        int sheetNumber = BiffRecords.ReadU16(d, 8);
        bool isBuiltIn = (optionFlag & 0x0020) != 0;
        if (isBuiltIn) return null; // 内置名（Print_Area 等）不作为用户命名区域
        bool multibyte = d[14] != 0;
        int namePos = 15;
        string name = multibyte
            ? (namePos + cch * 2 <= d.Length ? Encoding.Unicode.GetString(d, namePos, cch * 2) : "")
            : (namePos + cch <= d.Length ? Latin1.GetString(d, namePos, cch) : "");
        namePos += multibyte ? cch * 2 : cch;
        if (namePos + cce > d.Length) return null;

        // 公式：仅支持 PtgRef3d(0x3A) / PtgArea3d(0x3B) 的「ixti + 行列」简单引用；其它公式类型跳过
        var formula = new byte[cce];
        Array.Copy(d, namePos, formula, 0, cce);
        string? reference = DecodeRef3d(formula, externSheets, boundSheets);
        if (reference is null) return null;

        // 局部名：sheetNumber 字段 = 总表数 - sheetIndex（本 Excel 版本行为）；全局名 = 0 → -1
        int localSheetId;
        if (sheetNumber == 0) localSheetId = -1;
        else localSheetId = boundSheets.Count - sheetNumber;

        return new NamedRange
        {
            Name = name,
            Reference = reference,
            LocalSheetId = localSheetId,
        };
    }

    /// <summary>解码 PtgRef3d/PtgArea3d 为 "Sheet!$A$1" 文本；不支持返回 null。</summary>
    private static string? DecodeRef3d(byte[] formula, List<int> externSheets, List<(string Name, byte Type)> boundSheets)
    {
        if (formula.Length == 0) return null;
        int ptg = formula[0];
        int pos = 1;
        string? sheet = null;

        // 3D 引用：ixti(2) 指向 EXTERNSHEET → itab → 本地 sheet 名
        if (ptg == 0x3A || ptg == 0x3B)
        {
            if (pos + 2 > formula.Length) return null;
            int ixti = BiffRecords.ReadU16(formula, pos);
            pos += 2;
            if (ixti >= 0 && ixti < externSheets.Count)
            {
                int itab = externSheets[ixti];
                // itab → sheet 名：BOUNDSHEET 顺序（0-based），版本差异用 boundCount-1-itab 兜底
                if (itab >= 0 && itab < boundSheets.Count)
                    sheet = boundSheets[itab].Name;
                else if (itab >= 0 && itab < boundSheets.Count * 1 && boundSheets.Count - 1 - itab >= 0)
                {
                    int idx = boundSheets.Count - 1 - itab;
                    if (idx < boundSheets.Count) sheet = boundSheets[idx].Name;
                }
            }
            if (sheet is null) return null;
        }
        else
        {
            // 非 3D（纯局部引用）暂不支持——命名区域几乎总是带表名
            return null;
        }

        if (ptg == 0x3A) // PtgRef3d: rw(2) col(2)
        {
            if (pos + 4 > formula.Length) return null;
            int rw = BiffRecords.ReadU16(formula, pos);
            int col = BiffRecords.ReadU16(formula, pos + 2);
            return sheet + "!" + RefCellText(rw, col);
        }
        if (ptg == 0x3B) // PtgArea3d: rw1(2) rw2(2) col1(2) col2(2)
        {
            if (pos + 8 > formula.Length) return null;
            int rw1 = BiffRecords.ReadU16(formula, pos);
            int rw2 = BiffRecords.ReadU16(formula, pos + 2);
            int col1 = BiffRecords.ReadU16(formula, pos + 4);
            int col2 = BiffRecords.ReadU16(formula, pos + 6);
            return sheet + "!" + RefCellText(rw1, col1) + ":" + RefCellText(rw2, col2);
        }
        return null;
    }

    private static string RefCellText(int rw, int col)
    {
        var sb = new StringBuilder();
        int c = col;
        while (c >= 0)
        {
            sb.Insert(0, (char)('A' + (c % 26)));
            c = c / 26 - 1;
            if (c < 0) break;
        }
        sb.Append(rw + 1);
        return sb.ToString();
    }

    private static void ParseFormat(byte[] d, Dictionary<int, string> formats)
    {
        if (d.Length < 5) return;
        int ifmt = BiffRecords.ReadU16(d, 0);
        int cch = BiffRecords.ReadU16(d, 2);
        bool highByte = (d[4] & 0x01) != 0;
        string code;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, d.Length - 5);
            code = Encoding.Unicode.GetString(d, 5, bytes);
        }
        else
        {
            int bytes = Math.Min(cch, d.Length - 5);
            code = Latin1.GetString(d, 5, bytes);
        }
        formats[ifmt] = code;
    }

    private static void ParseSst(List<BiffRecords.Record> records, ref int i, List<string> sst)
    {
        var sstData = records[i].Data;
        if (sstData.Length < 8) return;

        int uniqueCount = BiffRecords.ReadS32(sstData, 4);

        // 首段从字符串区开始（跳过 cstTotal(4) + cstUnique(4)）
        var firstSegment = new byte[sstData.Length - 8];
        Array.Copy(sstData, 8, firstSegment, 0, firstSegment.Length);
        var segments = new List<byte[]> { firstSegment };
        while (i + 1 < records.Count && records[i + 1].Opcode == BiffRecords.OpContinue)
        {
            i++;
            segments.Add(records[i].Data);
        }

        var reader = new BiffStringReader(segments);
        for (int k = 0; k < uniqueCount; k++)
        {
            var s = reader.ReadString();
            if (s is null) break; // 流耗尽或格式错误
            sst.Add(s);
        }
    }

    // ── 工作表子流 ──

    private static SheetData ParseSheet(List<BiffRecords.Record> records, ref int i,
        List<string> sst, Dictionary<int, string> formats, List<int> xfIfmt, bool date1904)
    {
        var sheet = new SheetData();
        var cells = new Dictionary<int, Dictionary<int, Cell>>();
        int maxRow = -1;
        int maxCol = -1;
        var colWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();
        bool freeze = false;
        int freezeRows = 0;
        int freezeCols = 0;

        for (; i < records.Count; i++)
        {
            var rec = records[i];
            if (rec.Opcode == BiffRecords.OpEof) { i++; break; }

            switch (rec.Opcode)
            {
                case BiffRecords.OpNumber:
                    PutCell(cells, rec.Data, 0, 2, 4, (row, col, d) =>
                        CellFromNumber(BitConverter.ToDouble(d, 6), ReadU16At(d, 4), xfIfmt, formats, date1904), ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpRk:
                    // RK = rw(2) col(2) ixfe(2) rk(4)
                    PutCell(cells, rec.Data, 0, 2, 6, (row, col, d) => Cell.FromNumber(BiffShared.DecodeRk(ReadS32At(d, 6))), ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpMulRk:
                    ParseMulRk(cells, rec.Data, xfIfmt, formats, date1904, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpLabelSst:
                    PutCell(cells, rec.Data, 0, 2, 4, (row, col, d) =>
                    {
                        int isst = ReadS32At(d, 6);
                        return isst >= 0 && isst < sst.Count ? Cell.FromText(sst[isst]) : Cell.Empty;
                    }, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpLabel:
                    PutCell(cells, rec.Data, 0, 2, 4, (row, col, d) => Cell.FromText(ParseLabelString(d)), ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpBoolErr:
                    PutCell(cells, rec.Data, 0, 2, 4, (row, col, d) =>
                    {
                        bool isError = d.Length > 7 && d[7] != 0;
                        if (isError) return Cell.FromText(BiffShared.ErrorCode(d.Length > 6 ? d[6] : (byte)0));
                        bool v = d.Length > 6 && d[6] != 0;
                        return Cell.FromBoolean(v);
                    }, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpFormula:
                    ParseFormula(cells, rec.Data, xfIfmt, formats, date1904, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpMergedCells:
                    ParseMergedCells(sheet, rec.Data);
                    break;
                case BiffRecords.OpHlink:
                    ParseHlink(cells, rec.Data, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpHlinkTooltip:
                    ParseHlinkTooltip(cells, rec.Data, ref maxRow, ref maxCol);
                    break;
                case BiffRecords.OpColInfo:
                    ParseColInfo(rec.Data, colWidths);
                    break;
                case BiffRecords.OpRow:
                    ParseRowHeight(rec.Data, rowHeights);
                    break;
                case BiffRecords.OpPane:
                    // xSplit(2) ySplit(2) topRow(2) leftCol(2) activePane(1) fNoSplit(1)
                    if (rec.Data.Length >= 8)
                    {
                        int xSplit = BiffRecords.ReadU16(rec.Data, 0);
                        int ySplit = BiffRecords.ReadU16(rec.Data, 2);
                        freeze = ySplit >= 1 || xSplit >= 1;
                        if (freeze) { freezeRows = ySplit; freezeCols = xSplit; }
                    }
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

    /// <summary>解析 HLINK（0x01B8）：Ref(8) + CLSID(16) + Hyperlink 对象</summary>
    private static void ParseHlink(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d, ref int maxRow, ref int maxCol)
    {
        if (d.Length < 28) return;
        int row = BiffRecords.ReadU16(d, 0);
        int col = BiffRecords.ReadU16(d, 4);
        int off = 24; // Ref(8) + Hyperlink CLSID(16)

        int sVer = BiffRecords.ReadS32(d, off); off += 4;
        if (sVer != 2) return;
        int flags = BiffRecords.ReadU16(d, off); off += 2;
        off += 2; // 保留字段（通常 0）

        string displayName = "";
        string loc = "";
        string target = "";

        if ((flags & 0x0010) != 0) displayName = ReadHlinkString(d, ref off);
        if ((flags & 0x0080) != 0) ReadHlinkString(d, ref off); // targetFrameName（忽略）
        if ((flags & 0x0100) != 0 && (flags & 0x0001) != 0) target = ReadHlinkString(d, ref off); // 字符串 moniker
        else if ((flags & 0x0001) != 0) target = ReadHlinkMoniker(d, ref off); // URL/File moniker
        if ((flags & 0x0008) != 0) loc = ReadHlinkString(d, ref off);
        if ((flags & 0x0020) != 0) off += 16; // GUID
        if ((flags & 0x0040) != 0) off += 8;  // FILETIME

        if (!string.IsNullOrEmpty(loc))
        {
            target = string.IsNullOrEmpty(target) ? "#" + loc : target + "#" + loc;
        }
        // file:// 前缀重建：flags 0x0002 且路径以单个 '/' 开头（对齐 SheetJS）
        if ((flags & 0x0002) != 0 && target.StartsWith("/", StringComparison.Ordinal)
            && !target.StartsWith("//", StringComparison.Ordinal))
        {
            target = "file://" + target;
        }
        if (string.IsNullOrEmpty(target)) return;

        // displayName 是单元格显示文本而非 tooltip（tooltip 由 HLinkTooltip 记录提供）
        var existing = GetCell(cells, row, col);
        AttachCellHyperlink(cells, row, col, new Hyperlink
        {
            Target = target,
            IsInternal = target.StartsWith("#", StringComparison.Ordinal),
            Tooltip = existing?.Hyperlink?.Tooltip, // 保留先出现的 HLinkTooltip
        }, ref maxRow, ref maxCol);
    }

    /// <summary>解析 HLinkTooltip（0x0800）：0x0800(2) + Ref(8) + UTF-16LE(含 null)</summary>
    private static void ParseHlinkTooltip(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d, ref int maxRow, ref int maxCol)
    {
        if (d.Length < 12) return;
        int row = BiffRecords.ReadU16(d, 2);
        int col = BiffRecords.ReadU16(d, 6);
        int off = 10;
        var sb = new System.Text.StringBuilder();
        while (off + 1 < d.Length)
        {
            int ch = BiffRecords.ReadU16(d, off);
            if (ch == 0) break;
            sb.Append((char)ch);
            off += 2;
        }
        string tooltip = sb.ToString();
        if (tooltip.Length == 0) return;

        var cell = GetOrCreateCell(cells, row, col, ref maxRow, ref maxCol);
        if (cell.Hyperlink is null)
        {
            cell.Hyperlink = new Hyperlink { Tooltip = tooltip };
        }
        else if (string.IsNullOrEmpty(cell.Hyperlink.Tooltip))
        {
            cell.Hyperlink.Tooltip = tooltip;
        }
    }

    /// <summary>HyperlinkString：len(4) + UTF-16LE 字符（len 含 null 结尾）</summary>
    private static string ReadHlinkString(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return "";
        int len = BiffRecords.ReadS32(d, off); off += 4;
        if (len <= 0 || len > 0xFFFF) return "";
        if (off + len * 2 > d.Length) return "";
        var chars = new char[len];
        for (int i = 0; i < len; i++)
            chars[i] = (char)BiffRecords.ReadU16(d, off + i * 2);
        off += len * 2;
        return new string(chars).TrimEnd('\0');
    }

    /// <summary>Moniker：CLSID(16) + 内容。支持 URL（E0C9EA79...）与 File（03030000...）</summary>
    private static string ReadHlinkMoniker(byte[] d, ref int off)
    {
        if (off + 16 > d.Length) return "";
        var clsid = new byte[16];
        Array.Copy(d, off, clsid, 0, 16);
        off += 16;

        // URL Moniker：len(4) 字节数 + UTF-16LE（含 null）+ 可选 GUID(16)+FILETIME(8)
        if (clsid[0] == 0xE0 && clsid[1] == 0xC9 && clsid[2] == 0xEA && clsid[3] == 0x79)
        {
            if (off + 4 > d.Length) return "";
            int byteLen = BiffRecords.ReadS32(d, off); off += 4;
            if (byteLen <= 0 || byteLen > 0xFFFF) return "";
            if (off + byteLen > d.Length) return "";
            int charCount = byteLen / 2;
            var chars = new char[charCount];
            for (int i = 0; i < charCount; i++)
                chars[i] = (char)BiffRecords.ReadU16(d, off + i * 2);
            off += byteLen;
            var url = new string(chars).TrimEnd('\0');
            // 尾部可选 GUID + FILETIME（SheetJS 判定为 24 字节）
            if (off + 24 <= d.Length)
            {
                byte[] guid = { 0x79, 0x58, 0x81, 0xF4, 0x3B, 0x1D, 0x7F, 0x48, 0xAF, 0x2C, 0x82, 0x5D, 0xC4, 0x85, 0x27, 0x63 };
                bool match = true;
                for (int i = 0; i < 16; i++) if (d[off + i] != guid[i]) { match = false; break; }
                if (match) off += 24;
            }
            return url;
        }

        // File Moniker：cAnti(2) + ANSI 路径（含 null）+ 后续字段（忽略）
        if (clsid[0] == 0x03 && clsid[1] == 0x03)
        {
            if (off + 2 > d.Length) return "";
            int cAnti = BiffRecords.ReadU16(d, off); off += 2;
            string prefix = "";
            for (int i = 0; i < cAnti; i++) prefix += "../";
            int start = off;
            int ansiLen = 0;
            while (off + ansiLen < d.Length && d[start + ansiLen] != 0) ansiLen++;
            var ansi = System.Text.Encoding.GetEncoding(0).GetString(d, start, ansiLen);
            off += ansiLen + 1;
            return prefix + ansi;
        }

        return "";
    }

    private static Cell? GetCell(Dictionary<int, Dictionary<int, Cell>> cells, int row, int col)
    {
        return cells.TryGetValue(row, out var rowCells) && rowCells.TryGetValue(col, out var cell) ? cell : null;
    }

    private static void AttachCellHyperlink(Dictionary<int, Dictionary<int, Cell>> cells, int row, int col,
        Hyperlink link, ref int maxRow, ref int maxCol)
    {
        var cell = GetOrCreateCell(cells, row, col, ref maxRow, ref maxCol);
        cell.Hyperlink = link;
    }

    private static Cell GetOrCreateCell(Dictionary<int, Dictionary<int, Cell>> cells, int row, int col,
        ref int maxRow, ref int maxCol)
    {
        if (!cells.TryGetValue(row, out var rowCells))
        {
            rowCells = new Dictionary<int, Cell>();
            cells[row] = rowCells;
        }
        if (!rowCells.TryGetValue(col, out var cell))
        {
            cell = Cell.Empty;
            rowCells[col] = cell;
        }
        if (row > maxRow) maxRow = row;
        if (col > maxCol) maxCol = col;
        return cell;
    }

    private static void PutCell(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d,
        int rowOff, int colOff, int valueOff, Func<int, int, byte[], Cell> factory,
        ref int maxRow, ref int maxCol)
    {
        int row = BiffRecords.ReadU16(d, rowOff);
        int col = BiffRecords.ReadU16(d, colOff);
        if (!cells.TryGetValue(row, out var rowCells))
        {
            rowCells = new Dictionary<int, Cell>();
            cells[row] = rowCells;
        }
        rowCells[col] = factory(row, col, d);
        if (row > maxRow) maxRow = row;
        if (col > maxCol) maxCol = col;
    }

    private static void ParseMulRk(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d,
        List<int> xfIfmt, Dictionary<int, string> formats, bool date1904, ref int maxRow, ref int maxCol)
    {
        if (d.Length < 6) return;
        int row = BiffRecords.ReadU16(d, 0);
        int colFirst = BiffRecords.ReadU16(d, 2);
        // 数据区: (ixfe(2), rk(4)) * n, 最后 colLast(2)
        int n = (d.Length - 6) / 6;
        for (int k = 0; k < n; k++)
        {
            int ixfe = BiffRecords.ReadU16(d, 4 + k * 6);
            int rk = ReadS32At(d, 6 + k * 6);
            int col = colFirst + k;
            double val = BiffShared.DecodeRk(rk);
            var cell = CellFromNumber(val, ixfe, xfIfmt, formats, date1904);
            if (!cells.TryGetValue(row, out var rowCells))
            {
                rowCells = new Dictionary<int, Cell>();
                cells[row] = rowCells;
            }
            rowCells[col] = cell;
            if (row > maxRow) maxRow = row;
            if (col > maxCol) maxCol = col;
        }
    }

    private static void ParseFormula(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d,
        List<int> xfIfmt, Dictionary<int, string> formats, bool date1904, ref int maxRow, ref int maxCol)
    {
        if (d.Length < 14) return;
        int row = BiffRecords.ReadU16(d, 0);
        int col = BiffRecords.ReadU16(d, 2);
        int ixfe = BiffRecords.ReadU16(d, 4);

        Cell cell;
        // 值字段 8 字节（offset 6..13）
        bool isSpecial = d[6] == 0xFF && d[7] == 0xFF;
        if (isSpecial)
        {
            byte resultType = d[8];
            switch (resultType)
            {
                case 0x00: // 字符串结果
                    cell = Cell.FromText("");
                    break;
                case 0x01: // 布尔
                    cell = Cell.FromBoolean(d[9] != 0);
                    break;
                case 0x02: // 错误
                    cell = Cell.FromText(BiffShared.ErrorCode(d[9]));
                    break;
                default: // 0x03 空
                    cell = Cell.Empty;
                    break;
            }
        }
        else
        {
            var bits = BitConverter.ToInt64(d, 6);
            double val = BitConverter.Int64BitsToDouble(bits);
            cell = CellFromNumber(val, ixfe, xfIfmt, formats, date1904);
        }

        // 公式 RPN → 文本（FORMULA = 头部 22 字节 + cce(2) + RPN）
        if (d.Length >= 24)
        {
            int cce = BiffRecords.ReadU16(d, 20);
            if (cce > 0 && 22 + cce <= d.Length)
            {
                var rpn = new byte[cce];
                Array.Copy(d, 22, rpn, 0, cce);
                var text = FormulaParser.Parse(rpn, biff12: false);
                if (!string.IsNullOrEmpty(text))
                {
                    // P0-8: 公式串放入 Formula，不覆盖缓存值
                    cell.IsFormula = true;
                    cell.Formula = text;
                }
            }
        }

        PutCell(cells, d, 0, 2, -1, (r, c, dd) => cell, ref maxRow, ref maxCol);
    }

    // ── 单元格类型辅助 ──

    private static Cell CellFromNumber(double val, int ixfe, List<int> xfIfmt,
        Dictionary<int, string> formats, bool date1904)
        => FormatDetector.CellFromNumber(val, ixfe, xfIfmt, formats, date1904);

    private static string ParseLabelString(byte[] d)
    {
        // rw(2) col(2) ixfe(2) cch(2) grbit(1) data
        if (d.Length < 9) return "";
        int cch = BiffRecords.ReadU16(d, 6);
        bool highByte = (d[8] & 0x01) != 0;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, d.Length - 9);
            return Encoding.Unicode.GetString(d, 9, bytes);
        }
        int latinBytes = Math.Min(cch, d.Length - 9);
        return Latin1.GetString(d, 9, latinBytes);
    }

    private static void ParseMergedCells(SheetData sheet, byte[] d)
    {
        if (d.Length < 2) return;
        int count = BiffRecords.ReadU16(d, 0);
        for (int k = 0; k < count; k++)
        {
            int off = 2 + k * 8;
            if (off + 8 > d.Length) break;
            int rwFirst = BiffRecords.ReadU16(d, off);
            int rwLast = BiffRecords.ReadU16(d, off + 2);
            int colFirst = BiffRecords.ReadU16(d, off + 4);
            int colLast = BiffRecords.ReadU16(d, off + 6);
            sheet.MergedRanges.Add(new CellRange(rwFirst, rwLast, colFirst, colLast));
        }
    }

    private static void ParseColInfo(byte[] d, Dictionary<int, double> colWidths)
    {
        if (d.Length < 6) return;
        int colFirst = BiffRecords.ReadU16(d, 0);
        int colLast = BiffRecords.ReadU16(d, 2);
        int colw = BiffRecords.ReadU16(d, 4);
        double width = colw / 256.0;
        for (int c = colFirst; c <= colLast; c++)
            colWidths[c] = width;
    }

    private static void ParseRowHeight(byte[] d, Dictionary<int, double> rowHeights)
    {
        if (d.Length < 8) return;
        int rw = BiffRecords.ReadU16(d, 0);
        int miyRw = BiffRecords.ReadU16(d, 6); // BIFF8 ROW：rw(0)+colMic(2)+colMac(4)+miyRw(6)
        if (miyRw == 0 || miyRw == 0xFF) return; // 默认行高
        rowHeights[rw] = miyRw / 20.0; // 缇 → 磅
    }

    private static ushort ReadU16At(byte[] d, int offset) => BiffRecords.ReadU16(d, offset);
    private static int ReadS32At(byte[] d, int offset) => BiffRecords.ReadS32(d, offset);
}
