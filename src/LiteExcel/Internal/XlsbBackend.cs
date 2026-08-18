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

    // sharedStrings.bin
    private const int BrtBeginSst = 0x009F;
    private const int BrtSSTItem = 0x0013;
    private const int BrtEndSst = 0x00A0;

    // styles.bin
    private const int BrtFmt = 0x002C;        // 自定义数字格式
    private const int BrtXf = 0x002F;         // 单元格样式 XF
    private const int BrtBeginCellXfs = 0x0269;
    private const int BrtEndCellXfs = 0x026A;

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

    public static List<SheetData> ReadAll(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);

        var wbBytes = ReadEntry(zip, "xl/workbook.bin")
            ?? throw new LiteExcelException(".xlsb 文件中缺少 xl/workbook.bin");
        var (sheets, date1904) = ParseWorkbook(wbBytes);

        var sstBytes = ReadEntry(zip, "xl/sharedStrings.bin");
        var sst = sstBytes is not null ? ParseSharedStrings(sstBytes) : new List<string>();

        var stylesBytes = ReadEntry(zip, "xl/styles.bin");
        var (formats, cellXfs) = stylesBytes is not null
            ? ParseStyles(stylesBytes)
            : (new Dictionary<int, string>(), new List<int> { 0 });

        var sheetPaths = MapSheetPaths(zip, sheets);

        var result = new List<SheetData>(sheets.Count);
        for (int i = 0; i < sheets.Count; i++)
        {
            var data = ReadEntry(zip, sheetPaths[i]);
            if (data is null)
                throw new LiteExcelException($"缺少工作表文件: {sheetPaths[i]}");
            result.Add(ParseWorksheet(data, sheets[i].Name, sst, formats, cellXfs, date1904));
        }

        if (result.Count == 0)
            throw new LiteExcelException("这不是有效的 .xlsb 文件（未找到任何工作表）");
        return result;
    }

    // ── 包部件 ──

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
    private static List<string> MapSheetPaths(ZipArchive zip, List<(string Name, string RelId)> sheets)
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

    // ── workbook.bin ──

    private static (List<(string Name, string RelId)> Sheets, bool Date1904) ParseWorkbook(byte[] wb)
    {
        var records = Biff12Records.ReadAll(wb);
        if (records.Count == 0)
            throw new LiteExcelException("这不是有效的 .xlsb 文件（workbook.bin 为空）");

        var sheets = new List<(string, string)>();
        bool date1904 = false;

        foreach (var rec in records)
        {
            switch (rec.Rt)
            {
                case BrtBundleSh:
                {
                    var d = rec.Data;
                    if (d.Length < 8) break;
                    int off = 8; // Hidden(4) + iTabID(4)
                    var relId = Biff12Records.ReadWideString(d, ref off);
                    var name = Biff12Records.ReadWideString(d, ref off);
                    sheets.Add((name, relId));
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

    // ── sharedStrings.bin ──

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

    // ── styles.bin ──

    private static (Dictionary<int, string> Formats, List<int> CellXfs) ParseStyles(byte[] data)
    {
        var records = Biff12Records.ReadAll(data);
        var formats = new Dictionary<int, string>();
        var cellXfs = new List<int>(); // BrtBeginCellXFs 内按序排列，索引 0 即默认样式
        bool inCellXfs = false;

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
                case BrtBeginCellXfs:
                    inCellXfs = true;
                    break;
                case BrtEndCellXfs:
                    inCellXfs = false;
                    break;
                case BrtXf:
                    if (inCellXfs && d.Length >= 4)
                        cellXfs.Add(Biff12Records.ReadU16(d, 2)); // ixfeParent(2) ifmt(2)
                    break;
            }
        }
        return (formats, cellXfs);
    }

    // ── worksheet.bin ──

    private static SheetData ParseWorksheet(byte[] data, string sheetName, List<string> sst,
        Dictionary<int, string> formats, List<int> cellXfs, bool date1904)
    {
        var records = Biff12Records.ReadAll(data);
        var sheet = new SheetData { SheetName = sheetName };
        var cells = new Dictionary<int, Dictionary<int, Cell>>();
        int maxRow = -1;
        int maxCol = -1;
        var colWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();
        bool freeze = false;
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
                    PutCell(cells, d, isShort, ref prevCol, currentRow, (_) => Cell.Empty, ref maxRow, ref maxCol);
                    break;
                case BrtCellRk:
                case BrtShortRk:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => FormatDetector.CellFromNumber(BiffShared.DecodeRk(ReadS32(d, valOff)), StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904),
                        ref maxRow, ref maxCol);
                    break;
                case BrtCellError:
                case BrtShortError:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromText(BiffShared.ErrorCode(d[valOff])), ref maxRow, ref maxCol);
                    break;
                case BrtCellBool:
                case BrtShortBool:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromBoolean(d[valOff] != 0), ref maxRow, ref maxCol);
                    break;
                case BrtCellReal:
                case BrtShortReal:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => FormatDetector.CellFromNumber(BitConverter.ToDouble(d, valOff), StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904),
                        ref maxRow, ref maxCol);
                    break;
                case BrtCellSt:
                case BrtShortSt:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) => Cell.FromText(ReadStringAt(d, valOff)), ref maxRow, ref maxCol);
                    break;
                case BrtCellIsst:
                case BrtShortIsst:
                    PutCell(cells, d, isShort, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            int idx = ReadS32(d, valOff);
                            return idx >= 0 && idx < sst.Count ? Cell.FromText(sst[idx]) : Cell.Empty;
                        }, ref maxRow, ref maxCol);
                    break;
                case BrtFmlaString:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromText(ReadStringAt(d, valOff));
                            ApplyFormula(d, valOff + 4 + ReadWideLen(d, valOff), cell);
                            return cell;
                        }, ref maxRow, ref maxCol);
                    break;
                case BrtFmlaNum:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = FormatDetector.CellFromNumber(BitConverter.ToDouble(d, valOff), StyleRef(d, 4), cellXfs, formats, date1904);
                            ApplyFormula(d, valOff + 8, cell);
                            return cell;
                        }, ref maxRow, ref maxCol);
                    break;
                case BrtFmlaBool:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromBoolean(d[valOff] != 0);
                            ApplyFormula(d, valOff + 1, cell);
                            return cell;
                        }, ref maxRow, ref maxCol);
                    break;
                case BrtFmlaError:
                    PutCell(cells, d, false, ref prevCol, currentRow,
                        (valOff) =>
                        {
                            var cell = Cell.FromText(BiffShared.ErrorCode(d[valOff]));
                            ApplyFormula(d, valOff + 1, cell);
                            return cell;
                        }, ref maxRow, ref maxCol);
                    break;
                case BrtColInfo:
                    ParseColInfo(d, colWidths);
                    break;
                case BrtMergeCell:
                    ParseMergeCell(d, sheet);
                    break;
                case BrtPane:
                    // colFrozen(Xnum 8) + rowFrozen(Xnum 8)
                    if (d.Length >= 16)
                    {
                        double colFrozen = BitConverter.ToDouble(d, 0);
                        double rowFrozen = BitConverter.ToDouble(d, 8);
                        if (colFrozen >= 1.0 || rowFrozen >= 1.0) freeze = true;
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

        sheet.FreezeHeader = freeze;
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

    private static void PutCell(Dictionary<int, Dictionary<int, Cell>> cells, byte[] d, bool shortCell,
        ref int prevCol, int currentRow, Func<int, Cell> factory, ref int maxRow, ref int maxCol)
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
            cell.IsFormula = true;
            cell.Text = text;
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

    private static int ReadS32(byte[] d, int off) => Biff12Records.ReadS32(d, off);
}
