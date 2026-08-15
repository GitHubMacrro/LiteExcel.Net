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

    public static List<SheetData> ReadAll(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadAll(fs);
    }

    public static List<SheetData> ReadAll(Stream stream)
    {
        var cfb = CfbFile.Open(stream);
        var workbook = cfb.GetStream("Workbook") ?? cfb.GetStream("Book");
        if (workbook is null)
            throw new LiteExcelException(".xls 文件中缺少 Workbook/Book 流");
        return ParseWorkbook(workbook);
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

        int i = 1;
        for (; i < records.Count; i++)
        {
            var rec = records[i];
            if (rec.Opcode == BiffRecords.OpEof) { i++; break; }

            switch (rec.Opcode)
            {
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
            }
        }

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
                case BiffRecords.OpColInfo:
                    ParseColInfo(rec.Data, colWidths);
                    break;
                case BiffRecords.OpRow:
                    ParseRowHeight(rec.Data, rowHeights);
                    break;
                case BiffRecords.OpPane:
                    if (rec.Data.Length >= 6)
                    {
                        int ySplit = BiffRecords.ReadU16(rec.Data, 2);
                        int xSplit = BiffRecords.ReadU16(rec.Data, 0);
                        if (ySplit >= 1 || xSplit >= 1) freeze = true;
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
                    cell = Cell.FromText(""); // 公式文本暂不解析（后续可扩展）
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
        if (d.Length < 6) return;
        int rw = BiffRecords.ReadU16(d, 0);
        int miyRw = BiffRecords.ReadU16(d, 4);
        if (miyRw == 0 || miyRw == 0xFF) return; // 默认行高
        rowHeights[rw] = miyRw / 20.0; // 缇 → 磅
    }

    private static ushort ReadU16At(byte[] d, int offset) => BiffRecords.ReadU16(d, offset);
    private static int ReadS32At(byte[] d, int offset) => BiffRecords.ReadS32(d, offset);
}
