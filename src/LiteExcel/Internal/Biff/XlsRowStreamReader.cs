using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// XLS BIFF8 记录级流式读取器。
/// 从已解压的 Workbook 字节数组逐条读取记录，在目标工作表段内遇到单元格记录时
/// 按行号分组 yield，避免一次性构建全表字典后再组装行数组。
/// 共享字符串和样式表仍需预加载（工作簿级共享），但工作表行数据不再驻留内存。
/// </summary>
internal static class XlsRowStreamReader
{
    /// <summary>
    /// 预扫描全局段，返回 (boundSheets, sst, formats, xfIfmt, date1904, sheetStartPositions)。
    /// sheetStartPositions[i] = 第 i 张表的 BOF 记录在 workbookBytes 中的字节偏移。
    /// </summary>
    public static (List<string> boundSheets, List<string> sst,
        Dictionary<int, string> formats, List<int> xfIfmt, bool date1904,
        List<int> sheetStartPositions)
        PrepareStreaming(byte[] wb)
    {
        var boundSheets = new List<string>();
        var sst = new List<string>();
        var formats = new Dictionary<int, string>();
        var xfIfmt = new List<int>();
        var sheetStartPositions = new List<int>();
        bool date1904 = false;

        int pos = 0;
        int len = wb.Length;
        bool inGlobalSection = true;

        while (pos + 4 <= len)
        {
            ushort opcode = BiffRecords.ReadU16(wb, pos);
            ushort size = BiffRecords.ReadU16(wb, pos + 2);
            int dataStart = pos + 4;
            if (dataStart + size > len)
                size = (ushort)(len - dataStart);

            if (inGlobalSection)
            {
                switch (opcode)
                {
                    case BiffRecords.OpEof:
                        inGlobalSection = false;
                        break;
                    case BiffRecords.OpFilePass:
                        throw new LiteExcelException("该 .xls 文件已加密（带打开密码）。当前版本暂不支持读取加密工作簿，" +
                            "请在 Excel 中另存为无密码文件后再打开。");
                    case BiffRecords.OpBoundSheet:
                        boundSheets.Add(ParseBoundSheetName(wb, dataStart, size));
                        break;
                    case BiffRecords.OpSst:
                        ParseSstStreaming(wb, dataStart, size, ref pos, sst);
                        continue;
                    case BiffRecords.OpFormat:
                        ParseFormat(wb, dataStart, size, formats);
                        break;
                    case BiffRecords.OpXf:
                        if (size >= 4)
                            xfIfmt.Add(BiffRecords.ReadU16(wb, dataStart + 2));
                        break;
                    case BiffRecords.OpDateMode:
                        if (size >= 2)
                            date1904 = BiffRecords.ReadU16(wb, dataStart) == 1;
                        break;
                }
            }
            else
            {
                if (opcode == BiffRecords.OpBof)
                    sheetStartPositions.Add(pos);
            }

            pos = dataStart + size;
        }

        return (boundSheets, sst, formats, xfIfmt, date1904, sheetStartPositions);
    }

    /// <summary>从指定工作表起始位置流式读取行。</summary>
    public static IEnumerable<IReadOnlyList<Cell>> EnumerateRows(
        byte[] wb, int sheetStartPosition,
        List<string> sst, Dictionary<int, string> formats, List<int> xfIfmt, bool date1904)
    {
        var rowCells = new Dictionary<int, Cell>();
        int currentRow = -1;
        bool hasRow = false;

        int pos = sheetStartPosition;
        int len = wb.Length;

        while (pos + 4 <= len)
        {
            ushort opcode = BiffRecords.ReadU16(wb, pos);
            ushort size = BiffRecords.ReadU16(wb, pos + 2);
            int dataStart = pos + 4;
            if (dataStart + size > len)
                size = (ushort)(len - dataStart);

            var d = new byte[size];
            Array.Copy(wb, dataStart, d, 0, size);
            pos = dataStart + size;

            if (opcode == BiffRecords.OpEof)
                break;
            if (opcode == BiffRecords.OpBof)
                continue;

            int rowFromRecord;
            Cell cell;

            switch (opcode)
            {
                case BiffRecords.OpNumber:
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    cell = CellFromNumber(BitConverter.ToDouble(d, 6),
                        BiffRecords.ReadU16(d, 4), xfIfmt, formats, date1904);
                    PutCell(rowCells, d, 0, 2, cell);
                    break;

                case BiffRecords.OpRk:
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    cell = Cell.FromNumber(BiffShared.DecodeRk(BiffRecords.ReadS32(d, 6)));
                    PutCell(rowCells, d, 0, 2, cell);
                    break;

                case BiffRecords.OpMulRk:
                    if (d.Length < 6) break;
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    int colFirst = BiffRecords.ReadU16(d, 2);
                    int n = (d.Length - 6) / 6;
                    for (int k = 0; k < n; k++)
                    {
                        int ixfe = BiffRecords.ReadU16(d, 4 + k * 6);
                        int rk = BiffRecords.ReadS32(d, 6 + k * 6);
                        int col = colFirst + k;
                        rowCells[col] = CellFromNumber(BiffShared.DecodeRk(rk), ixfe, xfIfmt, formats, date1904);
                    }
                    break;

                case BiffRecords.OpLabelSst:
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    int isst = BiffRecords.ReadS32(d, 6);
                    cell = isst >= 0 && isst < sst.Count ? Cell.FromText(sst[isst]) : Cell.Empty;
                    PutCell(rowCells, d, 0, 2, cell);
                    break;

                case BiffRecords.OpLabel:
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    cell = Cell.FromText(ParseLabelString(d));
                    PutCell(rowCells, d, 0, 2, cell);
                    break;

                case BiffRecords.OpBoolErr:
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    bool isError = d.Length > 7 && d[7] != 0;
                    if (isError)
                        cell = Cell.FromText(BiffShared.ErrorCode(d.Length > 6 ? d[6] : (byte)0));
                    else
                        cell = Cell.FromBoolean(d.Length > 6 && d[6] != 0);
                    PutCell(rowCells, d, 0, 2, cell);
                    break;

                case BiffRecords.OpFormula:
                    if (d.Length < 14) break;
                    rowFromRecord = BiffRecords.ReadU16(d, 0);
                    if (rowFromRecord != currentRow)
                    {
                        if (hasRow)
                            yield return BuildRow(rowCells);
                        rowCells.Clear();
                        currentRow = rowFromRecord;
                        hasRow = true;
                    }
                    int ixfeF = BiffRecords.ReadU16(d, 4);
                    cell = ParseFormulaCell(d, ixfeF, xfIfmt, formats, date1904);
                    PutCell(rowCells, d, 0, 2, cell);
                    break;
            }
        }

        if (hasRow)
            yield return BuildRow(rowCells);
    }

    private static IReadOnlyList<Cell> BuildRow(Dictionary<int, Cell> rowCells)
    {
        if (rowCells.Count == 0)
            return Array.Empty<Cell>();
        int maxCol = -1;
        foreach (var col in rowCells.Keys)
            if (col > maxCol) maxCol = col;
        var arr = new Cell[maxCol + 1];
        for (int i = 0; i <= maxCol; i++)
            arr[i] = Cell.Empty;
        foreach (var kv in rowCells)
            arr[kv.Key] = kv.Value;
        return arr;
    }

    private static void PutCell(Dictionary<int, Cell> cells, byte[] d, int rowOff, int colOff, Cell cell)
    {
        int col = BiffRecords.ReadU16(d, colOff);
        cells[col] = cell;
    }

    private static Cell CellFromNumber(double val, int ixfe, List<int> xfIfmt,
        Dictionary<int, string> formats, bool date1904)
        => FormatDetector.CellFromNumber(val, ixfe, xfIfmt, formats, date1904);

    private static Cell ParseFormulaCell(byte[] d, int ixfe, List<int> xfIfmt,
        Dictionary<int, string> formats, bool date1904)
    {
        Cell cell;
        bool isSpecial = d[6] == 0xFF && d[7] == 0xFF;
        if (isSpecial)
        {
            byte resultType = d[8];
            cell = resultType switch
            {
                0x00 => Cell.FromText(""),
                0x01 => Cell.FromBoolean(d[9] != 0),
                0x02 => Cell.FromText(BiffShared.ErrorCode(d[9])),
                _ => Cell.Empty,
            };
        }
        else
        {
            var bits = BitConverter.ToInt64(d, 6);
            cell = CellFromNumber(BitConverter.Int64BitsToDouble(bits), ixfe, xfIfmt, formats, date1904);
        }
        return cell;
    }

    private static string ParseBoundSheetName(byte[] wb, int dataStart, int size)
    {
        if (size < 8) return "";
        int cch = wb[dataStart + 6];
        bool highByte = (wb[dataStart + 7] & 0x01) != 0;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, size - 8);
            return Encoding.Unicode.GetString(wb, dataStart + 8, bytes);
        }
        int latinBytes = Math.Min(cch, size - 8);
        return Encoding.GetEncoding(28591).GetString(wb, dataStart + 8, latinBytes);
    }

    private static void ParseFormat(byte[] wb, int dataStart, int size, Dictionary<int, string> formats)
    {
        if (size < 5) return;
        int ifmt = BiffRecords.ReadU16(wb, dataStart);
        int cch = BiffRecords.ReadU16(wb, dataStart + 2);
        bool highByte = (wb[dataStart + 4] & 0x01) != 0;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, size - 5);
            formats[ifmt] = Encoding.Unicode.GetString(wb, dataStart + 5, bytes);
        }
        else
        {
            int latinBytes = Math.Min(cch, size - 5);
            formats[ifmt] = Encoding.GetEncoding(28591).GetString(wb, dataStart + 5, latinBytes);
        }
    }

    private static void ParseSstStreaming(byte[] wb, int dataStart, int sstSize, ref int pos, List<string> sst)
    {
        if (sstSize < 8) { pos = dataStart + sstSize; return; }
        int uniqueCount = BiffRecords.ReadS32(wb, dataStart + 4);

        var segments = new List<byte[]>();
        var firstSegment = new byte[sstSize - 8];
        Array.Copy(wb, dataStart + 8, firstSegment, 0, firstSegment.Length);
        segments.Add(firstSegment);

        pos = dataStart + sstSize;
        while (pos + 4 <= wb.Length)
        {
            ushort cont = BiffRecords.ReadU16(wb, pos);
            ushort contSize = BiffRecords.ReadU16(wb, pos + 2);
            if (cont != BiffRecords.OpContinue) break;
            var segData = new byte[contSize];
            Array.Copy(wb, pos + 4, segData, 0, contSize);
            segments.Add(segData);
            pos += 4 + contSize;
        }

        var reader = new BiffStringReader(segments);
        for (int k = 0; k < uniqueCount; k++)
        {
            var s = reader.ReadString();
            if (s is null) break;
            sst.Add(s);
        }
    }

    private static string ParseLabelString(byte[] d)
    {
        if (d.Length < 9) return "";
        int cch = BiffRecords.ReadU16(d, 6);
        bool highByte = (d[8] & 0x01) != 0;
        if (highByte)
        {
            int bytes = Math.Min(cch * 2, d.Length - 9);
            return Encoding.Unicode.GetString(d, 9, bytes);
        }
        int latinBytes = Math.Min(cch, d.Length - 9);
        return Encoding.GetEncoding(28591).GetString(d, 9, latinBytes);
    }
}
