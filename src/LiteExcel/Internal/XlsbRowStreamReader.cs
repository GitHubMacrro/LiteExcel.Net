using System;
using System.Collections.Generic;
using System.IO;
using LiteExcel.Internal.Biff12;

namespace LiteExcel.Internal;

/// <summary>
/// XLSB BIFF12 记录级流式读取器。
/// 直接从已解压的 sheetN.bin 字节数组逐条读取记录，遇到 BrtRowHdr 时 yield 前一行，
/// 避免一次性 ReadAll 将整个 sheetN.bin 加载到字典后再组装。
/// 共享字符串和样式表仍需预加载（工作簿级共享），但工作表数据不再驻留内存。
/// </summary>
internal static class XlsbRowStreamReader
{
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

    /// <summary>从已解压的工作表字节数组流式读取行。</summary>
    public static IEnumerable<IReadOnlyList<Cell>> EnumerateRows(
        byte[] data, List<string> sst,
        Dictionary<int, string> formats, List<int> cellXfs, bool date1904)
    {
        var rowCells = new Dictionary<int, Cell>();
        bool hasRow = false;
        int prevCol = -1;
        int pos = 0;
        int len = data.Length;

        while (pos < len)
        {
            int rt = Biff12Records.ReadVarInt(data, ref pos);
            int cb = Biff12Records.ReadVarInt(data, ref pos);
            if (cb < 0 || pos + cb > len) break;

            var d = new byte[cb];
            Array.Copy(data, pos, d, 0, cb);
            pos += cb;

            bool isShort = rt >= BrtShortBlank;

            switch (rt)
            {
                case BrtRowHdr:
                    if (hasRow)
                        yield return BuildRow(rowCells);
                    rowCells.Clear();
                    prevCol = -1;
                    hasRow = true;
                    break;

                case BrtCellBlank:
                case BrtShortBlank:
                    PutCell(rowCells, d, isShort, ref prevCol, _ => Cell.Empty);
                    break;

                case BrtCellRk:
                case BrtShortRk:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff => FormatDetector.CellFromNumber(
                            BiffShared.DecodeRk(Biff12Records.ReadS32(d, valOff)),
                            StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904));
                    break;

                case BrtCellError:
                case BrtShortError:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff => Cell.FromText(BiffShared.ErrorCode(d[valOff])));
                    break;

                case BrtCellBool:
                case BrtShortBool:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff => Cell.FromBoolean(d[valOff] != 0));
                    break;

                case BrtCellReal:
                case BrtShortReal:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff => FormatDetector.CellFromNumber(
                            BitConverter.ToDouble(d, valOff),
                            StyleRef(d, isShort ? 0 : 4), cellXfs, formats, date1904));
                    break;

                case BrtCellSt:
                case BrtShortSt:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff => Cell.FromText(ReadStringAt(d, valOff)));
                    break;

                case BrtCellIsst:
                case BrtShortIsst:
                    PutCell(rowCells, d, isShort, ref prevCol,
                        valOff =>
                        {
                            int idx = Biff12Records.ReadS32(d, valOff);
                            return idx >= 0 && idx < sst.Count ? Cell.FromText(sst[idx]) : Cell.Empty;
                        });
                    break;

                case BrtFmlaString:
                    PutCell(rowCells, d, false, ref prevCol,
                        valOff => Cell.FromText(ReadStringAt(d, valOff)));
                    break;

                case BrtFmlaNum:
                    PutCell(rowCells, d, false, ref prevCol,
                        valOff => FormatDetector.CellFromNumber(
                            BitConverter.ToDouble(d, valOff),
                            StyleRef(d, 4), cellXfs, formats, date1904));
                    break;

                case BrtFmlaBool:
                    PutCell(rowCells, d, false, ref prevCol,
                        valOff => Cell.FromBoolean(d[valOff] != 0));
                    break;

                case BrtFmlaError:
                    PutCell(rowCells, d, false, ref prevCol,
                        valOff => Cell.FromText(BiffShared.ErrorCode(d[valOff])));
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

    private static void PutCell(Dictionary<int, Cell> cells, byte[] d, bool shortCell,
        ref int prevCol, Func<int, Cell> factory)
    {
        int valueOff = shortCell ? 4 : 8;
        int col;
        if (shortCell)
            col = prevCol + 1;
        else
        {
            col = Biff12Records.ReadS32(d, 0);
            if (col < 0) col = 0;
        }
        prevCol = col;
        cells[col] = factory(valueOff);
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
}
