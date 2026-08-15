using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 程序化构造最小 .xlsb（BIFF12）文件。记录头采用 [MS-XLSB] 变长编码：rt(LEB128) + cb(LEB128)。
/// 仅覆盖读取器关心的部件：workbook.bin / sharedStrings.bin / styles.bin / sheetN.bin。
/// </summary>
internal static class XlsbTestFile
{
    // workbook.bin
    private const int BrtBeginBook = 0x0083;
    private const int BrtEndBook = 0x0084;
    private const int BrtWbProp = 0x0099;
    private const int BrtBeginBundleShs = 0x008F;
    private const int BrtEndBundleShs = 0x0090;
    private const int BrtBundleSh = 0x009C;

    // sharedStrings.bin
    private const int BrtBeginSst = 0x009F;
    private const int BrtSSTItem = 0x0013;
    private const int BrtEndSst = 0x00A0;

    // styles.bin
    private const int BrtFmt = 0x002C;
    private const int BrtBeginCellXfs = 0x0269;
    private const int BrtEndCellXfs = 0x026A;
    private const int BrtXf = 0x002F;

    // worksheet.bin
    private const int BrtBeginSheet = 0x0081;
    private const int BrtEndSheet = 0x0082;
    private const int BrtWsDim = 0x0094;
    private const int BrtPane = 0x0097;
    private const int BrtBeginSheetData = 0x0091;
    private const int BrtEndSheetData = 0x0092;
    private const int BrtRowHdr = 0x0000;
    private const int BrtCellBlank = 0x0001;
    private const int BrtCellRk = 0x0002;
    private const int BrtCellBool = 0x0004;
    private const int BrtCellReal = 0x0005;
    private const int BrtCellSt = 0x0006;
    private const int BrtCellIsst = 0x0007;
    private const int BrtColInfo = 0x003C;
    private const int BrtMergeCell = 0x00B0;
    private const int BrtBeginMergeCells = 0x00B1;
    private const int BrtEndMergeCells = 0x00B2;

    public sealed class CellSpec
    {
        public int Col;
        public int Style = -1; // -1 = 默认样式 0
        public string? Text;
        public double? Number;
        public bool? Bool;
        public bool InlineText;
    }

    public sealed class RowSpec
    {
        public List<CellSpec> Cells { get; } = new();
        public double? Height;
    }

    public sealed class SheetSpec
    {
        public string Name = "Sheet1";
        public List<RowSpec> Rows { get; } = new();
        public List<(int R1, int R2, int C1, int C2)> Merges { get; } = new();
        public Dictionary<int, double> ColWidths { get; } = new();
        public int FrozenRows;
        public int FrozenCols;
    }

    public sealed class WorkbookSpec
    {
        public List<SheetSpec> Sheets { get; } = new();
        public List<string> SharedStrings { get; } = new();
        public Dictionary<int, string> Formats { get; } = new();
        public List<int> CellXfs { get; } = new() { 0 }; // 索引 0 = 默认样式
        public bool Date1904;
    }

    public static string Build(WorkbookSpec spec)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(zip, "xl/workbook.bin", BuildWorkbook(spec));
            // 先写工作表：WriteCell 会向 spec.SharedStrings 注册新字符串，随后再序列化 SST
            for (int i = 0; i < spec.Sheets.Count; i++)
                WriteEntry(zip, $"xl/worksheets/sheet{i + 1}.bin", BuildWorksheet(spec, spec.Sheets[i]));
            WriteEntry(zip, "xl/sharedStrings.bin", BuildSharedStrings(spec.SharedStrings));
            WriteEntry(zip, "xl/styles.bin", BuildStyles(spec));
        }
        return TempFile(ms.ToArray());
    }

    private static string TempFile(byte[] bytes)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsbtest_{System.Guid.NewGuid():N}.xlsb");
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private static void WriteEntry(ZipArchive zip, string name, byte[] content)
    {
        var entry = zip.CreateEntry(name);
        using var s = entry.Open();
        s.Write(content, 0, content.Length);
    }

    private static byte[] BuildWorkbook(WorkbookSpec spec)
    {
        using var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginBook, Empty());
        // BrtWbProp: flags(4) defaultThemeVersion(4) [CodeName]
        var wbProp = new byte[8];
        if (spec.Date1904) wbProp[0] = 0x01;
        WriteRecord(ms, BrtWbProp, wbProp);
        WriteRecord(ms, BrtBeginBundleShs, Empty());
        foreach (var s in spec.Sheets)
        {
            using var b = new MemoryStream();
            WriteU32(b, 0);          // Hidden
            WriteU32(b, 1);          // iTabID
            WriteWideString(b, "rId1");
            WriteWideString(b, s.Name);
            WriteRecord(ms, BrtBundleSh, b.ToArray());
        }
        WriteRecord(ms, BrtEndBundleShs, Empty());
        WriteRecord(ms, BrtEndBook, Empty());
        return ms.ToArray();
    }

    private static byte[] BuildSharedStrings(List<string> strings)
    {
        using var ms = new MemoryStream();
        using (var b = new MemoryStream())
        {
            WriteU32(b, (uint)strings.Count);
            WriteU32(b, (uint)strings.Count);
            WriteRecord(ms, BrtBeginSst, b.ToArray());
        }
        foreach (var s in strings)
        {
            using var b = new MemoryStream();
            b.WriteByte(0); // flags: 非富文本
            WriteWideString(b, s);
            WriteRecord(ms, BrtSSTItem, b.ToArray());
        }
        WriteRecord(ms, BrtEndSst, Empty());
        return ms.ToArray();
    }

    private static byte[] BuildStyles(WorkbookSpec spec)
    {
        using var ms = new MemoryStream();
        foreach (var kv in spec.Formats)
        {
            using var b = new MemoryStream();
            WriteU16(b, (ushort)kv.Key);
            WriteWideString(b, kv.Value);
            WriteRecord(ms, BrtFmt, b.ToArray());
        }
        WriteRecord(ms, BrtBeginCellXfs, Empty());
        foreach (var ifmt in spec.CellXfs)
        {
            using var b = new MemoryStream();
            WriteU16(b, 0);                  // ixfeParent
            WriteU16(b, (ushort)ifmt);       // ifmt
            for (int i = 0; i < 12; i++) b.WriteByte(0);
            WriteRecord(ms, BrtXf, b.ToArray());
        }
        WriteRecord(ms, BrtEndCellXfs, Empty());
        return ms.ToArray();
    }

    private static byte[] BuildWorksheet(WorkbookSpec spec, SheetSpec sheet)
    {
        using var ms = new MemoryStream();
        WriteRecord(ms, BrtBeginSheet, Empty());
        WriteRecord(ms, BrtWsDim, Empty());

        if (sheet.FrozenRows > 0 || sheet.FrozenCols > 0)
        {
            using var b = new MemoryStream();
            WriteDouble(b, sheet.FrozenCols); // xSplit Xnum
            WriteDouble(b, sheet.FrozenRows); // ySplit Xnum
            for (int i = 0; i < 13; i++) b.WriteByte(0);
            WriteRecord(ms, BrtPane, b.ToArray());
        }

        if (sheet.Merges.Count > 0)
        {
            using (var b = new MemoryStream())
            {
                WriteU32(b, (uint)sheet.Merges.Count);
                WriteRecord(ms, BrtBeginMergeCells, b.ToArray());
            }
            foreach (var m in sheet.Merges)
            {
                using var b = new MemoryStream();
                WriteU32(b, (uint)m.R1);
                WriteU32(b, (uint)m.R2);
                WriteU32(b, (uint)m.C1);
                WriteU32(b, (uint)m.C2);
                WriteRecord(ms, BrtMergeCell, b.ToArray());
            }
            WriteRecord(ms, BrtEndMergeCells, Empty());
        }

        WriteRecord(ms, BrtBeginSheetData, Empty());
        for (int r = 0; r < sheet.Rows.Count; r++)
        {
            WriteRowHeader(ms, r, sheet.Rows[r].Height);
            int prevCol = -1;
            foreach (var cell in sheet.Rows[r].Cells)
            {
                WriteCell(ms, spec, cell, prevCol);
                prevCol = cell.Col;
            }
        }
        WriteRecord(ms, BrtEndSheetData, Empty());

        if (sheet.ColWidths.Count > 0)
        {
            foreach (var kv in sheet.ColWidths)
            {
                using var b = new MemoryStream();
                WriteU32(b, (uint)kv.Key);
                WriteU32(b, (uint)kv.Key);
                WriteU32(b, (uint)(kv.Value * 256));
                WriteU32(b, 0);
                WriteU16(b, 0x0002); // 显式宽度标志
                WriteRecord(ms, BrtColInfo, b.ToArray());
            }
        }

        WriteRecord(ms, BrtEndSheet, Empty());
        return ms.ToArray();
    }

    private static void WriteRowHeader(MemoryStream ms, int rw, double? height)
    {
        using var b = new MemoryStream();
        WriteU32(b, (uint)rw);
        WriteU32(b, 0);            // ixfe
        WriteU16(b, height is { } h ? (ushort)(h * 20) : (ushort)0);
        b.WriteByte(0);            // top/bot padding
        b.WriteByte(height is null ? (byte)0 : (byte)0x20); // flags: 0x20 = 显式行高
        b.WriteByte(0);            // phonetic
        WriteU32(b, 0);            // ncolspan
        WriteRecord(ms, BrtRowHdr, b.ToArray());
    }

    private static void WriteCell(MemoryStream ms, WorkbookSpec spec, CellSpec cell, int prevCol)
    {
        bool shortCell = cell.Col == prevCol + 1;
        using var b = new MemoryStream();
        if (!shortCell)
            WriteU32(b, (uint)cell.Col);
        int style = cell.Style >= 0 ? cell.Style : 0;
        b.WriteByte((byte)(style & 0xFF));
        b.WriteByte((byte)((style >> 8) & 0xFF));
        b.WriteByte((byte)((style >> 16) & 0xFF));
        b.WriteByte(0); // fPhShow

        if (cell.InlineText)
        {
            WriteRecord(ms, shortCell ? 0x0011 : BrtCellSt, b, cell.Text ?? "");
        }
        else if (cell.Text is not null)
        {
            int idx = spec.SharedStrings.IndexOf(cell.Text);
            if (idx < 0)
            {
                idx = spec.SharedStrings.Count;
                spec.SharedStrings.Add(cell.Text);
            }
            WriteU32(b, (uint)idx);
            WriteRecord(ms, shortCell ? 0x0012 : BrtCellIsst, b, null);
        }
        else if (cell.Number is { } num)
        {
            if (num == (int)num && num is > -1000 and < 1000)
            {
                int rk = ((int)num) << 2 | 0x02;
                WriteU32(b, (uint)rk);
                WriteRecord(ms, shortCell ? 0x000D : BrtCellRk, b, null);
            }
            else
            {
                WriteDouble(b, num);
                WriteRecord(ms, shortCell ? 0x0010 : BrtCellReal, b, null);
            }
        }
        else if (cell.Bool is { } bo)
        {
            b.WriteByte(bo ? (byte)1 : (byte)0);
            WriteRecord(ms, shortCell ? 0x000F : BrtCellBool, b, null);
        }
        else
        {
            WriteRecord(ms, shortCell ? 0x000C : BrtCellBlank, b, null);
        }
    }

    // ── 记录与基础编码 ──

    private static byte[] Empty() => System.Array.Empty<byte>();

    private static void WriteRecord(MemoryStream ms, int rt, byte[] data)
    {
        WriteVarInt(ms, rt);
        WriteVarInt(ms, data.Length);
        ms.Write(data, 0, data.Length);
    }

    private static void WriteRecord(MemoryStream ms, int rt, MemoryStream body, string? wideString)
    {
        // 追加 wide string 后再写记录
        if (wideString is not null) WriteWideString(body, wideString);
        WriteRecord(ms, rt, body.ToArray());
    }

    private static void WriteVarInt(MemoryStream ms, int value)
    {
        uint v = (uint)value;
        while (v >= 0x80)
        {
            ms.WriteByte((byte)((v & 0x7F) | 0x80));
            v >>= 7;
        }
        ms.WriteByte((byte)v);
    }

    private static void WriteU16(MemoryStream ms, ushort v)
    {
        ms.WriteByte((byte)(v & 0xFF));
        ms.WriteByte((byte)((v >> 8) & 0xFF));
    }

    private static void WriteU32(MemoryStream ms, uint v)
    {
        ms.WriteByte((byte)(v & 0xFF));
        ms.WriteByte((byte)((v >> 8) & 0xFF));
        ms.WriteByte((byte)((v >> 16) & 0xFF));
        ms.WriteByte((byte)((v >> 24) & 0xFF));
    }

    private static void WriteWideString(MemoryStream ms, string s)
    {
        WriteU32(ms, (uint)s.Length);
        foreach (var ch in s)
        {
            ms.WriteByte((byte)(ch & 0xFF));
            ms.WriteByte((byte)((ch >> 8) & 0xFF));
        }
    }

    private static void WriteDouble(MemoryStream ms, double v)
    {
        var bytes = System.BitConverter.GetBytes(v);
        ms.Write(bytes, 0, bytes.Length);
    }
}
