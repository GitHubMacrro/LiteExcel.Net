using LiteExcel;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 测试辅助：程序化构造最小 .xls（OLE2/CFB 容器 + BIFF8 记录），用于验证 XlsBackend 读取。
/// </summary>
internal static class XlsTestFile
{
    private const int EndOfChain = unchecked((int)0xFFFFFFFE);
    private const int Free = unchecked((int)0xFFFFFFFF);

    public sealed class CellSpec
    {
        public int Row;
        public int Col;
        public CellType Kind;
        public string Text = "";
        public double Number;
        public bool Bool;
        public bool UseSst;
        public int SstIndex;
    }

    public sealed class SheetSpec
    {
        public string Name = "Sheet1";
        public List<CellSpec> Cells = new();
        public List<(int R1, int R2, int C1, int C2)> Merges = new();
        public List<(int Col, double Width)> ColWidths = new();
        public bool FreezeHeader;
    }

    /// <summary>构造完整 .xls 字节（CFB 封装 + BIFF8 工作簿） </summary>
    public static byte[] Build(params SheetSpec[] sheets) => Build(null, sheets);

    /// <summary>构造完整 .xls 字节，并写入共享字符串表（SST），单元格可用 UseSst/SstIndex 引用 </summary>
    public static byte[] Build(string[]? sst, params SheetSpec[] sheets)
    {
        var workbook = BuildWorkbook(sst, sheets);
        return BuildCfb(workbook);
    }

    /// <summary>把手工构造的工作簿字节封装为 CFB 容器 </summary>
    public static byte[] BuildCfbFromWorkbook(byte[] workbookBytes) => BuildCfb(workbookBytes);

    /// <summary>
    /// 构造一个 SST 字符串跨 CONTINUE 记录、且续接段重新声明编码（UTF-16）的工作簿。
    /// 字符串"你好世界"拆为 SST("你好") + CONTINUE(grbit + "世界")。
    /// </summary>
    public static byte[] BuildSstSplitWorkbook()
    {
        var ms = new MemoryStream();
        WriteRecord(ms, 0x0809, Bof(0x0005));

        var head = new List<byte>();
        head.AddRange(BitConverter.GetBytes((uint)1)); // cstTotal
        head.AddRange(BitConverter.GetBytes((uint)1)); // cstUnique
        head.AddRange(BitConverter.GetBytes((ushort)4)); // cch
        head.Add((byte)0x01);                            // grbit highByte
        head.AddRange(Encoding.Unicode.GetBytes("你好"));
        WriteRecord(ms, 0x00FC, head.ToArray());

        var cont = new List<byte>();
        cont.Add((byte)0x01); // 续接段重新声明 grbit（UTF-16）
        cont.AddRange(Encoding.Unicode.GetBytes("世界"));
        WriteRecord(ms, 0x003C, cont.ToArray());

        WriteRecord(ms, 0x00E0, Xf(0));
        WriteRecord(ms, 0x0022, new byte[] { 0, 0 });
        WriteBoundSheet(ms, "S1");
        WriteRecord(ms, 0x000A, Array.Empty<byte>());

        WriteRecord(ms, 0x0809, Bof(0x0010));
        WriteRecord(ms, 0x00FD, LabelSst(0, 0, 0, 0));
        WriteRecord(ms, 0x000A, Array.Empty<byte>());
        return ms.ToArray();
    }

    private static byte[] BuildWorkbook(string[]? sst, SheetSpec[] sheets)
    {
        var ms = new MemoryStream();

        // 全局流
        WriteRecord(ms, 0x0809, Bof(0x0005));
        WriteRecord(ms, 0x0022, new byte[] { 0, 0 });                       // DATEMODE 1900
        WriteRecord(ms, 0x00E0, Xf(0));                                     // XF 0: 常规
        WriteRecord(ms, 0x00E0, Xf(14));                                    // XF 1: 日期(yyyy-MM-dd)
        if (sst is not null)
            WriteRecord(ms, 0x00FC, Sst(sst));
        foreach (var sheet in sheets)
            WriteBoundSheet(ms, sheet.Name);
        WriteRecord(ms, 0x000A, Array.Empty<byte>());                       // 全局 EOF

        // 各工作表子流
        foreach (var sheet in sheets)
            WriteSheet(ms, sheet);

        return ms.ToArray();
    }

    private static void WriteSheet(MemoryStream ms, SheetSpec sheet)
    {
        WriteRecord(ms, 0x0809, Bof(0x0010));

        foreach (var cell in sheet.Cells.OrderBy(c => c.Row).ThenBy(c => c.Col))
        {
            if (cell.UseSst)
            {
                WriteRecord(ms, 0x00FD, LabelSst(cell.Row, cell.Col, 0, cell.SstIndex));
                continue;
            }
            switch (cell.Kind)
            {
                case CellType.Text:
                    WriteRecord(ms, 0x0204, Label(cell.Row, cell.Col, 0, cell.Text));
                    break;
                case CellType.Number:
                    WriteRecord(ms, 0x0203, Number(cell.Row, cell.Col, 0, cell.Number));
                    break;
                case CellType.Date:
                    WriteRecord(ms, 0x0203, Number(cell.Row, cell.Col, 1, cell.Number));
                    break;
                case CellType.Boolean:
                    WriteRecord(ms, 0x0205, BoolErr(cell.Row, cell.Col, cell.Bool));
                    break;
            }
        }

        if (sheet.Merges.Count > 0)
        {
            var data = new List<byte>();
            data.AddRange(BitConverter.GetBytes((ushort)sheet.Merges.Count));
            foreach (var m in sheet.Merges)
            {
                data.AddRange(BitConverter.GetBytes((ushort)m.R1));
                data.AddRange(BitConverter.GetBytes((ushort)m.R2));
                data.AddRange(BitConverter.GetBytes((ushort)m.C1));
                data.AddRange(BitConverter.GetBytes((ushort)m.C2));
            }
            WriteRecord(ms, 0x00E5, data.ToArray());
        }

        if (sheet.ColWidths.Count > 0)
        {
            foreach (var (col, width) in sheet.ColWidths)
            {
                var data = new List<byte>();
                data.AddRange(BitConverter.GetBytes((ushort)col));
                data.AddRange(BitConverter.GetBytes((ushort)col));
                data.AddRange(BitConverter.GetBytes((ushort)(width * 256)));
                data.AddRange(BitConverter.GetBytes((ushort)0)); // ixfe
                data.AddRange(BitConverter.GetBytes((ushort)0)); // grbit
                data.AddRange(BitConverter.GetBytes((ushort)0)); // cch
                WriteRecord(ms, 0x007D, data.ToArray());
            }
        }

        if (sheet.FreezeHeader)
        {
            var pane = new byte[9];
            BitConverter.GetBytes((ushort)0).CopyTo(pane, 0);  // xSplit
            BitConverter.GetBytes((ushort)1).CopyTo(pane, 2);  // ySplit
            BitConverter.GetBytes((ushort)1).CopyTo(pane, 4);  // topRow
            BitConverter.GetBytes((ushort)0).CopyTo(pane, 6);  // leftCol
            pane[8] = 0;                                       // activePane
            WriteRecord(ms, 0x0041, pane);
        }

        WriteRecord(ms, 0x000A, Array.Empty<byte>());
    }

    // ── BIFF 记录构建 ──

    private static void WriteRecord(MemoryStream ms, ushort opcode, byte[] data)
    {
        ms.Write(BitConverter.GetBytes(opcode), 0, 2);
        ms.Write(BitConverter.GetBytes((ushort)data.Length), 0, 2);
        ms.Write(data, 0, data.Length);
    }

    private static byte[] Bof(int docType)
    {
        var b = new byte[16];
        BitConverter.GetBytes((ushort)0x0600).CopyTo(b, 0); // version BIFF8
        BitConverter.GetBytes((ushort)docType).CopyTo(b, 2);
        return b;
    }

    private static byte[] Sst(string[] strings)
    {
        var data = new List<byte>();
        data.AddRange(BitConverter.GetBytes((uint)strings.Length)); // cstTotal
        data.AddRange(BitConverter.GetBytes((uint)strings.Length)); // cstUnique
        foreach (var s in strings)
            data.AddRange(Biff8String(s));
        return data.ToArray();
    }

    private static byte[] Biff8String(string text)
    {
        bool highByte = text.Any(c => c > 0xFF);
        var list = new List<byte>();
        list.AddRange(BitConverter.GetBytes((ushort)text.Length)); // cch
        list.Add((byte)(highByte ? 0x01 : 0x00));                  // grbit
        list.AddRange(highByte ? Encoding.Unicode.GetBytes(text) : Encoding.Latin1.GetBytes(text));
        return list.ToArray();
    }

    private static byte[] LabelSst(int row, int col, int ixfe, int isst)
    {
        var b = new byte[10];
        BitConverter.GetBytes((ushort)row).CopyTo(b, 0);
        BitConverter.GetBytes((ushort)col).CopyTo(b, 2);
        BitConverter.GetBytes((ushort)ixfe).CopyTo(b, 4);
        BitConverter.GetBytes((int)isst).CopyTo(b, 6);
        return b;
    }

    private static byte[] Xf(int ifmt)
    {
        var b = new byte[20];
        BitConverter.GetBytes((ushort)ifmt).CopyTo(b, 2);
        return b;
    }

    private static void WriteBoundSheet(MemoryStream ms, string name)
    {
        // lbPlyPos(4) grbit(2) cch(1) grbit(1) name
        bool highByte = name.Any(c => c > 0xFF);
        var nameData = highByte ? Encoding.Unicode.GetBytes(name) : Encoding.Latin1.GetBytes(name);
        var b = new byte[8 + nameData.Length];
        BitConverter.GetBytes((int)0).CopyTo(b, 0);          // lbPlyPos
        BitConverter.GetBytes((ushort)0).CopyTo(b, 4);       // grbit
        b[6] = (byte)name.Length;                            // cch
        b[7] = (byte)(highByte ? 0x01 : 0x00);               // grbit
        nameData.CopyTo(b, 8);
        WriteRecord(ms, 0x0085, b);
    }

    private static byte[] Label(int row, int col, int ixfe, string text)
    {
        bool highByte = text.Any(c => c > 0xFF);
        var textData = highByte ? Encoding.Unicode.GetBytes(text) : Encoding.Latin1.GetBytes(text);
        var b = new byte[9 + textData.Length];
        BitConverter.GetBytes((ushort)row).CopyTo(b, 0);
        BitConverter.GetBytes((ushort)col).CopyTo(b, 2);
        BitConverter.GetBytes((ushort)ixfe).CopyTo(b, 4);
        BitConverter.GetBytes((ushort)text.Length).CopyTo(b, 6);
        b[8] = (byte)(highByte ? 0x01 : 0x00);
        textData.CopyTo(b, 9);
        return b;
    }

    private static byte[] Number(int row, int col, int ixfe, double value)
    {
        var b = new byte[14];
        BitConverter.GetBytes((ushort)row).CopyTo(b, 0);
        BitConverter.GetBytes((ushort)col).CopyTo(b, 2);
        BitConverter.GetBytes((ushort)ixfe).CopyTo(b, 4);
        BitConverter.GetBytes(value).CopyTo(b, 6);
        return b;
    }

    private static byte[] BoolErr(int row, int col, bool value)
    {
        var b = new byte[8];
        BitConverter.GetBytes((ushort)row).CopyTo(b, 0);
        BitConverter.GetBytes((ushort)col).CopyTo(b, 2);
        BitConverter.GetBytes((ushort)0).CopyTo(b, 4); // ixfe
        b[6] = (byte)(value ? 1 : 0);
        b[7] = 0; // not error
        return b;
    }

    // ── CFB 容器构建 ──

    private const int SectorSize = 512;

    private static byte[] BuildCfb(byte[] workbook)
    {
        // 工作簿流按常规扇区存储：填充到 >= 4096（mini stream 阈值）
        int wbLen = Math.Max(workbook.Length, 4096);
        var wbPadded = new byte[wbLen];
        Array.Copy(workbook, wbPadded, workbook.Length);

        int wbSectors = wbLen / SectorSize;
        int dirSector = 1 + wbSectors;
        int totalSectors = dirSector + 1;
        int entriesPerSector = SectorSize / 4;

        var file = new byte[512 + totalSectors * SectorSize];

        // 头部
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(file, 0);
        BitConverter.GetBytes((ushort)0x003E).CopyTo(file, 24);          // minor
        BitConverter.GetBytes((ushort)0x0003).CopyTo(file, 26);          // major
        BitConverter.GetBytes((ushort)0xFFFE).CopyTo(file, 28);          // byte order
        BitConverter.GetBytes((ushort)0x0009).CopyTo(file, 30);          // sector shift 512
        BitConverter.GetBytes((ushort)0x0006).CopyTo(file, 32);          // mini sector shift 64
        BitConverter.GetBytes((uint)1).CopyTo(file, 44);                 // FAT 扇区数
        BitConverter.GetBytes((int)dirSector).CopyTo(file, 48);          // 首个目录扇区
        BitConverter.GetBytes((uint)0x1000).CopyTo(file, 56);            // mini stream 阈值 4096
        BitConverter.GetBytes(EndOfChain).CopyTo(file, 60);              // 首个 MiniFAT
        BitConverter.GetBytes(EndOfChain).CopyTo(file, 68);              // 首个 DIFAT
        BitConverter.GetBytes((int)0).CopyTo(file, 76);                  // DIFAT[0] = FAT 扇区 0
        for (int i = 1; i < 109; i++)
            BitConverter.GetBytes(Free).CopyTo(file, 76 + i * 4);

        // FAT 扇区（sector 0）
        var fat = new uint[entriesPerSector];
        for (int i = 0; i < fat.Length; i++) fat[i] = 0xFFFFFFFF;
        fat[0] = 0xFFFFFFFD; // FATSECT
        for (int s = 1; s <= wbSectors; s++) fat[s] = (uint)(s + 1);
        fat[wbSectors] = 0xFFFFFFFE; // ENDOFCHAIN
        fat[dirSector] = 0xFFFFFFFE;
        WriteSector(file, fat, 0);

        // 工作簿数据（sector 1..wbSectors）
        for (int s = 0; s < wbSectors; s++)
            Array.Copy(wbPadded, s * SectorSize, file, 512 + (s + 1) * SectorSize, SectorSize);

        // 目录（sector dirSector）
        var dir = new byte[SectorSize];
        WriteDirEntry(dir, 0, "Root Entry", 5, EndOfChain, 0);
        WriteDirEntry(dir, 1, "Workbook", 2, 1, (ulong)workbook.Length);
        Array.Copy(dir, 0, file, 512 + dirSector * SectorSize, SectorSize);

        return file;
    }

    private static void WriteSector(byte[] file, uint[] entries, int sector)
    {
        int off = 512 + sector * SectorSize;
        for (int i = 0; i < entries.Length; i++)
            BitConverter.GetBytes(entries[i]).CopyTo(file, off + i * 4);
    }

    private static void WriteDirEntry(byte[] dir, int index, string name, byte type, int startSector, ulong size)
    {
        int off = index * 128;
        var nameBytes = Encoding.Unicode.GetBytes(name + "\0");
        Array.Copy(nameBytes, 0, dir, off, nameBytes.Length);
        BitConverter.GetBytes((ushort)nameBytes.Length).CopyTo(dir, off + 64);
        dir[off + 66] = type;
        dir[off + 67] = 1; // 黑色
        BitConverter.GetBytes(Free).CopyTo(dir, off + 68);
        BitConverter.GetBytes(Free).CopyTo(dir, off + 72);
        BitConverter.GetBytes(Free).CopyTo(dir, off + 76);
        BitConverter.GetBytes((int)startSector).CopyTo(dir, off + 116);
        BitConverter.GetBytes((long)size).CopyTo(dir, off + 120);
    }
}
