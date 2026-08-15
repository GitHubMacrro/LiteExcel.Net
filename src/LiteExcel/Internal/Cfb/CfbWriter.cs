using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace LiteExcel.Internal.Cfb;

/// <summary>
/// OLE2 / CFB 复合文档写器（最小实现）。
/// 只写常规 FAT 链：Workbook 流若小于 mini stream 阈值（4096）则补齐到 4096，
/// 保证一律按普通扇区链存储，避免实现 mini stream，同时保证所有读取器按普通流处理。
/// </summary>
internal static class CfbWriter
{
    private const int SectorSize = 512;
    private const int MiniStreamCutoff = 4096;
    private const int FreeSect = unchecked((int)0xFFFFFFFF);
    private const int EndOfChain = unchecked((int)0xFFFFFFFE);
    private const int FatSect = unchecked((int)0xFFFFFFFD);

    /// <summary>构建单流 CFB 文件字节（仅需 Workbook 流即可被 Excel 打开） </summary>
    public static byte[] Build(string streamName, byte[] stream)
    {
        // 小于阈值补齐到 4096，保证按常规 FAT 链存储
        byte[] data = stream;
        if (data.Length < MiniStreamCutoff)
        {
            data = new byte[MiniStreamCutoff];
            Array.Copy(stream, data, stream.Length);
        }

        int wSectors = (data.Length + SectorSize - 1) / SectorSize;

        // FAT 扇区数：F 满足 F*(512/4) >= F + 1(目录) + W，迭代求解
        int fatSectors = 1;
        while (fatSectors * (SectorSize / 4) < fatSectors + 1 + wSectors)
            fatSectors++;
        if (fatSectors > 109)
            throw new LiteExcelException("工作簿过大，暂不支持写出 .xls（FAT 扇区超出单 DIFAT 头容量）");

        int dirSector = fatSectors;
        int firstDataSector = fatSectors + 1;
        int totalSectors = fatSectors + 1 + wSectors;

        // FAT 表（按扇区号索引）
        var fat = new int[totalSectors];
        for (int i = 0; i < totalSectors; i++) fat[i] = FreeSect;
        for (int i = 0; i < fatSectors; i++) fat[i] = FatSect; // FAT 扇区自身须标记 FATSECT
        fat[dirSector] = EndOfChain;
        for (int i = 0; i < wSectors; i++)
            fat[firstDataSector + i] = i == wSectors - 1 ? EndOfChain : firstDataSector + i + 1;

        var header = BuildHeader(fatSectors, dirSector);
        var directory = BuildDirectory(firstDataSector, data.Length);

        using var ms = new MemoryStream(totalSectors * SectorSize);
        ms.Write(header, 0, header.Length);
        int entriesPerSector = SectorSize / 4;
        for (int s = 0; s < fatSectors; s++)
        {
            int start = s * entriesPerSector;
            int count = Math.Min(entriesPerSector, totalSectors - start);
            var sec = new byte[SectorSize];
            for (int i = 0; i < count; i++)
                WriteU32(sec, i * 4, (uint)fat[start + i]);
            for (int i = count; i < entriesPerSector; i++)
                WriteU32(sec, i * 4, unchecked((uint)FreeSect));
            ms.Write(sec, 0, sec.Length);
        }
        ms.Write(directory, 0, directory.Length);
        for (int i = 0; i < wSectors; i++)
        {
            var sec = new byte[SectorSize];
            int off = i * SectorSize;
            Array.Copy(data, off, sec, 0, Math.Min(SectorSize, data.Length - off));
            ms.Write(sec, 0, sec.Length);
        }
        return ms.ToArray();
    }

    private static byte[] BuildHeader(int fatSectors, int dirSector)
    {
        var h = new byte[SectorSize];
        WriteBytes(h, 0, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
        WriteU16(h, 24, 0x003E);           // minor version
        WriteU16(h, 26, 0x0003);           // major version (v3, 512 字节扇区)
        WriteU16(h, 28, 0xFFFE);           // byte order
        WriteU16(h, 30, 9);                // sector shift
        WriteU16(h, 32, 6);                // mini sector shift
        WriteU32(h, 40, 0);                // number of directory sectors（v3 恒为 0，按链推导）
        WriteU32(h, 44, (uint)fatSectors); // number of FAT sectors
        WriteU32(h, 48, (uint)dirSector);  // first directory sector
        WriteU32(h, 52, 0);                // transaction signature
        WriteU32(h, 56, MiniStreamCutoff); // mini stream cutoff
        WriteU32(h, 60, unchecked((uint)EndOfChain)); // first miniFAT sector（无 mini stream）
        WriteU32(h, 64, 0);                // number of miniFAT sectors
        WriteU32(h, 68, unchecked((uint)EndOfChain)); // first DIFAT sector（无）
        WriteU32(h, 72, 0);                // number of DIFAT sectors
        for (int i = 0; i < 109; i++)
            WriteU32(h, 76 + i * 4, i < fatSectors ? (uint)i : unchecked((uint)FreeSect));
        return h;
    }

    private static byte[] BuildDirectory(int firstDataSector, long streamSize)
    {
        var dir = new byte[SectorSize];
        // 根目录条目
        WriteUtf16Name(dir, 0, "Root Entry");
        WriteU16(dir, 64, 22);     // name length（含终止符，"Root Entry" = 11 字符）
        dir[66] = 5;               // root storage
        dir[67] = 1;               // black
        WriteS32(dir, 68, -1);     // left
        WriteS32(dir, 72, -1);     // right
        WriteS32(dir, 76, 1);      // child -> Workbook 条目
        WriteS32(dir, 116, EndOfChain); // 无 mini stream
        WriteU64(dir, 120, 0);
        // Workbook 流条目
        int o = 128;
        WriteUtf16Name(dir, o, "Workbook");
        WriteU16(dir, o + 64, 18); // name length（含终止符："Workbook" = 8 字符 = 16 字节 + 2）
        dir[o + 66] = 2;           // stream
        dir[o + 67] = 1;           // black
        WriteS32(dir, o + 68, -1); // left
        WriteS32(dir, o + 72, -1); // right
        WriteS32(dir, o + 76, -1); // child
        WriteS32(dir, o + 116, firstDataSector);
        WriteU64(dir, o + 120, (ulong)streamSize);
        return dir;
    }

    private static void WriteUtf16Name(byte[] d, int offset, string name)
    {
        var bytes = Encoding.Unicode.GetBytes(name);
        Array.Copy(bytes, 0, d, offset, bytes.Length);
    }

    private static void WriteBytes(byte[] d, int offset, byte[] v) => Array.Copy(v, 0, d, offset, v.Length);

    private static void WriteU16(byte[] d, int offset, int v)
    {
        d[offset] = (byte)v;
        d[offset + 1] = (byte)(v >> 8);
    }

    private static void WriteS32(byte[] d, int offset, int v)
    {
        d[offset] = (byte)v;
        d[offset + 1] = (byte)(v >> 8);
        d[offset + 2] = (byte)(v >> 16);
        d[offset + 3] = (byte)(v >> 24);
    }

    private static void WriteU32(byte[] d, int offset, uint v)
    {
        d[offset] = (byte)v;
        d[offset + 1] = (byte)(v >> 8);
        d[offset + 2] = (byte)(v >> 16);
        d[offset + 3] = (byte)(v >> 24);
    }

    private static void WriteU64(byte[] d, int offset, ulong v)
    {
        WriteU32(d, offset, (uint)v);
        WriteU32(d, offset + 4, (uint)(v >> 32));
    }
}
