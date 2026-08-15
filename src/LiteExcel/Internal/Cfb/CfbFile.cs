using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace LiteExcel.Internal.Cfb;

/// <summary>
/// OLE2 / CFB（Compound File Binary）容器只读解析器。
/// 用于读取传统 .xls 文件，仅实现读取所需的最小集合（目录、FAT、MiniFAT、流提取）。
/// </summary>
internal sealed class CfbFile
{
    private static readonly byte[] SignatureBytes =
        { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    private readonly byte[] _data;
    private readonly int _sectorSize;
    private readonly int _miniSectorSize;
    private readonly uint[] _fat;
    private readonly DirectoryEntry[] _dirEntries;

    private CfbFile(byte[] data, int sectorSize, int miniSectorSize,
        uint[] fat, DirectoryEntry[] dirEntries)
    {
        _data = data;
        _sectorSize = sectorSize;
        _miniSectorSize = miniSectorSize;
        _fat = fat;
        _dirEntries = dirEntries;
    }

    /// <summary>打开 CFB 文件。文件不是有效 CFB 时抛 <see cref="LiteExcelException"/> </summary>
    public static CfbFile Open(Stream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        var data = ms.ToArray();
        if (data.Length < 512 || !IsCfb(data))
            throw new LiteExcelException("这不是有效的 .xls 文件（OLE2 复合文档签名缺失）");

        int sectorShift = data[30] | (data[31] << 8);
        int miniSectorShift = data[32] | (data[33] << 8);
        if (sectorShift <= 0 || miniSectorShift <= 0 || sectorShift >= 24 || miniSectorShift >= 12)
            throw new LiteExcelException("这不是有效的 .xls 文件（扇区参数非法）");

        int sectorSize = 1 << sectorShift;
        int miniSectorSize = 1 << miniSectorShift;

        // 读取 DIFAT（前 109 个 FAT 扇区在头部，其余在 DIFAT 扇区链）
        var fatSectorList = ReadDifatSectors(data, sectorSize);
        var fat = ReadFat(data, sectorSize, fatSectorList);

        int firstDirSector = ReadInt32(data, 48);
        // v3 不记录目录扇区数，直接沿 FAT 链走到底
        var dirBytes = ReadChainBytesWalk(data, sectorSize, fat, firstDirSector);
        var dirEntries = ParseDirectory(dirBytes);

        return new CfbFile(data, sectorSize, miniSectorSize, fat, dirEntries);
    }

    /// <summary>按名称提取流字节；不存在返回 null </summary>
    public byte[]? GetStream(string name)
    {
        foreach (var entry in _dirEntries)
        {
            if (entry.ObjectType != 2) continue; // stream
            if (string.Equals(entry.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return ReadDirectoryStream(entry);
            }
        }
        return null;
    }

    private byte[] ReadDirectoryStream(DirectoryEntry entry)
    {
        // 小于 mini stream 阈值（4096）的流放在 mini stream 内
        if (entry.StreamSize < 4096)
        {
            ulong offset = (ulong)(uint)entry.StartSector * (uint)_miniSectorSize;
            var mini = ReadMiniStreamBytes();
            if (offset + entry.StreamSize <= (ulong)mini.Length)
            {
                var buf = new byte[(int)entry.StreamSize];
                Array.Copy(mini, (int)offset, buf, 0, buf.Length);
                return buf;
            }
        }

        return ReadStreamBytesFromChain(_data, _sectorSize, _fat, entry.StartSector, entry.StreamSize);
    }

    private byte[] ReadMiniStreamBytes()
    {
        // mini stream 自身以普通扇区链存储（从根目录条目的 StartSector 开始）
        if (_dirEntries.Length == 0) return Array.Empty<byte>();
        var root = _dirEntries[0];
        return ReadStreamBytesFromChain(_data, _sectorSize, _fat, root.StartSector, root.StreamSize);
    }

    // ── 基础读取 ──

    private static bool IsCfb(byte[] data)
    {
        for (int i = 0; i < 8; i++)
        {
            if (data[i] != SignatureBytes[i]) return false;
        }
        return true;
    }

    private static int ReadInt32(byte[] d, int offset) =>
        d[offset] | (d[offset + 1] << 8) | (d[offset + 2] << 16) | (d[offset + 3] << 24);

    private static uint ReadUInt32(byte[] d, int offset) =>
        (uint)(d[offset] | (d[offset + 1] << 8) | (d[offset + 2] << 16) | (d[offset + 3] << 24));

    private static int SectorOffset(int sector, int sectorSize) => 512 + sector * sectorSize;

    /// <summary>DIFAT：头部 109 个 FAT 扇区 + DIFAT 扇区链中的其余扇区 </summary>
    private static List<int> ReadDifatSectors(byte[] data, int sectorSize)
    {
        var list = new List<int>();
        for (int i = 0; i < 109; i++)
        {
            int s = ReadInt32(data, 76 + i * 4);
            if (s < 0) break;
            list.Add(s);
        }

        // DIFAT 扇区链
        int nextDifat = ReadInt32(data, 68);
        while (nextDifat >= 0)
        {
            int off = SectorOffset(nextDifat, sectorSize);
            int entriesPerSector = sectorSize / 4;
            for (int i = 0; i < entriesPerSector - 1; i++)
            {
                int s = ReadInt32(data, off + i * 4);
                if (s < 0) break;
                list.Add(s);
            }
            nextDifat = ReadInt32(data, off + (entriesPerSector - 1) * 4);
        }
        return list;
    }

    private static uint[] ReadFat(byte[] data, int sectorSize, List<int> fatSectors)
    {
        var entries = new List<uint>();
        foreach (var sector in fatSectors)
        {
            int off = SectorOffset(sector, sectorSize);
            int count = sectorSize / 4;
            for (int i = 0; i < count; i++)
                entries.Add(ReadUInt32(data, off + i * 4));
        }
        return entries.ToArray();
    }

    /// <summary>沿 FAT 链读取字节 </summary>
    private static byte[] ReadStreamBytesFromChain(byte[] data, int sectorSize, uint[] fat, int start, ulong size)
    {
        int len = (int)Math.Min(size, int.MaxValue);
        var result = new byte[len];
        int sector = start;
        int written = 0;
        while (sector >= 0 && written < len)
        {
            int off = SectorOffset(sector, sectorSize);
            int take = Math.Min(sectorSize, len - written);
            Array.Copy(data, off, result, written, take);
            written += take;
            if (sector >= fat.Length) break;
            sector = (int)fat[sector];
        }
        return result;
    }

    /// <summary>沿 FAT 链读取字节直到链结束（用于无长度计数的目录流） </summary>
    private static byte[] ReadChainBytesWalk(byte[] data, int sectorSize, uint[] fat, int start)
    {
        using var ms = new MemoryStream();
        int sector = start;
        int guard = 0;
        while (sector >= 0 && guard++ < 1_000_000)
        {
            int off = SectorOffset(sector, sectorSize);
            int take = Math.Min(sectorSize, data.Length - off);
            ms.Write(data, off, take);
            if (sector >= fat.Length) break;
            sector = (int)fat[sector];
        }
        return ms.ToArray();
    }

    private static DirectoryEntry[] ParseDirectory(byte[] dirBytes)
    {
        var list = new List<DirectoryEntry>();
        int count = dirBytes.Length / 128;
        for (int i = 0; i < count; i++)
        {
            int off = i * 128;
            var entry = new DirectoryEntry
            {
                Name = ReadUtf16Name(dirBytes, off, 64),
                ObjectType = dirBytes[off + 66],
                StartSector = ReadInt32(dirBytes, off + 116),
                StreamSize = ReadUInt64(dirBytes, off + 120),
            };
            if (entry.ObjectType == 0) continue; // unused
            list.Add(entry);
        }
        return list.ToArray();
    }

    private static ulong ReadUInt64(byte[] d, int offset) =>
        (ulong)(uint)ReadInt32(d, offset) | ((ulong)(uint)ReadInt32(d, offset + 4) << 32);

    private static string ReadUtf16Name(byte[] d, int offset, int maxBytes)
    {
        int len = 0;
        while (len + 1 < maxBytes && !(d[offset + len] == 0 && d[offset + len + 1] == 0))
            len += 2;
        if (len == 0) return "";
        return Encoding.Unicode.GetString(d, offset, len);
    }

    private sealed class DirectoryEntry
    {
        public string Name = "";
        public byte ObjectType;
        public int StartSector = -1;
        public ulong StreamSize;
    }
}
