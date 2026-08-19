using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LiteExcel.Internal.Cfb;

/// <summary>
/// ECMA-376 加密 CFB 写器（忠实移植 msoffcrypto-tool 的 ECMA376Encrypted，BSD-3）。
/// 布局：difat | fat | miniFat | directory | miniData | encryptedPackage，
/// 目录/miniFat/miniData 按普通扇区链交错，FAT 扇区由迭代求解。
/// </summary>
internal static class EncryptedCfbWriter
{
    private const int SectorSize = 512;
    private const int MiniSectorSize = 64;
    private const int MiniStreamCutoff = 4096;
    private const int FirstNumDifat = 109;

    private const int EndOfChain = unchecked((int)0xFFFFFFFE);
    private const int FreeSect = unchecked((int)0xFFFFFFFF);
    private const int FatSect = unchecked((int)0xFFFFFFFD);
    private const int DifSect = unchecked((int)0xFFFFFFFC);

    // MS-OFFCRYPTO DefaultContent 固定字节（来自 Herumi/msoffice）
    private static readonly byte[] VersionContent = HexToBytes(
        "3C0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E004400610074006100530070006100630065007300010000000100000001000000");
    private static readonly byte[] PrimaryContent = HexToBytes(
        "58000000010000004C0000007B00460046003900410033004600300033002D0035003600450046002D0034003600310033002D0042004400440035002D003500410034003100430031004400300037003200340036007D004E0000004D006900630072006F0073006F00660074002E0043006F006E007400610069006E00650072002E0045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D00000001000000010000000100000000000000000000000000000004000000");
    private static readonly byte[] DataSpaceMapContent = HexToBytes(
        "08000000010000006800000001000000000000002000000045006E0063007200790070007400650064005000610063006B00610067006500320000005300740072006F006E00670045006E006300720079007000740069006F006E004400610074006100530070006100630065000000");
    private static readonly byte[] StrongEncryptionDataSpaceContent = HexToBytes(
        "0800000001000000320000005300740072006F006E00670045006E006300720079007000740069006F006E005400720061006E00730066006F0072006D000000");

    private sealed class Dir
    {
        public string Name = "";
        public byte Type;
        public byte Color;
        public int Left = -1;
        public int Right = -1;
        public int Child = -1;
        public byte[] Content = Array.Empty<byte>();
        public int Start = -1;
    }

    // 目录顺序（与 msoffcrypto DSPos 一致）
    private const int IEncryptionPackage = 1;
    private const int IDataSpaces = 2;
    private const int IVersion = 3;
    private const int IDataSpaceMap = 4;
    private const int IDataSpaceInfo = 5;
    private const int IStrongEncryptionDataSpace = 6;
    private const int ITransformInfo = 7;
    private const int IStrongEncryptionTransform = 8;
    private const int IPrimary = 9;
    private const int IEncryptionInfo = 10;

    public static byte[] Build(byte[] encryptedPackage, byte[] encryptionInfo)
    {
        long ft = DateTimeToFileTime(DateTime.Now);

        var dirs = new List<Dir>
        {
            new() { Name = "Root Entry", Type = 5, Color = 0, Child = IEncryptionInfo },
            new() { Name = "EncryptedPackage", Type = 2, Color = 0 },
            new() { Name = "\x06DataSpaces", Type = 1, Color = 0, Child = IDataSpaceMap },
            new() { Name = "Version", Type = 2, Color = 1, Content = VersionContent },
            new() { Name = "DataSpaceMap", Type = 2, Color = 1, Left = IVersion, Right = IDataSpaceInfo, Content = DataSpaceMapContent },
            new() { Name = "DataSpaceInfo", Type = 1, Color = 1, Right = ITransformInfo, Child = IStrongEncryptionDataSpace },
            new() { Name = "StrongEncryptionDataSpace", Type = 2, Color = 1, Content = StrongEncryptionDataSpaceContent },
            new() { Name = "TransformInfo", Type = 1, Color = 0, Child = IStrongEncryptionTransform },
            new() { Name = "StrongEncryptionTransform", Type = 1, Color = 1, Child = IPrimary },
            new() { Name = "\x06Primary", Type = 2, Color = 1, Content = PrimaryContent },
            new() { Name = "EncryptionInfo", Type = 2, Color = 1, Left = IDataSpaces, Right = IEncryptionPackage, Content = encryptionInfo },
        };
        dirs[IEncryptionPackage].Content = encryptedPackage;

        // ── 布局求解 ──
        var layout = new Layout(SectorSize);
        layout.DirectoryEntrySectorNum = BlockNum(dirs.Count, SectorSize / 128);

        // EncryptedPackage：< 4096 走 mini stream（符合 CFB 规范，CfbFile 同样判定），否则普通扇区
        bool encPkgInMini = encryptedPackage.Length < MiniStreamCutoff;
        layout.EncryptionPackageSectorNum = encPkgInMini ? 0 : BlockNum(encryptedPackage.Length, SectorSize);

        // mini 流（EncryptedPackage 若小则纳入）
        var miniStreams = dirs.Where(d => d.Type == 2 &&
            (d.Name != "EncryptedPackage" || encPkgInMini)).ToList();
        layout.MiniFatSectors = miniStreams.Select(s => BlockNum(s.Content.Length, MiniSectorSize)).ToList();
        layout.MiniFatNum = layout.MiniFatSectors.Sum();
        layout.MiniFatDataSectorNum = BlockNum(layout.MiniFatNum, SectorSize / MiniSectorSize);
        layout.NumMiniFatSectors = 1;

        // 迭代求 FAT/DIFAT 扇区数
        int entriesPerSector = SectorSize / 4;
        int difatSectors = 0, fatSectors = 0;
        for (int i = 0; i < 10; i++)
        {
            int a = BlockNum(difatSectors + fatSectors + layout.ContentSectorNum, entriesPerSector);
            int b = a <= FirstNumDifat ? 0 : BlockNum(a - FirstNumDifat, entriesPerSector - 1);
            if (b == difatSectors && a == fatSectors)
            {
                layout.FatSectorNum = fatSectors;
                layout.DifatSectorNum = difatSectors;
                break;
            }
            difatSectors = b;
            fatSectors = a;
        }

        // 扇区位置（派生属性自动计算）
        layout.DifatPos = 0;
        layout.TotalSectors = layout.DifatSectorNum + layout.FatSectorNum + layout.ContentSectorNum;

        // 目录条目起始扇区
        int miniPos = 0;
        foreach (var s in miniStreams)
        {
            s.Start = miniPos;
            miniPos += BlockNum(s.Content.Length, MiniSectorSize);
        }
        dirs[IEncryptionPackage].Start = encPkgInMini ? dirs[IEncryptionPackage].Start : layout.EncryptionPackagePos;
        dirs[0].Start = layout.MiniFatDataPos;
        dirs[0].Content = new byte[64 * layout.MiniFatNum];

        // ── 写输出（纯字节数组偏移写入，避免流位置错乱） ──
        var buf = new byte[SectorSize + layout.TotalSectors * SectorSize];

        // 头部
        {
            int hp = 0;
            WriteBytes(buf, ref hp, new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 });
            hp += 16; // CLSID
            WriteU16(buf, ref hp, 0x003E);      // minor
            WriteU16(buf, ref hp, 0x0003);      // major v3
            WriteU16(buf, ref hp, 0xFFFE);      // byte order
            WriteU16(buf, ref hp, 9);           // sector shift
            WriteU16(buf, ref hp, 6);           // mini sector shift
            WriteU16(buf, ref hp, 0); WriteU16(buf, ref hp, 0); WriteU16(buf, ref hp, 0);
            WriteU32(buf, ref hp, 0);                      // num dir sectors (v3)
            WriteU32(buf, ref hp, (uint)layout.FatSectorNum);
            WriteU32(buf, ref hp, (uint)layout.DirectoryEntryPos);
            WriteU32(buf, ref hp, 0);                      // transaction
            WriteU32(buf, ref hp, (uint)MiniStreamCutoff);
            WriteU32(buf, ref hp, (uint)layout.MiniFatPos);
            WriteU32(buf, ref hp, (uint)layout.NumMiniFatSectors);
            WriteU32(buf, ref hp, layout.DifatSectorNum > 0 ? (uint)layout.DifatPos : unchecked((uint)EndOfChain));
            WriteU32(buf, ref hp, (uint)layout.DifatSectorNum);
            // 头部 DIFAT
            for (int i = 0; i < FirstNumDifat; i++)
                WriteU32(buf, ref hp, i < layout.FatSectorNum ? (uint)(layout.FatPos + i) : unchecked((uint)FreeSect));
        }

        // 目录条目
        for (int i = 0; i < dirs.Count; i++)
        {
            int off = SectorSize + layout.DirectoryEntryPos * SectorSize + i * 128;
            WriteDirectoryEntry(buf, off, dirs[i], ft);
        }
        // mini data
        for (int i = 0; i < miniStreams.Count; i++)
        {
            var s = miniStreams[i];
            int off = SectorSize + layout.MiniFatDataPos * SectorSize + s.Start * MiniSectorSize;
            Array.Copy(s.Content, 0, buf, off, s.Content.Length);
        }
        // EncryptedPackage（非 mini 时单独写普通扇区；mini 时已在上面的 mini data 循环写入）
        if (!encPkgInMini)
        {
            Array.Copy(encryptedPackage, 0, buf,
                SectorSize + layout.EncryptionPackagePos * SectorSize, encryptedPackage.Length);
        }

        // FAT 表
        var fat = BuildFat(layout);
        int entriesPerSec = SectorSize / 4;
        for (int fs = 0; fs < layout.FatSectorNum; fs++)
        {
            int off = SectorSize + (layout.FatPos + fs) * SectorSize;
            for (int j = 0; j < entriesPerSec; j++)
            {
                int idx = fs * entriesPerSec + j;
                WriteU32At(buf, off + j * 4, idx < fat.Count ? (uint)fat[idx] : unchecked((uint)FreeSect));
            }
        }
        // DIFAT 扇区
        for (int ds = 0; ds < layout.DifatSectorNum; ds++)
        {
            int off = SectorSize + (layout.DifatPos + ds) * SectorSize;
            for (int j = 0; j < entriesPerSec - 1; j++)
            {
                int idx = FirstNumDifat + ds * (entriesPerSec - 1) + j;
                WriteU32At(buf, off + j * 4, idx < layout.FatSectorNum ? (uint)(layout.FatPos + idx) : unchecked((uint)FreeSect));
            }
            WriteU32At(buf, off + (entriesPerSec - 1) * 4,
                ds == layout.DifatSectorNum - 1 ? unchecked((uint)EndOfChain) : (uint)(layout.DifatPos + ds + 1));
        }
        // MiniFAT
        {
            int off = SectorSize + layout.MiniFatPos * SectorSize;
            int idx = 0;
            for (int i = 0; i < miniStreams.Count; i++)
            {
                int n = BlockNum(miniStreams[i].Content.Length, MiniSectorSize);
                for (int k = 0; k < n; k++)
                {
                    WriteU32At(buf, off + idx * 4, idx == layout.MiniFatNum - 1 ? unchecked((uint)EndOfChain) : (uint)(idx + 1));
                    idx++;
                }
            }
            for (int i = idx; i < entriesPerSec; i++)
                WriteU32At(buf, off + i * 4, unchecked((uint)FreeSect));
        }

        return buf;
    }

    private static List<int> BuildFat(Layout layout)
    {
        var fat = new int[layout.TotalSectors];
        for (int i = 0; i < layout.TotalSectors; i++) fat[i] = FreeSect;
        // FAT 扇区自身
        for (int i = 0; i < layout.FatSectorNum; i++)
            fat[layout.FatPos + i] = FatSect;
        // DIFAT 扇区
        for (int i = 0; i < layout.DifatSectorNum; i++)
            fat[layout.DifatPos + i] = DifSect;
        // miniFAT
        fat[layout.MiniFatPos] = EndOfChain;
        // 目录链
        for (int i = 0; i < layout.DirectoryEntrySectorNum; i++)
            fat[layout.DirectoryEntryPos + i] = i == layout.DirectoryEntrySectorNum - 1
                ? EndOfChain : layout.DirectoryEntryPos + i + 1;
        // mini data 链
        for (int i = 0; i < layout.MiniFatDataSectorNum; i++)
            fat[layout.MiniFatDataPos + i] = i == layout.MiniFatDataSectorNum - 1
                ? EndOfChain : layout.MiniFatDataPos + i + 1;
        // EncryptedPackage 普通扇区链
        for (int i = 0; i < layout.EncryptionPackageSectorNum; i++)
            fat[layout.EncryptionPackagePos + i] = i == layout.EncryptionPackageSectorNum - 1
                ? EndOfChain : layout.EncryptionPackagePos + i + 1;
        return fat.ToList();
    }

    private static void WriteDirectoryEntry(byte[] buf, int off, Dir d, long ft)
    {
        var nameBytes = Encoding.Unicode.GetBytes(d.Name);
        Array.Copy(nameBytes, 0, buf, off, nameBytes.Length);
        int p = off + 64;
        WriteU16(buf, ref p, nameBytes.Length + 2);
        buf[p++] = d.Type;
        buf[p++] = d.Color;
        WriteS32(buf, ref p, d.Left);
        WriteS32(buf, ref p, d.Right);
        WriteS32(buf, ref p, d.Child);
        p += 16; // CLSID
        WriteU32(buf, ref p, 0); // state bits
        WriteU32(buf, ref p, (uint)(ft & 0xFFFFFFFF));
        WriteU32(buf, ref p, (uint)(ft >> 32));
        WriteU32(buf, ref p, (uint)(ft & 0xFFFFFFFF));
        WriteU32(buf, ref p, (uint)(ft >> 32));
        WriteS32(buf, ref p, d.Start);
        WriteU64(buf, ref p, (ulong)d.Content.Length);
    }

    private static void WriteS32(byte[] d, ref int p, int v) { d[p++] = (byte)v; d[p++] = (byte)(v >> 8); d[p++] = (byte)(v >> 16); d[p++] = (byte)(v >> 24); }
    private static void WriteU64(byte[] d, ref int p, ulong v) { WriteU32(d, ref p, (uint)v); WriteU32(d, ref p, (uint)(v >> 32)); }

    private sealed class Layout
    {
        public Layout(int sectorSize) { }
        public int DifatPos;
        public int FatSectorNum;
        public int DifatSectorNum;
        public int NumMiniFatSectors;
        public int MiniFatNum;
        public int MiniFatDataSectorNum;
        public int DirectoryEntrySectorNum;
        public int EncryptionPackageSectorNum;
        public List<int> MiniFatSectors = new();

        public int FatPos => DifatPos + DifatSectorNum;
        public int MiniFatPos => FatPos + FatSectorNum;
        public int DirectoryEntryPos => MiniFatPos + NumMiniFatSectors;
        public int MiniFatDataPos => DirectoryEntryPos + DirectoryEntrySectorNum;
        public int EncryptionPackagePos => MiniFatDataPos + MiniFatDataSectorNum;
        public int ContentSectorNum => NumMiniFatSectors + DirectoryEntrySectorNum + MiniFatDataSectorNum + EncryptionPackageSectorNum;
        public int TotalSectors;
    }

    private static int BlockNum(int x, int block) => (x + block - 1) / block;

    private static long DateTimeToFileTime(DateTime dt)
    {
        var epoch = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        return (long)(dt.ToUniversalTime() - epoch).TotalSeconds * 10000000;
    }

    private static byte[] HexToBytes(string hex)
    {
        var r = new byte[hex.Length / 2];
        for (int i = 0; i < r.Length; i++) r[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return r;
    }

    private static void WriteBytes(byte[] d, ref int p, byte[] v) { Array.Copy(v, 0, d, p, v.Length); p += v.Length; }
    private static void WriteU16(byte[] d, ref int p, int v) { d[p++] = (byte)v; d[p++] = (byte)(v >> 8); }
    private static void WriteU32(byte[] d, ref int p, uint v) { d[p++] = (byte)v; d[p++] = (byte)(v >> 8); d[p++] = (byte)(v >> 16); d[p++] = (byte)(v >> 24); }
    private static void WriteU32At(byte[] d, int off, uint v) { d[off] = (byte)v; d[off + 1] = (byte)(v >> 8); d[off + 2] = (byte)(v >> 16); d[off + 3] = (byte)(v >> 24); }
}
