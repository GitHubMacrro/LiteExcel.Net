using System.Collections.Generic;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// BIFF12（.xlsb 内部二进制格式）记录读取器。
/// 记录头 = RecordType(LEB128 变长，最多 2 字节) + RecordSize(LEB128 变长，最多 4 字节)，无 Instance 字段。
/// 与 [MS-XLSB] 2.2.2 以及 SheetJS/Excel 实测一致（注意：pyxlsb 的 8 位移位解码是错的）。
/// </summary>
internal static class Biff12Records
{
    public readonly struct Record
    {
        public readonly int Rt;
        public readonly byte[] Data;

        public Record(int rt, byte[] data)
        {
            Rt = rt;
            Data = data;
        }
    }

    public static List<Record> ReadAll(byte[] stream)
    {
        var result = new List<Record>();
        int pos = 0;
        int len = stream.Length;
        while (pos < len)
        {
            int rt = ReadVarInt(stream, ref pos);
            int cb = ReadVarInt(stream, ref pos);
            if (cb < 0 || pos + cb > len) break; // 损坏数据，安全截断
            var data = new byte[cb];
            System.Array.Copy(stream, pos, data, 0, cb);
            pos += cb;
            result.Add(new Record(rt, data));
        }
        return result;
    }

    /// <summary>LEB128 变长整数（7 位一组，高位置续） </summary>
    public static int ReadVarInt(byte[] b, ref int pos)
    {
        int v = 0;
        for (int i = 0; i < 4; i++)
        {
            byte x = b[pos++];
            v |= (x & 0x7F) << (7 * i);
            if ((x & 0x80) == 0) return v;
        }
        return v;
    }

    public static ushort ReadU16(byte[] d, int off) =>
        (ushort)(d[off] | (d[off + 1] << 8));

    public static uint ReadU32(byte[] d, int off) =>
        (uint)(d[off] | (d[off + 1] << 8) | (d[off + 2] << 16) | (d[off + 3] << 24));

    public static int ReadS32(byte[] d, int off) =>
        d[off] | (d[off + 1] << 8) | (d[off + 2] << 16) | (d[off + 3] << 24);

    /// <summary>XLWideString：cch(4) + UTF-16LE 字符。返回字符串并推进 offset。</summary>
    public static string ReadWideString(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return "";
        uint cch = ReadU32(d, off);
        off += 4;
        int bytes = (int)cch * 2;
        if (off + bytes > d.Length) return "";
        var s = System.Text.Encoding.Unicode.GetString(d, off, bytes);
        off += bytes;
        return s;
    }
}
