using System;
using System.Collections.Generic;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// BIFF8 Unicode 字符串读取。
/// 处理 SST / 公式字符串 跨 CONTINUE 记录的续接：
/// 当字符数据跨记录边界时，续接记录开头会重新出现一个选项字节（grbit），
/// 本读取器在跨段续接字符时按该规则重新读取编码模式。
/// </summary>
internal sealed class BiffStringReader
{
    private readonly List<byte[]> _segments;
    private int _segIndex;
    private int _pos;

    public BiffStringReader(List<byte[]> segments)
    {
        _segments = segments;
        _segIndex = 0;
        _pos = 0;
    }

    public bool IsEnd => _segIndex >= _segments.Count || _pos >= _segments[_segIndex].Length;

    /// <summary>当前段剩余字节数 </summary>
    public int RemainingInSegment =>
        _segIndex >= _segments.Count ? 0 : _segments[_segIndex].Length - _pos;

    /// <summary>读取一个 BIFF8 字符串（含富文本/拼音扩展）。格式错误时返回 null 并跳过该字符串。 </summary>
    public string? ReadString()
    {
        // 若上一字符串恰好结束在记录边界，前进到下一个续接段再读 cch
        if (RemainingInSegment < 3)
        {
            if (!MoveNextRaw()) return null;
            if (RemainingInSegment < 3) return null;
        }

        int cch = ReadU16();
        // cch 可能跨段（防御性处理）；grbit 紧随其后
        if (RemainingInSegment < 1 && !MoveNextRaw()) return null;
        byte flags = ReadByte();
        bool highByte = (flags & 0x01) != 0;
        bool fExtSt = (flags & 0x02) != 0;   // 富文本
        bool fRichSt = (flags & 0x04) != 0;  // 拼音扩展

        int rtCount = fExtSt ? ReadU16() : 0;
        int extSize = fRichSt ? ReadS32() : 0;

        var sb = new StringBuilder(cch);
        int remaining = cch;

        // 首段字符可能已在开头，跨段时续接段首是新的 grbit
        while (remaining > 0)
        {
            int avail = RemainingInSegment;
            if (avail <= 0)
            {
                if (!MoveToNextSegmentForChars(ref highByte)) return null;
                avail = RemainingInSegment;
            }

            int take = Math.Min(avail / (highByte ? 2 : 1), remaining);
            if (take <= 0)
            {
                // 段内不足以容纳一个字符：跨段，续接段首重新读取 grbit
                if (!MoveToNextSegmentForChars(ref highByte)) return null;
                continue;
            }

            if (highByte)
            {
                sb.Append(Encoding.Unicode.GetString(_segments[_segIndex], _pos, take * 2));
                _pos += take * 2;
            }
            else
            {
                AppendLatin1(sb, _segments[_segIndex], _pos, take);
                _pos += take;
            }
            remaining -= take;
        }

        // 富文本格式运行（每项 4 字节，无 grbit 续接）
        if (rtCount > 0)
        {
            if (!SkipBytes(rtCount * 4)) return null;
        }

        // 拼音扩展数据（无 grbit 续接）
        if (extSize > 0)
        {
            if (!SkipBytes(extSize)) return null;
        }

        return sb.ToString();
    }

    /// <summary>跨段续接字符数据：段首是新的 grbit（0x01 = UTF-16，0x00 = 压缩 Latin1） </summary>
    private bool MoveToNextSegmentForChars(ref bool highByte)
    {
        _segIndex++;
        _pos = 0;
        if (_segIndex >= _segments.Count) return false;
        if (RemainingInSegment < 1) return false;
        byte newFlags = ReadByte();
        highByte = (newFlags & 0x01) != 0;
        return true;
    }

    private bool SkipBytes(int count)
    {
        int remaining = count;
        while (remaining > 0)
        {
            int avail = RemainingInSegment;
            if (avail <= 0)
            {
                _segIndex++;
                _pos = 0;
                if (_segIndex >= _segments.Count) return false;
                continue;
            }
            int take = Math.Min(avail, remaining);
            _pos += take;
            remaining -= take;
        }
        return true;
    }

    private byte ReadByte()
    {
        var b = _segments[_segIndex][_pos];
        _pos++;
        return b;
    }

    private ushort ReadU16()
    {
        int avail = RemainingInSegment;
        if (avail >= 2)
        {
            var d = _segments[_segIndex];
            ushort v = (ushort)(d[_pos] | (d[_pos + 1] << 8));
            _pos += 2;
            return v;
        }

        // 跨段读取（2 字节跨边界时直接拼接）
        if (avail == 1)
        {
            int first = ReadByte();
            if (RemainingInSegment < 1 && !MoveNextRaw()) return 0;
            int second = ReadByte();
            return (ushort)(first | (second << 8));
        }

        if (!MoveNextRaw()) return 0;
        return ReadU16();
    }

    private int ReadS32()
    {
        var bytes = new byte[4];
        for (int i = 0; i < 4; i++)
        {
            if (RemainingInSegment < 1 && !MoveNextRaw()) return 0;
            bytes[i] = ReadByte();
        }
        return bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24);
    }

    private bool MoveNextRaw()
    {
        _segIndex++;
        _pos = 0;
        return _segIndex < _segments.Count;
    }

    private static void AppendLatin1(StringBuilder sb, byte[] data, int offset, int count)
    {
        for (int i = 0; i < count; i++)
            sb.Append((char)data[offset + i]);
    }
}
