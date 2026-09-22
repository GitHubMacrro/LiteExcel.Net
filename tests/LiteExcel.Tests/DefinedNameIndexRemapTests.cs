using LiteExcel.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 回归：删除工作表后，存活表的数据验证公式（BrtDVal 0x0040）与单元格公式以 PtgName(0x23)
/// 按「workbook.bin 名表 0 基下标」引用定义名。挂被删表的局部名被移除后名表缩短，
/// 这些下标必须同步重映射；否则残留越界下标会让 Excel 打开时报
/// 「已修复的记录: /xl/worksheets/sheetN.bin 部分的 公式」。
/// 旧实现只改写 3D 引用令牌，未处理名下标，故删第 1 张表（带走多个局部名）时必坏。
/// </summary>
public class DefinedNameIndexRemapTests
{
    private static byte[] VarInt(int v)
    {
        var ms = new MemoryStream();
        while (v >= 0x80) { ms.WriteByte((byte)((v & 0x7F) | 0x80)); v >>= 7; }
        ms.WriteByte((byte)v);
        return ms.ToArray();
    }

    private static byte[] Record(int rt, byte[] data)
    {
        using var ms = new MemoryStream();
        var r = VarInt(rt); ms.Write(r, 0, r.Length);
        var l = VarInt(data.Length); ms.Write(l, 0, l.Length);
        ms.Write(data, 0, data.Length);
        return ms.ToArray();
    }

    private static byte[] U32(uint v) => new[] { (byte)v, (byte)(v >> 8), (byte)(v >> 16), (byte)(v >> 24) };

    /// <summary>构造含一条 BrtDVal 的 sheet bin；rgce1 = PtgName(idx)，rgce2 空。</summary>
    private static byte[] SheetWithDValName(uint nameIndex)
    {
        using var body = new MemoryStream();
        body.Write(U32(0));         // flags
        body.Write(U32(1));         // cRefs
        for (int i = 0; i < 4; i++) body.Write(U32(0)); // rwF/rwL/colF/colL
        for (int i = 0; i < 4; i++) body.Write(U32(0xFFFFFFFF)); // 4 个可空宽串 = null
        body.Write(U32(5));         // cce1
        body.WriteByte(0x23);       // PtgName
        body.Write(U32(nameIndex));
        body.Write(U32(0));         // reserved
        body.Write(U32(0));         // cce2 = 0
        body.Write(U32(0));         // reserved
        return Record(0x0040, body.ToArray());
    }

    /// <summary>从 sheet bin 中取出唯一 BrtDVal 的 rgce1 PtgName 下标。</summary>
    private static uint ReadDValNameIndex(byte[] sheetBin)
    {
        int pos = 0;
        while (pos < sheetBin.Length)
        {
            int rt = ReadVarInt(sheetBin, ref pos);
            int len = ReadVarInt(sheetBin, ref pos);
            var d = new byte[len];
            Array.Copy(sheetBin, pos, d, 0, len);
            pos += len;
            if (rt != 0x0040) continue;
            // flags(4)+cRefs(4)+16+4×4(null strings)
            int o = 8 + 16 + 16;
            int cce1 = BitConverter.ToInt32(d, o); o += 4;
            Assert.Equal(5, cce1);
            Assert.Equal(0x23, d[o]);
            return BitConverter.ToUInt32(d, o + 1);
        }
        throw new InvalidOperationException("no BrtDVal");
    }

    private static int ReadVarInt(byte[] b, ref int pos)
    {
        int v = 0, shift = 0;
        while (true)
        {
            byte x = b[pos++];
            v |= (x & 0x7F) << shift;
            if ((x & 0x80) == 0) break;
            shift += 7;
        }
        return v;
    }

    [Fact]
    public void DValNameIndex_Remapped_WhenEarlierNamesRemoved()
    {
        // 名表原有 10 个名；删表移除下标 2 与 5。原下标 8 应重映射为 8-2=6。
        var sheet = SheetWithDValName(8);
        var removed = new HashSet<int> { 2, 5 };
        var result = XlsbWriter.RewriteSheetRefsToDeleted(sheet, new HashSet<int>(), removed);
        Assert.Equal(6u, ReadDValNameIndex(result));
    }

    [Fact]
    public void DValNameIndex_Unchanged_WhenNoNamesRemoved()
    {
        var sheet = SheetWithDValName(7);
        var result = XlsbWriter.RewriteSheetRefsToDeleted(sheet, new HashSet<int>(), new HashSet<int>());
        Assert.Equal(7u, ReadDValNameIndex(result));
    }

    [Fact]
    public void DValNameIndex_Remapped_OnlyByEarlierRemovals()
    {
        // 移除下标 9（在引用下标 8 之后）不应影响下标 8。
        var sheet = SheetWithDValName(8);
        var result = XlsbWriter.RewriteSheetRefsToDeleted(sheet, new HashSet<int>(), new HashSet<int> { 9 });
        Assert.Equal(8u, ReadDValNameIndex(result));
    }

    [Fact]
    public void DValNameIndex_RemapsMultiple_ToExact()
    {
        // 移除下标 0,1,2（均 < 5）→ 原下标 5 变为 2。
        var sheet = SheetWithDValName(5);
        var result = XlsbWriter.RewriteSheetRefsToDeleted(sheet, new HashSet<int>(), new HashSet<int> { 0, 1, 2 });
        Assert.Equal(2u, ReadDValNameIndex(result));
    }
}
