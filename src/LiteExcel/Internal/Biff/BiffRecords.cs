using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// BIFF8 记录读取器：把 Workbook 流切分为 (opcode, data) 记录序列。
/// 仅需记录级访问，不解析字段。
/// </summary>
internal static class BiffRecords
{
    public const ushort OpBof = 0x0809;
    public const ushort OpEof = 0x000A;
    public const ushort OpContinue = 0x003C;
    public const ushort OpBoundSheet = 0x0085;
    public const ushort OpSst = 0x00FC;
    public const ushort OpFormat = 0x041E;
    public const ushort OpXf = 0x00E0;
    public const ushort OpDateMode = 0x0022;
    public const ushort OpFont = 0x0031;
    public const ushort OpRow = 0x0208;
    public const ushort OpNumber = 0x0203;
    public const ushort OpRk = 0x027E;
    public const ushort OpMulRk = 0x00BD;
    public const ushort OpLabelSst = 0x00FD;
    public const ushort OpLabel = 0x0204;
    public const ushort OpBoolErr = 0x0205;
    public const ushort OpFormula = 0x0006;
    public const ushort OpMergedCells = 0x00E5;
    public const ushort OpColInfo = 0x007D;
    public const ushort OpPane = 0x0041;
    public const ushort OpWindow2 = 0x023E;
    public const ushort OpFilePass = 0x002F;
    public const ushort OpHlink = 0x01B8;
    public const ushort OpHlinkTooltip = 0x0800;
    public const ushort OpExternSheet = 0x0017;
    public const ushort OpDefinedName = 0x0018;

    public readonly struct Record
    {
        public readonly ushort Opcode;
        public readonly byte[] Data;

        public Record(ushort opcode, byte[] data)
        {
            Opcode = opcode;
            Data = data;
        }
    }

    public static List<Record> ReadAll(byte[] stream)
    {
        var records = new List<Record>();
        int pos = 0;
        int len = stream.Length;
        while (pos + 4 <= len)
        {
            ushort opcode = ReadU16(stream, pos);
            ushort size = ReadU16(stream, pos + 2);
            pos += 4;
            if (pos + size > len)
            {
                // 截断的最后一个记录：按剩余字节读取，避免抛异常
                size = (ushort)(len - pos);
            }
            var data = new byte[size];
            Array.Copy(stream, pos, data, 0, size);
            records.Add(new Record(opcode, data));
            pos += size;
        }
        return records;
    }

    internal static ushort ReadU16(byte[] d, int offset) =>
        (ushort)(d[offset] | (d[offset + 1] << 8));

    internal static int ReadS32(byte[] d, int offset) =>
        d[offset] | (d[offset + 1] << 8) | (d[offset + 2] << 16) | (d[offset + 3] << 24);

    internal static uint ReadU32(byte[] d, int offset) => (uint)ReadS32(d, offset);
}
