using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace LiteExcel.Internal.Biff12;

/// <summary>
/// 把 xlsb 的 <c>connections.bin</c> / <c>queryTableN.bin</c>（BIFF12）转码为 OOXML XML，
/// 供 xlsb → xlsx/xlsm 跨格式转换使用（Power Query / Power Pivot 连接层）。
///
/// 记录布局经真实 Excel 样本逐字节标定：
/// <list type="bullet">
/// <item><c>BrtConnection</c>(0x00C9)：<c>refreshedVersion(u8) minRefreshableVersion(u8) reserved(4) flags(u8) reserved(1)
///   width?(u16) type(u16) reserved(2) reserved(4) id(u32) credentials(u8) [description(XLWideString)] name(XLWideString)</c>。</item>
/// <item><c>BrtDbPr</c>(0x00CB)：<c>flags(u32) reserved(u8) conn(XLWideString) cmd(XLWideString)</c>。</item>
/// <item><c>BrtOlapPr</c>(0x00CD)：<c>flags(u8) rowDrillCount(u32)</c>。</item>
/// <item><c>0x083D</c>：<c>reserved(5) name(XLWideString)</c>（x15 连接 id）。</item>
/// <item><c>0x0844</c>：<c>reserved(4) name(XLWideString)</c>（rangePr sourceName）。</item>
/// </list>
/// </summary>
internal static class XlsbConnectionTranscoder
{
    private const int RtConnection = 0x00C9;
    private const int RtUid = 0x0C00;
    private const int RtDbPr = 0x00CB;
    private const int RtOlapPr = 0x00CD;
    private const int RtConnX15Id = 0x083D;
    private const int RtConnRange = 0x0844;

    private const string ConnNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string X15Ns = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

    internal sealed class ConnectionInfo
    {
        public uint RefreshedVersion;
        public uint MinRefreshableVersion;
        public bool KeepAlive;
        public bool SaveData;
        public byte Credentials;
        public int Type;
        public uint Id;
        public string Name = "";
        public string? Description;
        public string? DbConnection;
        public string? DbCommand;
        public uint DbFlags;
        public bool HasOlapPr;
        public byte OlapFlags;
        public uint RowDrillCount;
        public string? X15Id;
        public string? RangeSourceName;
        public Guid Uid;
    }

    public static List<ConnectionInfo> Parse(byte[] data)
    {
        var list = new List<ConnectionInfo>();
        ConnectionInfo? cur = null;
        Guid pendingUid = Guid.Empty;
        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case RtUid:
                    if (d.Length >= 16)
                    {
                        var gb = new byte[16];
                        Array.Copy(d, 0, gb, 0, 16);
                        pendingUid = new Guid(gb);
                    }
                    break;
                case RtConnection:
                    cur = ParseConnection(d);
                    if (cur is not null) { cur.Uid = pendingUid; pendingUid = Guid.Empty; list.Add(cur); }
                    break;
                case RtDbPr:
                    if (cur is not null) ParseDbPr(d, cur);
                    break;
                case RtOlapPr:
                    if (cur is not null && d.Length >= 5)
                    {
                        cur.HasOlapPr = true;
                        cur.OlapFlags = d[0];
                        cur.RowDrillCount = Biff12Records.ReadU32(d, 1);
                    }
                    break;
                case RtConnX15Id:
                    if (cur is not null) cur.X15Id = ReadWideAt(d, 5);
                    break;
                case RtConnRange:
                    if (cur is not null) cur.RangeSourceName = ReadWideAt(d, 4);
                    break;            }
        }
        return list;
    }

    private static ConnectionInfo? ParseConnection(byte[] d)
    {
        if (d.Length < 23) return null;
        var c = new ConnectionInfo
        {
            RefreshedVersion = d[0],
            MinRefreshableVersion = d[1],
            KeepAlive = (d[6] & 0x01) != 0,
            SaveData = (d[6] & 0x40) != 0,
            Type = Biff12Records.ReadU16(d, 10),
            Id = Biff12Records.ReadU32(d, 18),
            Credentials = d[22],
        };
        // 尾部：可选的 description + name（均为 XLWideString）。
        int off = 23;
        var first = ReadWideString(d, ref off);
        if (off >= d.Length)
        {
            c.Name = first; // 仅一个字符串 → name
        }
        else
        {
            c.Description = first;
            c.Name = ReadWideString(d, ref off);
        }
        return c;
    }

    private static void ParseDbPr(byte[] d, ConnectionInfo c)
    {
        if (d.Length < 5) return;
        c.DbFlags = Biff12Records.ReadU32(d, 0);
        int off = 5; // flags(4) + reserved(1)
        c.DbConnection = ReadWideString(d, ref off);
        c.DbCommand = ReadWideString(d, ref off);
    }

    private static string ReadWideString(byte[] d, ref int off)
    {
        if (off + 4 > d.Length) return "";
        uint cch = Biff12Records.ReadU32(d, off);
        off += 4;
        if (cch == 0 || cch > (uint)((d.Length - off) / 2)) return "";
        var s = Encoding.Unicode.GetString(d, off, (int)cch * 2);
        off += (int)cch * 2;
        return s;
    }

    private static string? ReadWideAt(byte[] d, int start)
    {
        if (start + 4 > d.Length) return null;
        int off = start;
        uint cch = Biff12Records.ReadU32(d, off);
        off += 4;
        // 合理的名称长度保护：0 或越界视为“无名称”。
        if (cch == 0 || cch > (uint)((d.Length - off) / 2)) return null;
        return Encoding.Unicode.GetString(d, off, (int)cch * 2);
    }

    /// <summary>16 字节 GUID（LE 前三段）→ <c>{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}</c>。</summary>
    private static string FormatXr16Uid(byte[] d, int off)
    {
        var g = new Guid(
            Biff12Records.ReadU32(d, off), Biff12Records.ReadU16(d, off + 4), Biff12Records.ReadU16(d, off + 6),
            d[off + 8], d[off + 9], d[off + 10], d[off + 11], d[off + 12], d[off + 13], d[off + 14], d[off + 15]);
        return "{" + g.ToString().ToUpperInvariant() + "}";
    }

    /// <summary>生成 connections.xml 全文。</summary>
    public static string ToXml(IReadOnlyList<ConnectionInfo> conns)
    {
        var sb = new StringBuilder(1024);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<connections xmlns=\"{ConnNs}\" xmlns:mc=\"{McNs}\" mc:Ignorable=\"xr16\" " +
                  "xmlns:xr16=\"http://schemas.microsoft.com/office/spreadsheetml/2017/revision16\">");
        foreach (var c in conns)
        {
            sb.Append($"<connection id=\"{c.Id.ToString(CultureInfo.InvariantCulture)}\" name=\"{Esc(c.Name)}\" type=\"{c.Type}\" " +
                      $"refreshedVersion=\"{c.RefreshedVersion}\"");
            if (c.MinRefreshableVersion > 0) sb.Append($" minRefreshableVersion=\"{c.MinRefreshableVersion}\"");
            if (c.Uid != Guid.Empty) sb.Append($" xr16:uid=\"{{{c.Uid.ToString().ToUpperInvariant()}}}\"");
            if (c.KeepAlive) sb.Append(" keepAlive=\"1\"");
            if (!string.IsNullOrEmpty(c.Description)) sb.Append($" description=\"{Esc(c.Description!)}\"");
            if (c.SaveData) sb.Append(" saveData=\"1\"");
            if (c.Credentials != 0) sb.Append(" credentials=\"none\"");
            sb.Append('>');

            bool isDataModel = string.Equals(c.DbCommand, "Model", StringComparison.Ordinal);
            if (c.DbConnection is not null)
            {
                sb.Append($"<dbPr connection=\"{Esc(c.DbConnection)}\" command=\"{Esc(c.DbCommand ?? "")}\"");
                if (isDataModel) sb.Append(" commandType=\"1\"");
                sb.Append("/>");
            }
            if (c.HasOlapPr)
                sb.Append($"<olapPr sendLocale=\"1\" rowDrillCount=\"{c.RowDrillCount}\"/>");

            if (c.X15Id is not null || c.RangeSourceName is not null || isDataModel)
            {
                sb.Append("<extLst><ext uri=\"{DE250136-89BD-433C-8126-D09CA5730AF9}\" ");
                sb.Append($"xmlns:x15=\"{X15Ns}\">");
                if (c.RangeSourceName is not null)
                    sb.Append($"<x15:connection id=\"{Esc(c.X15Id ?? "")}\"><x15:rangePr sourceName=\"{Esc(c.RangeSourceName)}1\"/></x15:connection>");
                else
                    sb.Append($"<x15:connection id=\"{Esc(c.X15Id ?? "")}\" model=\"1\"/>");
                sb.Append("</ext></extLst>");
            }
            sb.Append("</connection>");
        }
        sb.Append("</connections>");
        return sb.ToString();
    }

    private static string Esc(string s) => s
        .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private const int RtBeginQueryTable = 0x01BF;
    private const int RtBeginQueryTableRefresh = 0x01C1;
    private const int RtQueryTableField = 0x01C9;
    private const string QueryTableNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>生成 queryTableN.xml 全文。</summary>
    public static string ToQueryTableXml(byte[] data)
    {
        string name = "";
        uint connectionId = 0;
        ushort autoFormatId = 0;
        uint nextId = 0;
        string uid = "";
        bool unboundColumnsRight = false;
        var fields = new List<(uint Id, uint ColId, string Name)>();

        foreach (var rec in Biff12Records.ReadAll(data))
        {
            var d = rec.Data;
            switch (rec.Rt)
            {
                case 0x0C00:
                    if (d.Length >= 16) uid = FormatXr16Uid(d, 0);
                    break;
                case RtBeginQueryTable:
                    if (d.Length >= 14)
                    {
                        autoFormatId = Biff12Records.ReadU16(d, 4);
                        connectionId = Biff12Records.ReadU32(d, 6);
                        int off = 10;
                        name = ReadWideString(d, ref off);
                    }
                    break;
                case RtBeginQueryTableRefresh:
                    if (d.Length >= 6) nextId = Biff12Records.ReadU32(d, 2);
                    if (d.Length >= 9 && d[8] != 0) unboundColumnsRight = true;
                    break;
                case RtQueryTableField:
                    if (d.Length >= 12)
                    {
                        uint id = Biff12Records.ReadU32(d, 4);
                        uint colId = Biff12Records.ReadU32(d, 8);
                        int off = 12;
                        var fn = ReadWideString(d, ref off);
                        fields.Add((id, colId, fn));
                    }
                    break;
            }
        }

        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<queryTable xmlns=\"{QueryTableNs}\"");
        if (uid.Length > 0)
            sb.Append($" xmlns:mc=\"{McNs}\" mc:Ignorable=\"xr16\" " +
                      "xmlns:xr16=\"http://schemas.microsoft.com/office/spreadsheetml/2017/revision16\"");
        sb.Append($" name=\"{Esc(name)}\" backgroundRefresh=\"0\" " +
                  $"connectionId=\"{connectionId}\"");
        if (uid.Length > 0) sb.Append($" xr16:uid=\"{uid}\"");
        sb.Append($" autoFormatId=\"{autoFormatId}\" " +
                  "applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" " +
                  "applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"0\">");
        sb.Append($"<queryTableRefresh nextId=\"{nextId}\"");
        if (unboundColumnsRight) sb.Append(" unboundColumnsRight=\"1\"");
        sb.Append($"><queryTableFields count=\"{fields.Count}\">");
        foreach (var (id, colId, fn) in fields)
        {
            if (fn.Length == 0)
                sb.Append($"<queryTableField id=\"{id}\" dataBound=\"0\" tableColumnId=\"{colId}\"/>");
            else
                sb.Append($"<queryTableField id=\"{id}\" name=\"{Esc(fn)}\" tableColumnId=\"{colId}\"/>");
        }
        sb.Append("</queryTableFields></queryTableRefresh></queryTable>");
        return sb.ToString();
    }
}
