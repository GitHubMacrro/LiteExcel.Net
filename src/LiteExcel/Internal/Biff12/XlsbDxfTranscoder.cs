// XlsbDxfTranscoder：xlsb styles.bin 的 BrtDXF(0x01FB) → xlsx <dxf> XML 转码。
// 记录体布局（真实 Excel 样本标定）：header(u32) + blockCount(u16) + 块序列。
// 每块 = type(u16) + size(u16,=dataLen+4) + data(dataLen)。
// 块类型 → OOXML <dxf> 子元素：
//   0x00(1B u8)=fill patternType(fls)  0x01(8B BrtColor)=fill fgColor  0x02(8B)=fill bgColor
//   0x05(8B BrtColor)=font color  0x06/07/08/09/0A(10B)=border top/bottom/left/right/diagonal
//   0x13(1B u8)=alignment readingOrder  0x14(1B u8)=alignment wrapText
//   0x18(XLWideString)=font name  0x19(2B u16)=font weight(→b)  0x24(4B u32 twips)=font sz
//   0x25(1B)=italic(0/2=否)  0x26(XLWideString)=numFmt formatCode
// 其余块类型（rich dxf 子字段）当前跳过；pivot 引用的 dxf 全部可解码。
// 采用恒等索引（不做去重）：xlsb dxf 索引 = xlsx dxf 索引，<formats> 的 dxfId 直接复用 0x012F 值。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace LiteExcel.Internal.Biff12
{
    internal static class XlsbDxfTranscoder
    {
        private const int BrtDXF = 0x01FB;

        /// <summary>dxf 转码结果：去重后的 &lt;dxf&gt; 列表 + 原始索引→去重索引映射 + 需注册的自定义格式。</summary>
        internal sealed class DxfTranscodeResult
        {
            public readonly List<string> Dxfs = new();
            public readonly Dictionary<int, int> IndexMap = new();   // 原始 xlsb dxf 序号 → 去重后 dxfId
            public readonly Dictionary<int, string> CustomNumFmts = new(); // numFmtId(>=164) → formatCode
        }

        /// <summary>读取 styles.bin 全部 BrtDXF，转码为 &lt;dxf&gt; 并按内容去重（保留首次出现顺序，与 Excel 一致），
        /// 返回去重列表 + 原始序号映射，及 dxf 引用的自定义 numFmt（须注册到 &lt;numFmts&gt;）。</summary>
        public static DxfTranscodeResult TranscodeAll(byte[] stylesBin)
        {
            var result = new DxfTranscodeResult();
            if (stylesBin is null || stylesBin.Length == 0) return result;
            var dedupIndex = new Dictionary<string, int>(StringComparer.Ordinal);
            int orig = 0;
            foreach (var rec in Biff12Records.ReadAll(stylesBin))
            {
                if (rec.Rt != BrtDXF) continue;
                var xml = ToDxfXml(rec.Data, result.CustomNumFmts);
                if (!dedupIndex.TryGetValue(xml, out var id))
                {
                    id = result.Dxfs.Count;
                    dedupIndex[xml] = id;
                    result.Dxfs.Add(xml);
                }
                result.IndexMap[orig] = id;
                orig++;
            }
            return result;
        }

        /// <summary>由内层 dxf 列表合成完整 &lt;dxfs count="N"&gt;...&lt;/dxfs&gt; 字符串；空则返回 null。</summary>
        public static string? BuildDxfsXml(List<string> dxfs)
        {
            if (dxfs is null || dxfs.Count == 0) return null;
            var sb = new StringBuilder(dxfs.Count * 64);
            sb.Append("<dxfs count=\"").Append(dxfs.Count.ToString(CultureInfo.InvariantCulture)).Append("\">");
            foreach (var dxf in dxfs) sb.Append(dxf);
            sb.Append("</dxfs>");
            return sb.ToString();
        }

        private static string ToDxfXml(byte[] body, Dictionary<int, string> customNumFmts)
        {
            // header(u32) + blockCount(u16) + blocks
            if (body.Length < 6) return "<dxf/>";
            int q = 0;
            q += 4; // header（pivot=0、CF=0x8000；不影响 XML）
            int count = Biff12Records.ReadU16(body, q); q += 2;

            // 解析所有块
            int? fls = null; string? fgColor = null, bgColor = null;
            string? fontColor = null; int? weight = null; int? fontSize = null;
            bool italic = false; string? fontName = null; string? numFmt = null; int numFmtId = 0;
            string? borderTop = null, borderBottom = null, borderLeft = null, borderRight = null;
            int? readingOrder = null, wrapText = null;

            for (int i = 0; i < count && q + 4 <= body.Length; i++)
            {
                int type = Biff12Records.ReadU16(body, q); q += 2;
                int size = Biff12Records.ReadU16(body, q); q += 2;
                int dataLen = size - 4;
                if (dataLen < 0 || q + dataLen > body.Length) break;
                int start = q; q += dataLen;

                switch (type)
                {
                    case 0x00: if (dataLen >= 1) fls = body[start]; break;
                    case 0x01: fgColor = ColorAttr(body, start); break;
                    case 0x02: bgColor = ColorAttr(body, start); break;
                    case 0x05: fontColor = ColorAttr(body, start); break;
                    case 0x06: borderTop = BorderEdgeXml(body, start, "top"); break;
                    case 0x07: borderBottom = BorderEdgeXml(body, start, "bottom"); break;
                    case 0x08: borderLeft = BorderEdgeXml(body, start, "left"); break;
                    case 0x09: borderRight = BorderEdgeXml(body, start, "right"); break;
                    case 0x13: if (dataLen >= 1) readingOrder = body[start]; break;
                    case 0x14: if (dataLen >= 1) wrapText = body[start]; break;
                    case 0x18: { int o = start; fontName = ReadDxfString(body, ref o); break; }
                    case 0x19: if (dataLen >= 2) weight = Biff12Records.ReadU16(body, start); break;
                    case 0x24: if (dataLen >= 4) fontSize = (int)Biff12Records.ReadU32(body, start); break;
                    case 0x25: if (dataLen >= 1) italic = body[start] == 1; break;
                    case 0x26: { int o = start; numFmt = ReadDxfString(body, ref o); break; }
                    case 0x29: if (dataLen >= 2) numFmtId = Biff12Records.ReadU16(body, start); break;
                    default: break; // 未识别块跳过（rich dxf 子字段，pivot 不引用）
                }
            }

            // 按 OOXML CT_Dxf 序：font, numFmt, fill, alignment, border（font 必须在 numFmt 之前，否则 Excel 拒开）
            var sb = new StringBuilder(80);
            sb.Append("<dxf>");

            bool anyFont = weight is not null || italic || fontColor is not null || fontName is not null || fontSize is not null;
            if (anyFont)
            {
                sb.Append("<font>");
                if (weight is not null && weight >= 700) sb.Append("<b/>");
                if (italic) sb.Append("<i/>");
                if (fontSize is not null) sb.Append("<sz val=\"").Append((fontSize.Value / 20.0).ToString("R", CultureInfo.InvariantCulture)).Append("\"/>");
                if (fontColor is not null) sb.Append("<color ").Append(fontColor).Append("/>");
                if (fontName is not null) sb.Append("<name val=\"").Append(XmlEsc(fontName)).Append("\"/>");
                sb.Append("</font>");
            }

            if (numFmt is not null)
            {
                if (numFmtId >= 164) customNumFmts[numFmtId] = numFmt;
                sb.Append("<numFmt numFmtId=\"").Append(numFmtId.ToString(CultureInfo.InvariantCulture)).Append("\" formatCode=\"").Append(XmlEsc(numFmt)).Append("\"/>");
            }

            if (fls is not null || fgColor is not null || bgColor is not null)
            {
                sb.Append("<fill><patternFill");
                if (fls is not null) sb.Append(" patternType=\"").Append(PatternName(fls.Value)).Append("\"");
                sb.Append(">");
                if (fgColor is not null) sb.Append("<fgColor ").Append(fgColor).Append("/>");
                if (bgColor is not null) sb.Append("<bgColor ").Append(bgColor).Append("/>");
                sb.Append("</patternFill></fill>");
            }

            if (readingOrder is not null || wrapText is not null)
            {
                sb.Append("<alignment");
                if (wrapText is not null && wrapText.Value != 0) sb.Append(" wrapText=\"1\"");
                else if (wrapText is not null) sb.Append(" wrapText=\"0\"");
                if (readingOrder is not null) sb.Append(" readingOrder=\"").Append(readingOrder.Value.ToString(CultureInfo.InvariantCulture)).Append("\"");
                sb.Append("/>");
            }

            if (borderTop is not null || borderBottom is not null || borderLeft is not null || borderRight is not null)
            {
                sb.Append("<border>");
                sb.Append(borderLeft ?? "<left/>");
                sb.Append(borderRight ?? "<right/>");
                sb.Append(borderTop ?? "<top/>");
                sb.Append(borderBottom ?? "<bottom/>");
                sb.Append("</border>");
            }

            sb.Append("</dxf>");
            return sb.ToString();
        }

        /// <summary>BrtDXF 块内字符串：cch(u16) + UTF-16LE（与 XLWideString 的 u32 前缀不同，真实样本标定）。</summary>
        private static string ReadDxfString(byte[] d, ref int off)
        {
            if (off + 2 > d.Length) return "";
            int cch = Biff12Records.ReadU16(d, off);
            off += 2;
            int bytes = cch * 2;
            if (off + bytes > d.Length) bytes = d.Length - off;
            var s = System.Text.Encoding.Unicode.GetString(d, off, bytes);
            off += bytes;
            return s;
        }

        /// <summary>BrtColor(8B) → 颜色属性串（"rgb=..." / "theme=... [tint=...]" / "indexed=..."）；无颜色返回 null。</summary>
        private static string? ColorAttr(byte[] d, int off)
        {
            if (off + 8 > d.Length) return null;
            byte type = d[off];
            switch (type)
            {
                case 0x05: // RGB：byte4-6 = R G B
                    return "rgb=\"FF" + d[off + 4].ToString("X2", CultureInfo.InvariantCulture)
                        + d[off + 5].ToString("X2", CultureInfo.InvariantCulture)
                        + d[off + 6].ToString("X2", CultureInfo.InvariantCulture) + "\"";
                case 0x07:
                    {
                        int theme = d[off + 1];
                        int tintU16 = Biff12Records.ReadU16(d, off + 2);
                        var sb = new StringBuilder(32);
                        sb.Append("theme=\"").Append(theme.ToString(CultureInfo.InvariantCulture)).Append("\"");
                        if (tintU16 != 0)
                        {
                            double tint = (double)(short)tintU16 / 32767.0;
                            sb.Append(" tint=\"").Append(tint.ToString("R", CultureInfo.InvariantCulture)).Append("\"");
                        }
                        return sb.ToString();
                    }
                case 0x03: // indexed
                    return "indexed=\"" + d[off + 1].ToString(CultureInfo.InvariantCulture) + "\"";
                default:
                    return null; // 0x00=auto / 0x01=无色 → 不输出
            }
        }

        /// <summary>border 块(10B：BrtColor 8 + style u16) → &lt;{edge} style="..."&gt;&lt;color .../&gt;&lt;/{edge}&gt;；无边返回 null。</summary>
        private static string? BorderEdgeXml(byte[] d, int off, string edge)
        {
            if (off + 10 > d.Length) return null;
            int style = Biff12Records.ReadU16(d, off + 8);
            if (style == 0) return null; // 无边
            string name = BorderStyleName(style);
            string? color = ColorAttr(d, off);
            return "<" + edge + " style=\"" + name + "\">" + (color is not null ? "<color " + color + "/>" : "") + "</" + edge + ">";
        }

        private static string PatternName(int fls) => fls switch
        {
            0 => "none",
            1 => "solid",
            2 => "mediumGray",
            3 => "darkGray",
            4 => "lightGray",
            5 => "darkHorizontal",
            6 => "darkVertical",
            7 => "darkDown",
            8 => "darkUp",
            9 => "darkGrid",
            10 => "darkTrellis",
            11 => "lightHorizontal",
            12 => "lightVertical",
            13 => "lightDown",
            14 => "lightUp",
            15 => "lightGrid",
            16 => "lightTrellis",
            17 => "gray125",
            18 => "gray0625",
            _ => "solid",
        };

        private static string BorderStyleName(int style) => style switch
        {
            1 => "thin", 2 => "medium", 3 => "dashed", 4 => "dotted", 5 => "thick",
            6 => "double", 7 => "hair", 8 => "mediumDashed", 9 => "dashDot",
            10 => "mediumDashDot", 11 => "dashDotDot", 12 => "mediumDashDotDot", 13 => "slantDashDot",
            _ => "thin",
        };

        private static string XmlEsc(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var sb = new StringBuilder(s.Length);
            foreach (var c in s)
            {
                switch (c)
                {
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '&': sb.Append("&amp;"); break;
                    case '"': sb.Append("&quot;"); break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }
    }
}
