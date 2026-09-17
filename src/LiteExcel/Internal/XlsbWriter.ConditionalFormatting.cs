using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using LiteExcel.Internal.Biff;

namespace LiteExcel.Internal;

/// <summary>
/// xlsb 条件格式（conditionalFormatting）写出。
/// 记录层级（[MS-XLSB]，经真实 Excel 样本逐字节标定）：
///   BrtBeginCF(0x01CD) → BrtCFRule(0x01CF) → [类型子记录] → BrtEndCFRule(0x01D0) → BrtEndCF(0x01CE)
/// 类型子记录：
///   iconSet    = BrtBeginIconSet(0x01D1) + BrtCFVO(0x01D7)×N + BrtEndIconSet(0x01D2)
///   colorScale = BrtBeginColorScale(0x01D5) + BrtCFVO×N + BrtColor(0x0234)×N + BrtEndColorScale(0x01D6)
///   dataBar    = BrtBeginDataBar(0x01D3) + BrtCFVO×2 + BrtColor(0x0234) + BrtEndDataBar(0x01D4)
/// DXF（条件格式样式）落在 styles.bin 的 BrtDXF(0x01FB)。
/// 支持全部 18 种 OOXML 规则类型：cellIs / expression / colorScale / dataBar / iconSet / top10 /
///   uniqueValues / duplicateValues / containsText / notContainsText / beginsWith / endsWith /
///   containsBlanks / notContainsBlanks / containsErrors / notContainsErrors / timePeriod / aboveAverage（含 belowAverage）。
/// 注意：textLength 不是合法的 OOXML cfRule 类型（Excel 拒开），Excel 以 cellIs + LEN(ref) op value 表达，
///   本实现同样如此；模型层 ConditionalFormatType.TextLength 映射到该形式。
/// </summary>
internal static partial class XlsbWriter
{
    /// <summary>该条件格式类型是否支持 xlsb 写出。</summary>
    internal static bool IsCfTypeSupported(ConditionalFormatType t) => t switch
    {
        ConditionalFormatType.CellIs => true,
        ConditionalFormatType.Expression => true,
        ConditionalFormatType.ColorScale => true,
        ConditionalFormatType.DataBar => true,
        ConditionalFormatType.IconSet => true,
        ConditionalFormatType.Top10 => true,
        ConditionalFormatType.AboveAverage => true,
        ConditionalFormatType.BelowAverage => true,
        ConditionalFormatType.ContainsText => true,
        ConditionalFormatType.BeginsWith => true,
        ConditionalFormatType.EndsWith => true,
        ConditionalFormatType.NotContainsText => true,
        ConditionalFormatType.Unique => true,
        ConditionalFormatType.Duplicate => true,
        ConditionalFormatType.Blanks => true,
        ConditionalFormatType.NoBlanks => true,
        ConditionalFormatType.Errors => true,
        ConditionalFormatType.NoErrors => true,
        ConditionalFormatType.TextLength => true,
        ConditionalFormatType.TimePeriod => true,
        _ => false,
    };

    /// <summary>把工作表的条件格式写出到 sheet bin。getDxfId 由调用方提供全局 dxf 去重（与 styles.bin 顺序一致）。</summary>
    private static void WriteConditionalFormats(MemoryStream ms, SheetData sheet, Func<CellStyle, int> getDxfId)
    {
        if (sheet.ConditionalFormats is not { Count: > 0 }) return;

        int autoPriority = 1;
        foreach (var cf in sheet.ConditionalFormats)
        {
            if (string.IsNullOrEmpty(cf.Sqref)) continue;
            if (!IsCfTypeSupported(cf.Type)) continue;
            var ranges = ParseSqrefRanges(cf.Sqref);
            if (ranges.Count == 0) continue;

            int dxfId = cf.Style is not null ? getDxfId(cf.Style) : -1;
            int priority = cf.Priority > 0 ? cf.Priority : autoPriority++;

            // 公式引用锚点 = sqref 左上角（首个区域）
            var (anchorRf, _, anchorCf, _) = ranges[0];

            // BrtBeginCF 容器
            var head = new MemoryStream();
            WriteU32(head, 1);                  // cCF（本块规则数；逐规则单块写出）
            WriteU32(head, 0);                  // reserved
            WriteU32(head, (uint)ranges.Count); // cSqref
            foreach (var (rf, rl, cFirst, cLast) in ranges)
            {
                WriteS32(head, rf); WriteS32(head, rl);
                WriteS32(head, cFirst); WriteS32(head, cLast);
            }
            WriteRecord(ms, BrtBeginCF, head.ToArray());

            // BrtCFRule
            WriteRecord(ms, BrtCFRule, BuildCFRule(cf, dxfId, priority, anchorRf, anchorCf));

            // 类型子记录
            switch (cf.Type)
            {
                case ConditionalFormatType.IconSet:
                    WriteIconSetRecords(ms, cf);
                    break;
                case ConditionalFormatType.ColorScale:
                    WriteColorScaleRecords(ms, cf);
                    break;
                case ConditionalFormatType.DataBar:
                    WriteDataBarRecords(ms, cf);
                    break;
            }

            WriteRecord(ms, BrtEndCFRule, Array.Empty<byte>());
            WriteRecord(ms, BrtEndCF, Array.Empty<byte>());
        }
    }

    /// <summary>BrtCFRule(0x01CF) 规则体：46 字节固定头 + 可选文本前缀 + 可选公式段。</summary>
    private static byte[] BuildCFRule(ConditionalFormat cf, int dxfId, int priority, int anchorRow, int anchorCol)
    {
        var (cfType, subType) = MapCfType(cf.Type);
        int param = 0;
        ushort flags = 0;
        string? textPrefix = null;

        switch (cf.Type)
        {
            case ConditionalFormatType.CellIs:
            case ConditionalFormatType.TextLength:
                param = OperatorToXlsb(cf.Operator);
                break;
            case ConditionalFormatType.Top10:
                param = cf.Rank;
                if (cf.Percent) flags |= 0x10;
                break;
            case ConditionalFormatType.AboveAverage:
                flags |= 0x04; // fAboveAverage
                break;
            case ConditionalFormatType.ContainsText:
                param = 0; textPrefix = cf.Text; break;
            case ConditionalFormatType.NotContainsText:
                param = 1; textPrefix = cf.Text; break;
            case ConditionalFormatType.BeginsWith:
                param = 2; textPrefix = cf.Text; break;
            case ConditionalFormatType.EndsWith:
                param = 3; textPrefix = cf.Text; break;
            case ConditionalFormatType.TimePeriod:
            {
                var tp = ResolveTimePeriod(cf.TimePeriod);
                subType = tp.SubType;
                param = tp.Param;
                break;
            }
        }

        // 公式（cellIs 1~2 条；expression/文本类/空值类/时间段 1 条；其余无）。引用相对 sqref 左上角锚点。
        byte[]? rgce1 = null, rgce2 = null;
        int exp1 = 0, exp2 = 0;
        string? formula1 = cf.Formula;
        string refCell = CellRefA1(anchorRow, anchorCol);

        if (cf.Type == ConditionalFormatType.CellIs)
        {
            rgce1 = EncodeCfFormula(formula1, anchorRow, anchorCol, out exp1);
            if (cf.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
                rgce2 = EncodeCfFormula(cf.Formula2, anchorRow, anchorCol, out exp2);
        }
        else if (cf.Type == ConditionalFormatType.TextLength)
        {
            // Excel 以 cellIs + LEN(ref) <op> <value> 表达文本长度（lengthIs 非合法 OOXML 类型）
            var (op1, op2) = TextLengthOperators(cf.Operator);
            rgce1 = EncodeCfFormula($"LEN({refCell}){op1}{formula1}", anchorRow, anchorCol, out exp1);
            if (cf.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
                rgce2 = EncodeCfFormula($"LEN({refCell}){op2}{cf.Formula2}", anchorRow, anchorCol, out exp2);
        }
        else if (cf.Type is ConditionalFormatType.Expression)
        {
            rgce1 = EncodeCfFormula(formula1, anchorRow, anchorCol, out exp1);
        }
        else if (cf.Type == ConditionalFormatType.TimePeriod)
        {
            // 时间段公式为 Excel 固定模板（含前置 PtgAttr），与锚点无关，逐字节标定后原样写出。
            var tp = ResolveTimePeriod(cf.TimePeriod);
            rgce1 = HexToBytes(tp.Rgce);
            exp1 = tp.ExpandedLength;
        }
        else if (textPrefix is not null)
        {
            // 文本类：公式为 Excel 约定（ref 指锚点单元格）
            string f = formula1 ?? cf.Type switch
            {
                ConditionalFormatType.ContainsText => $"NOT(ISERROR(SEARCH(\"{textPrefix}\",{refCell})))",
                ConditionalFormatType.NotContainsText => $"ISERROR(SEARCH(\"{textPrefix}\",{refCell}))",
                ConditionalFormatType.BeginsWith => $"LEFT({refCell},LEN(\"{textPrefix}\"))=\"{textPrefix}\"",
                ConditionalFormatType.EndsWith => $"RIGHT({refCell},LEN(\"{textPrefix}\"))=\"{textPrefix}\"",
                _ => "",
            };
            rgce1 = EncodeCfFormula(f, anchorRow, anchorCol, out exp1);
        }
        else if (cf.Type is ConditionalFormatType.Blanks or ConditionalFormatType.NoBlanks
                 or ConditionalFormatType.Errors or ConditionalFormatType.NoErrors)
        {
            string f = cf.Type switch
            {
                ConditionalFormatType.Blanks => $"LEN(TRIM({refCell}))=0",
                ConditionalFormatType.NoBlanks => $"LEN(TRIM({refCell}))>0",
                ConditionalFormatType.Errors => $"ISERROR({refCell})",
                _ => $"NOT(ISERROR({refCell}))",
            };
            rgce1 = EncodeCfFormula(f, anchorRow, anchorCol, out exp1);
        }

        var ms = new MemoryStream();
        WriteU32(ms, (uint)cfType);
        WriteU32(ms, (uint)subType);
        WriteS32(ms, dxfId);
        WriteU32(ms, (uint)priority);
        WriteU32(ms, (uint)param);
        WriteU32(ms, 0);
        WriteU32(ms, 0);
        WriteU16(ms, flags);
        WriteU16(ms, (ushort)exp1);   // 公式1 展开长度
        WriteU16(ms, 0);
        WriteU16(ms, (ushort)exp2);   // 公式2 展开长度
        WriteU32(ms, 0);
        WriteU16(ms, 0);
        // 文本类：末字段为文本字符数；其余为 -1
        WriteU32(ms, textPrefix is null ? 0xFFFFFFFF : (uint)textPrefix.Length);

        // 文本前缀：原始 UTF-16LE（无长度头），紧随固定头
        if (!string.IsNullOrEmpty(textPrefix))
        {
            var bytes = System.Text.Encoding.Unicode.GetBytes(textPrefix);
            ms.Write(bytes, 0, bytes.Length);
        }
        if (rgce1 is not null)
        {
            WriteU32(ms, (uint)rgce1.Length);
            ms.Write(rgce1, 0, rgce1.Length);
            WriteU32(ms, 0);
        }
        if (rgce2 is not null)
        {
            WriteU32(ms, (uint)rgce2.Length);
            ms.Write(rgce2, 0, rgce2.Length);
            WriteU32(ms, 0);
        }
        return ms.ToArray();
    }

    /// <summary>编码 CF 公式；expandedLength = 长度 + 2×引用数（Excel 标定），引用相对锚点。</summary>
    private static byte[]? EncodeCfFormula(string? formula, int anchorRow, int anchorCol, out int expandedLength)
    {
        expandedLength = 0;
        if (string.IsNullOrEmpty(formula)) return null;
        var rgce = FormulaEncoder.TryEncodeCf(formula, anchorRow, anchorCol, out int refCount);
        if (rgce is null) return null;
        expandedLength = rgce.Length + 2 * refCount;
        return rgce;
    }

    /// <summary>BrtBeginIconSet + BrtCFVO×N + BrtEndIconSet。</summary>
    private static void WriteIconSetRecords(MemoryStream ms, ConditionalFormat cf)
    {
        var iset = cf.IconSet ?? new IconSetInfo();
        int iconCount = iset.IconCount;

        // BrtBeginIconSet: iTemplate(4) + flags(2)。
        // 标定：基值 0x78；bit1(0x02) = !showValue；bit2(0x04) = reverse。
        ushort flags = 0x78;
        if (!iset.ShowValue) flags |= 0x02;
        if (iset.Reverse) flags |= 0x04;
        var head = new MemoryStream();
        WriteU32(head, (uint)IconTemplateId(iset));
        WriteU16(head, flags);
        WriteRecord(ms, BrtBeginIconSet, head.ToArray());

        // BrtCFVO × iconCount（首项 0，其后为阈值）；percent → type=4，num → type=1
        var thresholds = iset.EffectiveThresholds();
        int cfvoType = iset.Percent ? 4 : 1;
        for (int i = 0; i < iconCount; i++)
        {
            double val = i < thresholds.Length ? thresholds[i] : 0;
            WriteRecord(ms, BrtCFVO, BuildCfvo(cfvoType, val, iconSet: true));
        }
        WriteRecord(ms, BrtEndIconSet, Array.Empty<byte>());
    }

    /// <summary>BrtBeginColorScale + BrtCFVO×N + BrtColor×N + BrtEndColorScale。
    /// 2 色 = min/max；3 色 = min/percentile(50)/max（与 xlsx 写出语义一致）。</summary>
    private static void WriteColorScaleRecords(MemoryStream ms, ConditionalFormat cf)
    {
        var cs = cf.ColorScale ?? new ColorScaleInfo();
        bool threeColor = cs.MidColor is not null;

        WriteRecord(ms, BrtBeginColorScale, Array.Empty<byte>());

        // CFVO 类型：num=1, min=2, max=3, percent=4, percentile=5
        if (threeColor)
        {
            WriteRecord(ms, BrtCFVO, BuildCfvo(2, 0, iconSet: false));   // min
            WriteRecord(ms, BrtCFVO, BuildCfvo(4, 50, iconSet: false));  // percent 50
            WriteRecord(ms, BrtCFVO, BuildCfvo(3, 0, iconSet: false));   // max
            WriteRecord(ms, BrtColor, ColorBytes(cs.LowColor));
            WriteRecord(ms, BrtColor, ColorBytes(cs.MidColor!));
            WriteRecord(ms, BrtColor, ColorBytes(cs.HighColor));
        }
        else
        {
            WriteRecord(ms, BrtCFVO, BuildCfvo(2, 0, iconSet: false));   // min
            WriteRecord(ms, BrtCFVO, BuildCfvo(3, 0, iconSet: false));   // max
            WriteRecord(ms, BrtColor, ColorBytes(cs.LowColor));
            WriteRecord(ms, BrtColor, ColorBytes(cs.HighColor));
        }

        WriteRecord(ms, BrtEndColorScale, Array.Empty<byte>());
    }

    /// <summary>BrtBeginDataBar + BrtCFVO(min/max) + BrtColor + BrtEndDataBar。</summary>
    private static void WriteDataBarRecords(MemoryStream ms, ConditionalFormat cf)
    {
        var db = cf.DataBar ?? new DataBarInfo();

        // BrtBeginDataBar: minLength(u8) + maxLength(u8) + flags(u8)；bit0 = showValue
        byte flags = (byte)(db.ShowValue ? 0x01 : 0x00);
        var head = new[] { (byte)db.MinLengthPercent, (byte)db.MaxLengthPercent, flags };
        WriteRecord(ms, BrtBeginDataBar, head);

        WriteRecord(ms, BrtCFVO, BuildCfvo(2, 0, iconSet: false)); // min
        WriteRecord(ms, BrtCFVO, BuildCfvo(3, 0, iconSet: false)); // max
        WriteRecord(ms, BrtColor, ColorBytes(db.Color));

        WriteRecord(ms, BrtEndDataBar, Array.Empty<byte>());
    }

    /// <summary>BrtCFVO(0x01D7)：type(4) + value(f64) + f1(4) + f2(4) + 0(4)。iconSet 时 f1=f2=1，其余为 0。</summary>
    private static byte[] BuildCfvo(int type, double value, bool iconSet)
    {
        var ms = new MemoryStream();
        WriteU32(ms, (uint)type);
        WriteDouble(ms, value);
        WriteU32(ms, iconSet ? 1u : 0u);
        WriteU32(ms, iconSet ? 1u : 0u);
        WriteU32(ms, 0);
        return ms.ToArray();
    }

    /// <summary>解析 sqref（可含空格分隔的多区域，如 "A1:A10 C1:C10"）。</summary>
    private static List<(int rf, int rl, int cFirst, int cLast)> ParseSqrefRanges(string sqref)
    {
        var list = new List<(int, int, int, int)>();
        foreach (var part in sqref.Split(new[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var p = part.Replace("$", "");
            if (string.IsNullOrEmpty(p)) continue;
            var (rf, rl, cFirst, cLast) = ParseRef(p);
            list.Add((rf, rl, cFirst, cLast));
        }
        return list;
    }

    /// <summary>(row, col) 0-based → A1 单元格引用。</summary>
    private static string CellRefA1(int row, int col)
    {
        int c = col;
        var sb = new System.Text.StringBuilder();
        do
        {
            sb.Insert(0, (char)('A' + c % 26));
            c = c / 26 - 1;
        } while (c >= 0);
        return sb.ToString() + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>模型规则类型 → (BrtCFRule.cfType, subType)。TimePeriod 的 subType 由时间段决定，此处用 today 兜底。</summary>
    private static (int cfType, int subType) MapCfType(ConditionalFormatType t) => t switch
    {
        ConditionalFormatType.CellIs => (1, 0),
        ConditionalFormatType.TextLength => (1, 0),   // 以 cellIs + LEN 表达
        ConditionalFormatType.Expression => (2, 1),
        ConditionalFormatType.ColorScale => (3, 2),
        ConditionalFormatType.DataBar => (4, 3),
        ConditionalFormatType.IconSet => (6, 4),
        ConditionalFormatType.Top10 => (5, 5),
        ConditionalFormatType.Unique => (2, 7),
        ConditionalFormatType.Duplicate => (2, 27),
        ConditionalFormatType.ContainsText => (2, 8),
        ConditionalFormatType.NotContainsText => (2, 8),
        ConditionalFormatType.BeginsWith => (2, 8),
        ConditionalFormatType.EndsWith => (2, 8),
        ConditionalFormatType.Blanks => (2, 9),
        ConditionalFormatType.NoBlanks => (2, 10),
        ConditionalFormatType.Errors => (2, 11),
        ConditionalFormatType.NoErrors => (2, 12),
        ConditionalFormatType.TimePeriod => (2, 15),  // 兜底 today
        ConditionalFormatType.AboveAverage => (2, 25),
        ConditionalFormatType.BelowAverage => (2, 26),
        _ => (2, 1),
    };

    /// <summary>cellIs 比较操作 → xlsb operator 码（与 xlFormatConditionOperator 一致）。</summary>
    private static int OperatorToXlsb(ConditionalOperator op) => op switch
    {
        ConditionalOperator.Between => 1,
        ConditionalOperator.NotBetween => 2,
        ConditionalOperator.Equal => 3,
        ConditionalOperator.NotEqual => 4,
        ConditionalOperator.GreaterThan => 5,
        ConditionalOperator.LessThan => 6,
        ConditionalOperator.GreaterThanOrEqual => 7,
        ConditionalOperator.LessThanOrEqual => 8,
        _ => 5,
    };

    /// <summary>textLength 比较操作 → LEN 公式运算符（Excel 原生形式）。</summary>
    private static (string op1, string op2) TextLengthOperators(ConditionalOperator op) => op switch
    {
        ConditionalOperator.Between => (">=", "<="),
        ConditionalOperator.NotBetween => ("<", ">"),
        ConditionalOperator.Equal => ("=", ""),
        ConditionalOperator.NotEqual => ("<>", ""),
        ConditionalOperator.LessThan => ("<", ""),
        ConditionalOperator.LessThanOrEqual => ("<=", ""),
        ConditionalOperator.GreaterThanOrEqual => (">=", ""),
        _ => (">", ""),
    };

    /// <summary>时间段名 → (BrtCFRule.subType, param, 展开长度, rgce 十六进制)。rgce 为 Excel 固定模板，与锚点无关。</summary>
    private static (int SubType, int Param, int ExpandedLength, string Rgce) ResolveTimePeriod(string? period)
    {
        var key = string.IsNullOrEmpty(period) ? "today" : period;
        return TimePeriodMap.TryGetValue(key, out var v) ? v : TimePeriodMap["today"];
    }

    // 时间段 rgce 逐字节来自真实 Excel 样本（c_tp_*.xlsb）。前置 4 字节为 Excel 固定前缀。
    private static readonly Dictionary<string, (int SubType, int Param, int ExpandedLength, string Rgce)> TimePeriodMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["today"] = (15, 0, 23, "1901FEFF4C0000000000C01E0100411D0141DD000B"),
            ["yesterday"] = (17, 1, 27, "1901FEFF4C0000000000C01E0100411D0141DD001E0100040B"),
            ["tomorrow"] = (16, 6, 27, "1901FEFF4C0000000000C01E0100411D0141DD001E0100030B"),
            ["last7Days"] = (18, 2, 50, "1901000041DD004C0000000000C01E0100411D01041E06000A4C0000000000C01E0100411D0141DD000A42022400"),
            ["thisWeek"] = (21, 3, 70, "1901000041DD004C0000000000C01E000041D5000441DD00220146001E0100040A4C0000000000C01E000041D50041DD00041E070041DD0022014600040A42022400"),
            ["lastWeek"] = (23, 4, 68, "1901000041DD004C0000000000C01E000041D5000441DD0022014600150C41DD004C0000000000C01E000041D5000441DD00220146001E070003150942022400"),
            ["nextWeek"] = (22, 7, 72, "1901FCFF4C0000000000C01E000041D50041DD00041E070041DD002201460004150D4C0000000000C01E000041D50041DD00041E0F0041DD002201460004150942022400"),
            ["thisMonth"] = (24, 9, 46, "1901FCFF4C0000000000C041440041DD004144000B4C0000000000C041450041DD004145000B42022400"),
            ["lastMonth"] = (19, 5, 68, "1901FCFF4C0000000000C041440041DD001E00001E0100044202C1014144000B4C0000000000C041450041DD001E00001E0100044202C1014145000B42022400"),
            ["nextMonth"] = (20, 8, 68, "1901FCFF4C0000000000C041440041DD001E00001E0100034202C1014144000B4C0000000000C041450041DD001E00001E0100034202C1014145000B42022400"),
        };

    /// <summary>模型图标集样式 → BIFF12 CfIconSetTemplate 枚举。</summary>
    private static int IconTemplateId(IconSetInfo iset) => iset.Style switch
    {
        IconSetStyle.ThreeArrows => 0,
        IconSetStyle.ThreeArrowsGray => 1,
        IconSetStyle.ThreeFlags => 2,
        IconSetStyle.ThreeTrafficLights => 3,
        IconSetStyle.ThreeTrafficLights2 => 4,
        IconSetStyle.ThreeSigns => 5,
        IconSetStyle.ThreeSymbols => 6,
        IconSetStyle.ThreeSymbols2 => 7,
        IconSetStyle.FourArrows => 8,
        IconSetStyle.FourArrowsGray => 9,
        IconSetStyle.FourRedToBlack => 10,
        IconSetStyle.FourRating => 11,
        IconSetStyle.FourTrafficLights => 12,
        IconSetStyle.FiveArrows => 13,
        IconSetStyle.FiveArrowsGray => 14,
        IconSetStyle.FiveRating => 15,
        _ => 16, // FiveQuarters
    };

    /// <summary>DXF 样式全局去重注册表：保证 CF 规则引用的 dxfId 与 styles.bin 中 BrtDXF 顺序一致。</summary>
    internal sealed class DxfRegistry
    {
        private readonly List<CellStyle> _styles = new();
        private readonly Dictionary<CellStyle, int> _index = new();

        public IReadOnlyList<CellStyle> Styles => _styles;

        public int GetOrCreate(CellStyle style)
        {
            if (_index.TryGetValue(style, out var id)) return id;
            id = _styles.Count;
            _styles.Add(style);
            _index[style] = id;
            return id;
        }
    }

    /// <summary>BrtDXF(0x01FB) 记录体：固定头 + 块计数 + 块序列。</summary>
    private static byte[] BuildDxf(CellStyle style)
    {
        var ms = new MemoryStream();
        WriteU32(ms, 0x00008000); // 固定头（Excel 原生输出）
        var blocks = new MemoryStream();
        int count = 0;

        void Block(ushort type, byte[] data)
        {
            WriteU16(blocks, type);
            WriteU16(blocks, (ushort)(data.Length + 4));
            blocks.Write(data, 0, data.Length);
            count++;
        }

        // 块类型：2=填充色、5=字体色、25=字重、26=下划线、28=斜体、29=删除线；
        // 边框 6/7/8/9 = 上/下/左/右（经真实样本标定）。颜色均为 BrtColor(8)。
        if (!string.IsNullOrEmpty(style.FillColor))
            Block(2, ColorBytes(style.FillColor!));

        if (!string.IsNullOrEmpty(style.FontColor))
            Block(5, ColorBytes(style.FontColor!));

        if (style.Bold || style.Italic || style.Underline || style.Strikeout)
        {
            Block(25, UInt16Bytes(style.Bold ? (ushort)700 : (ushort)400));
            Block(28, new[] { (byte)(style.Italic ? 1 : 0) });
            if (style.Underline) Block(26, new byte[] { 0x01, 0x00 });
            if (style.Strikeout) Block(29, new byte[] { 0x01 });
        }

        if (style.Border is { } border)
        {
            EmitBorder(Block, 6, border.Top);
            EmitBorder(Block, 7, border.Bottom);
            EmitBorder(Block, 8, border.Left);
            EmitBorder(Block, 9, border.Right);
        }

        WriteU16(ms, (ushort)count);
        var blockBytes = blocks.ToArray();
        ms.Write(blockBytes, 0, blockBytes.Length);
        return ms.ToArray();
    }

    private static void EmitBorder(Action<ushort, byte[]> block, ushort type, BorderEdge? edge)
    {
        if (edge is null || string.IsNullOrEmpty(edge.Style)) return;
        var data = new MemoryStream();
        var color = ColorBytes(edge.Color ?? "");
        data.Write(color, 0, color.Length);
        WriteU16(data, (ushort)BorderStyleToXlsb(edge.Style));
        block(type, data.ToArray());
    }

    /// <summary>BrtColor(8)：05 FF 00 00 R G B FF；无颜色 = 全 0（auto）。</summary>
    private static byte[] ColorBytes(string rgb)
    {
        if (string.IsNullOrEmpty(rgb)) return new byte[8];
        var hex = rgb.StartsWith("#") ? rgb.Substring(1) : rgb;
        if (hex.Length != 6) return new byte[8];
        byte r = Convert.ToByte(hex.Substring(0, 2), 16);
        byte g = Convert.ToByte(hex.Substring(2, 2), 16);
        byte b = Convert.ToByte(hex.Substring(4, 2), 16);
        return new byte[] { 0x05, 0xFF, 0x00, 0x00, r, g, b, 0xFF };
    }

    private static byte[] HexToBytes(string hex)
    {
        var bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++)
            bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return bytes;
    }

    private static byte[] UInt16Bytes(ushort v) => new[] { (byte)v, (byte)(v >> 8) };

    private static int BorderStyleToXlsb(string style) => style switch
    {
        "thin" => 1,
        "medium" => 2,
        "dashed" => 3,
        "dotted" => 4,
        "thick" => 5,
        "double" => 6,
        "hair" => 7,
        "mediumDashed" => 8,
        "dashDot" => 9,
        "mediumDashDot" => 10,
        "dashDotDot" => 11,
        "mediumDashDotDot" => 12,
        "slantDashDot" => 13,
        _ => 1,
    };
}
