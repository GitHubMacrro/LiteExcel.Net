using LiteExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace LiteExcel.Tests;

/// <summary>
/// xlsb 条件格式写入：BrtBeginCF(0x01CD) → BrtCFRule(0x01CF) → [iconSet 子记录] → BrtEndCFRule/BrtEndCF，
/// 以及 styles.bin 的 BrtDXF(0x01FB)。记录布局经真实 Excel 样本逐字节标定。
/// </summary>
public class XlsbConditionalFormatTests
{
    private const int BrtBeginCF = 0x01CD;
    private const int BrtCFRule = 0x01CF;
    private const int BrtEndCFRule = 0x01D0;
    private const int BrtEndCF = 0x01CE;
    private const int BrtBeginIconSet = 0x01D1;
    private const int BrtEndIconSet = 0x01D2;
    private const int BrtCFVO = 0x01D7;
    private const int BrtBeginColorScale = 0x01D5;
    private const int BrtEndColorScale = 0x01D6;
    private const int BrtBeginDataBar = 0x01D3;
    private const int BrtEndDataBar = 0x01D4;
    private const int BrtColor = 0x0234;
    private const int BrtDXF = 0x01FB;

    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"xlsbcf_{Guid.NewGuid():N}.xlsb");

    private static List<(int Rt, byte[] Data)> Records(string file, string entry)
    {
        using var zip = ZipFile.OpenRead(file);
        var e = zip.GetEntry(entry)!;
        using var ms = new MemoryStream();
        using (var s = e.Open()) s.CopyTo(ms);
        var b = ms.ToArray();
        var list = new List<(int, byte[])>();
        int pos = 0;
        while (pos < b.Length)
        {
            int rt = ReadVarInt(b, ref pos);
            int len = ReadVarInt(b, ref pos);
            if (len < 0 || pos + len > b.Length) break;
            var data = new byte[len];
            Array.Copy(b, pos, data, 0, len);
            list.Add((rt, data));
            pos += len;
        }
        return list;
    }

    private static int ReadVarInt(byte[] b, ref int pos)
    {
        int v = 0, sh = 0;
        for (int i = 0; i < 4; i++)
        {
            byte x = b[pos++];
            v |= (x & 0x7F) << sh;
            if ((x & 0x80) == 0) return v;
            sh += 7;
        }
        return v;
    }

    private static uint U32(byte[] d, int off) => BitConverter.ToUInt32(d, off);

    [Fact]
    public void CellIs_WritesCfRecordsAndDxf()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs,
                Sqref = "A1:A10",
                Operator = ConditionalOperator.GreaterThan,
                Formula = "5",
                Style = new CellStyle { FillColor = "#FF0000" },
            });
            wb.SaveAs(file);

            var recs = Records(file, "xl/worksheets/sheet1.bin");
            Assert.Contains(recs, r => r.Rt == BrtBeginCF);
            var rule = recs.First(r => r.Rt == BrtCFRule);
            Assert.Equal(1u, U32(rule.Data, 0));       // cfType = cellIs
            Assert.Equal(0u, U32(rule.Data, 4));       // subType
            Assert.Equal(0u, U32(rule.Data, 8));       // dxfId
            Assert.Equal(1u, U32(rule.Data, 12));      // priority
            Assert.Equal(5u, U32(rule.Data, 16));      // operator = greaterThan
            Assert.Contains(recs, r => r.Rt == BrtEndCFRule);
            Assert.Contains(recs, r => r.Rt == BrtEndCF);

            // DXF 写入 styles.bin
            var srecs = Records(file, "xl/styles.bin");
            var dxf = srecs.First(r => r.Rt == BrtDXF);
            // 固定头 00 80 00 00 + 块计数 1 + 块 type=2（填充）
            Assert.Equal(0x00008000u, U32(dxf.Data, 0));
            Assert.Equal(1, BitConverter.ToUInt16(dxf.Data, 4));
            Assert.Equal(2, BitConverter.ToUInt16(dxf.Data, 6)); // fill block
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Between_WritesTwoFormulas()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs, Sqref = "A1:A10",
                Operator = ConditionalOperator.Between, Formula = "5", Formula2 = "10",
                Style = new CellStyle { FillColor = "#00FF00" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            Assert.Equal(1u, U32(rule.Data, 16));  // between operator
            // 46 字节固定头 + 两条公式段（各 cce(4)+rgce+reserved(4)）
            Assert.True(rule.Data.Length > 46 + 8);
            // 第一条公式 rgce 长度 = 3（PtgInt 5）
            Assert.Equal(3u, U32(rule.Data, 46));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void IconSet_WritesSubRecordsAndTemplate()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.IconSet, Sqref = "A1:A10",
                IconSet = new IconSetInfo { Style = IconSetStyle.FourRating },
            });
            wb.SaveAs(file);

            var recs = Records(file, "xl/worksheets/sheet1.bin");
            var rule = recs.First(r => r.Rt == BrtCFRule);
            Assert.Equal(6u, U32(rule.Data, 0));   // cfType = iconSet
            Assert.Equal(4u, U32(rule.Data, 4));   // subType

            var iconSet = recs.First(r => r.Rt == BrtBeginIconSet);
            Assert.Equal(11u, U32(iconSet.Data, 0)); // 4Rating template
            Assert.Contains(recs, r => r.Rt == BrtEndIconSet);
            // 4 图标 → 4 个 CFVO
            Assert.Equal(4, recs.Count(r => r.Rt == BrtCFVO));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Expression_RefIsAnchorRelative()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("B1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Expression, Sqref = "B1:B10",
                Formula = "MOD(B1,2)=0",
                Style = new CellStyle { FontColor = "#0000FF" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            Assert.Equal(2u, U32(rule.Data, 0));   // cfType = expression
            Assert.Equal(1u, U32(rule.Data, 4));   // subType
            // 公式段起点 46：cce(4) + rgce。rgce 首字节为 0x4C（CF PtgRef）。
            Assert.Equal(0x4C, rule.Data[46 + 4]);
            // B1 相对锚点 B1 的偏移 = (0,0)，col 字段带相对标志 0xC000。
            // rgce: 4C(1) + rw(u32)=0 + col(u16)=0xC000 → col 低字节 0x00、高字节 0xC0。
            Assert.Equal(0x00, rule.Data[46 + 4 + 1 + 4]); // col low
            Assert.Equal(0xC0, rule.Data[46 + 4 + 1 + 5]); // col high
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Top10_WritesRankAndPercent()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Top10, Sqref = "A1:A10", Rank = 20, Percent = true,
                Style = new CellStyle { FillColor = "#FFFF00" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            Assert.Equal(5u, U32(rule.Data, 0));   // cfType = top10
            Assert.Equal(5u, U32(rule.Data, 4));   // subType
            Assert.Equal(20u, U32(rule.Data, 16)); // rank
            // 标志位低 16 位含 percent (0x10)
            Assert.Equal(0x10, BitConverter.ToUInt16(rule.Data, 28));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ColorScale_WritesCfvoAndColors()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.ColorScale, Sqref = "A1:A10",
                ColorScale = new ColorScaleInfo { LowColor = "F8696B", HighColor = "63BE7B" },
            });
            wb.SaveAs(file);

            var recs = Records(file, "xl/worksheets/sheet1.bin");
            var rule = recs.First(r => r.Rt == BrtCFRule);
            Assert.Equal(3u, U32(rule.Data, 0));   // cfType = colorScale
            Assert.Equal(2u, U32(rule.Data, 4));   // subType
            Assert.Contains(recs, r => r.Rt == BrtBeginColorScale);
            Assert.Contains(recs, r => r.Rt == BrtEndColorScale);
            // 2 色 → 2 CFVO（min=2, max=3）+ 2 BrtColor
            var cfvos = recs.Where(r => r.Rt == BrtCFVO).ToList();
            Assert.Equal(2, cfvos.Count);
            Assert.Equal(2u, U32(cfvos[0].Data, 0));
            Assert.Equal(3u, U32(cfvos[1].Data, 0));
            Assert.Equal(2, recs.Count(r => r.Rt == BrtColor));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ColorScale3_WritesPercentileAndThreeColors()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.ColorScale, Sqref = "A1:A10",
                ColorScale = new ColorScaleInfo { LowColor = "FF0000", MidColor = "FFFF00", HighColor = "00FF00" },
            });
            wb.SaveAs(file);

            var recs = Records(file, "xl/worksheets/sheet1.bin");
            var cfvos = recs.Where(r => r.Rt == BrtCFVO).ToList();
            Assert.Equal(3, cfvos.Count);
            Assert.Equal(2u, U32(cfvos[0].Data, 0));  // min
            Assert.Equal(4u, U32(cfvos[1].Data, 0));  // percent
            Assert.Equal(50d, BitConverter.ToDouble(cfvos[1].Data, 4));
            Assert.Equal(3u, U32(cfvos[2].Data, 0));  // max
            Assert.Equal(3, recs.Count(r => r.Rt == BrtColor));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DataBar_WritesHeaderCfvoColor()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.DataBar, Sqref = "A1:A10",
                DataBar = new DataBarInfo { Color = "638EC6", MinLengthPercent = 10, MaxLengthPercent = 90, ShowValue = false },
            });
            wb.SaveAs(file);

            var recs = Records(file, "xl/worksheets/sheet1.bin");
            var rule = recs.First(r => r.Rt == BrtCFRule);
            Assert.Equal(4u, U32(rule.Data, 0));   // cfType = dataBar
            Assert.Equal(3u, U32(rule.Data, 4));   // subType

            var head = recs.First(r => r.Rt == BrtBeginDataBar);
            Assert.Equal(10, head.Data[0]);        // minLength
            Assert.Equal(90, head.Data[1]);        // maxLength
            Assert.Equal(0, head.Data[2]);         // showValue = false
            Assert.Equal(2, recs.Count(r => r.Rt == BrtCFVO));
            Assert.Contains(recs, r => r.Rt == BrtEndDataBar);
            // 颜色 638EC6
            var color = recs.First(r => r.Rt == BrtColor);
            Assert.Equal(0x63, color.Data[4]);
            Assert.Equal(0x8E, color.Data[5]);
            Assert.Equal(0xC6, color.Data[6]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void TextLength_WritesCellIsWithLenFormula()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "hello");
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.TextLength, Sqref = "A1:A10",
                Operator = ConditionalOperator.GreaterThan, Formula = "5",
                Style = new CellStyle { FillColor = "#FF0000" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            // lengthIs 非合法 OOXML → 以 cellIs 写出
            Assert.Equal(1u, U32(rule.Data, 0));   // cfType = cellIs
            Assert.Equal(0u, U32(rule.Data, 4));   // subType
            Assert.Equal(5u, U32(rule.Data, 16));  // operator = greaterThan
            // 公式段起点 46：cce(4) + rgce。rgce 首字节 0x4C（PtgRef 锚点 A1）
            Assert.Equal(0x4C, rule.Data[46 + 4]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Duplicate_WritesSubType27()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Duplicate, Sqref = "A1:A10",
                Style = new CellStyle { FillColor = "#00FF00" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            Assert.Equal(2u, U32(rule.Data, 0));   // cfType = expression 族
            Assert.Equal(27u, U32(rule.Data, 4));  // subType = duplicateValues
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void TimePeriod_WritesSubTypeAndTemplate()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.TimePeriod, Sqref = "A1:A10",
                TimePeriod = "today",
                Style = new CellStyle { FillColor = "#FFFF00" },
            });
            wb.SaveAs(file);

            var rule = Records(file, "xl/worksheets/sheet1.bin").First(r => r.Rt == BrtCFRule);
            Assert.Equal(2u, U32(rule.Data, 0));    // cfType
            Assert.Equal(15u, U32(rule.Data, 4));   // subType = today
            Assert.Equal(0u, U32(rule.Data, 16));   // param = 0
            // 公式段 rgce 首字节 0x19（PtgAttr 前置），展开长度 23
            Assert.Equal(23u, BitConverter.ToUInt16(rule.Data, 30));
            Assert.Equal(0x19, rule.Data[46 + 4]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void UnsupportedType_ReportsDegradation()
    {
        // 全部 OOXML 规则类型均已支持；此处用越界枚举值模拟未知类型。
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = (ConditionalFormatType)999, Sqref = "A1:A10",
            });
            wb.SaveAs(file);

            Assert.Contains(wb.SaveDegradations, d => d.Capability == DegradationCapability.ConditionalFormatting);
            var recs = Records(file, "xl/worksheets/sheet1.bin");
            Assert.DoesNotContain(recs, r => r.Rt == BrtCFRule);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SupportedType_NoDegradation()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs, Sqref = "A1:A10",
                Operator = ConditionalOperator.GreaterThan, Formula = "5",
                Style = new CellStyle { FillColor = "#FF0000" },
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.IconSet, Sqref = "B1:B10",
                IconSet = new IconSetInfo { Style = IconSetStyle.ThreeArrows },
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.ColorScale, Sqref = "C1:C10",
                ColorScale = new ColorScaleInfo(),
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.DataBar, Sqref = "D1:D10",
                DataBar = new DataBarInfo(),
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.TextLength, Sqref = "E1:E10",
                Operator = ConditionalOperator.GreaterThan, Formula = "5",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.TimePeriod, Sqref = "F1:F10",
                TimePeriod = "lastWeek",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Duplicate, Sqref = "G1:G10",
            });
            wb.SaveAs(file);

            Assert.DoesNotContain(wb.SaveDegradations, d => d.Capability == DegradationCapability.ConditionalFormatting);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Dxf_BorderEdges_UseTopBottomLeftRight()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs, Sqref = "A1:A10",
                Operator = ConditionalOperator.GreaterThan, Formula = "5",
                Style = new CellStyle
                {
                    Border = new BorderStyle { Bottom = new BorderEdge { Style = "thin", Color = "#FF0000" } },
                },
            });
            wb.SaveAs(file);

            var dxf = Records(file, "xl/styles.bin").First(r => r.Rt == BrtDXF);
            // 固定头(4) + 块计数(2) + 块 type(u16)
            Assert.Equal(1, BitConverter.ToUInt16(dxf.Data, 4));   // 1 块
            Assert.Equal(7, BitConverter.ToUInt16(dxf.Data, 6));   // 下边框 = type 7
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
