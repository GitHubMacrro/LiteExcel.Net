using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 2.5.0 批 3：iconSet 条件格式写出/读回。
/// </summary>
public class IconSetTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"ico_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void IconSet_WritesXml()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("B1", "分数");
            for (int i = 2; i <= 6; i++) ws.SetValue($"B{i}", i * 10);

            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.IconSet,
                Sqref = "B2:B6",
                IconSet = new IconSetInfo { Style = IconSetStyle.FourRating, Percent = true, ShowValue = true },
            });

            wb.SaveAs(file);

            string sheetXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();

            Assert.Contains("type=\"iconSet\"", sheetXml);
            Assert.Contains("iconSet=\"4Rating\"", sheetXml);
            Assert.Contains("percent=\"1\"", sheetXml);
            Assert.Contains("showValue=\"1\"", sheetXml);
            // 4 图标 → 4 个 cfvo 阈值 0/25/50/75
            Assert.Contains("<cfvo type=\"percent\" val=\"0\"/>", sheetXml);
            Assert.Contains("<cfvo type=\"percent\" val=\"25\"/>", sheetXml);
            Assert.Contains("<cfvo type=\"percent\" val=\"50\"/>", sheetXml);
            Assert.Contains("<cfvo type=\"percent\" val=\"75\"/>", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void IconSet_All16Enums_RoundTrip()
    {
        foreach (var style in (IconSetStyle[])Enum.GetValues(typeof(IconSetStyle)))
        {
            var file = GetTempFile();
            try
            {
                var wb = Excel.Create("S");
                var ws = wb.Worksheets[0];
                ws.SetValue("A1", "v"); ws.SetValue("A2", "1");
                ws.ConditionalFormats.Add(new ConditionalFormat
                {
                    Type = ConditionalFormatType.IconSet,
                    Sqref = "A1:A2",
                    IconSet = new IconSetInfo { Style = style },
                });
                wb.SaveAs(file);

                var opened = Excel.Open(file);
                var cf = opened.Worksheets[0].ConditionalFormats.SingleOrDefault();
                Assert.NotNull(cf);
                Assert.Equal(ConditionalFormatType.IconSet, cf!.Type);
                Assert.NotNull(cf.IconSet);
                Assert.Equal(style, cf.IconSet!.Style);
            }
            finally { if (File.Exists(file)) File.Delete(file); }
        }
    }

    [Fact]
    public void IconSet_CustomThresholds_Written()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "v"); ws.SetValue("A2", "1");
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.IconSet,
                Sqref = "A1:A2",
                IconSet = new IconSetInfo { Style = IconSetStyle.ThreeArrows, Thresholds = new[] { 10.0, 50.0 } },
            });
            wb.SaveAs(file);

            string sheetXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();
            Assert.Contains("val=\"10\"", sheetXml);
            Assert.Contains("val=\"50\"", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void IconSet_ReadRealExcelSample()
    {
        var src = Path.Combine("D:\\OneDriverData\\OneDrive - Private\\Visual Studio Project\\CC\\dotnet\\Customwin.Utils.Xlsx\\LiteXlsx\\files\\iconSet", "iconSet.xlsx");
        if (!File.Exists(src)) return;

        var opened = Excel.Open(src);
        var cfs = opened.Worksheets[0].ConditionalFormats.Where(cf => cf.Type == ConditionalFormatType.IconSet).ToList();
        Assert.True(cfs.Count >= 5, $"iconSet rules: {cfs.Count}");
        // 含 4Rating / 3Arrows / 3TrafficLights2；缺省 iconSet 属性那条回 ThreeArrows
        Assert.Contains(cfs, cf => cf.IconSet!.Style == IconSetStyle.FourRating);
        Assert.Contains(cfs, cf => cf.IconSet!.Style == IconSetStyle.ThreeArrows);
        Assert.Contains(cfs, cf => cf.IconSet!.Style == IconSetStyle.ThreeTrafficLights2);
    }
}
