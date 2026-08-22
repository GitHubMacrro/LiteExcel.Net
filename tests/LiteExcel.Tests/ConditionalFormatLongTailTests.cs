using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// A2：条件格式长尾类型一期（文本/时间周期/空值/错误/唯一重复/前N/平均线）。
/// 验证写→读往返 + sheet XML 结构合法性。
/// </summary>
public class ConditionalFormatLongTailTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"a2_{Guid.NewGuid():N}.xlsx");

    private static string ReadSheetXml(string file)
    {
        using var zip = ZipFile.OpenRead(file);
        using var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    [Fact]
    public void TextConditions_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.ContainsText,
                Sqref = "A1:A10",
                Text = "urgent",
                Style = new CellStyle { FontColor = "#FF0000" },
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.BeginsWith,
                Sqref = "B1:B10",
                Text = "ID-",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.EndsWith,
                Sqref = "C1:C10",
                Text = ".xlsx",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.NotContainsText,
                Sqref = "D1:D10",
                Text = "temp",
            });
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("type=\"containsText\"", xml);
            Assert.Contains("text=\"urgent\"", xml);
            Assert.Contains("type=\"beginsWith\"", xml);
            Assert.Contains("type=\"endsWith\"", xml);
            Assert.Contains("type=\"notContainsText\"", xml);

            var rules = Excel.Open(file).Worksheets[0].ConditionalFormats;
            Assert.Equal(4, rules.Count);
            Assert.Equal(ConditionalFormatType.ContainsText, rules[0].Type);
            Assert.Equal("urgent", rules[0].Text);
            Assert.Equal(ConditionalFormatType.BeginsWith, rules[1].Type);
            Assert.Equal("ID-", rules[1].Text);
            Assert.Equal(ConditionalFormatType.EndsWith, rules[2].Type);
            Assert.Equal(ConditionalFormatType.NotContainsText, rules[3].Type);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BlanksErrorsUniqueDuplicate_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.Blanks, Sqref = "A1:A20" });
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.NoBlanks, Sqref = "B1:B20" });
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.Errors, Sqref = "C1:C20" });
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.NoErrors, Sqref = "D1:D20" });
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.Unique, Sqref = "E1:E20" });
            ws.ConditionalFormats.Add(new ConditionalFormat { Type = ConditionalFormatType.Duplicate, Sqref = "F1:F20" });
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("type=\"containsBlanks\"", xml);
            Assert.Contains("type=\"notContainsBlanks\"", xml);
            Assert.Contains("type=\"containsErrors\"", xml);
            Assert.Contains("type=\"notContainsErrors\"", xml);
            Assert.Contains("type=\"uniqueValues\"", xml);
            Assert.Contains("type=\"duplicateValues\"", xml);

            var rules = Excel.Open(file).Worksheets[0].ConditionalFormats;
            Assert.Equal(6, rules.Count);
            Assert.Equal(ConditionalFormatType.Duplicate, rules[5].Type);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void TimePeriod_Top10_AboveAverage_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.TimePeriod,
                Sqref = "A1:A10",
                TimePeriod = "thisMonth",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Top10,
                Sqref = "B1:B10",
                Rank = 3,
                Percent = true,
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.AboveAverage,
                Sqref = "C1:C10",
            });
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.BelowAverage,
                Sqref = "D1:D10",
            });
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("type=\"timePeriod\"", xml);
            Assert.Contains("timePeriod=\"thisMonth\"", xml);
            Assert.Contains("type=\"top10\"", xml);
            Assert.Contains("rank=\"3\"", xml);
            Assert.Contains("percent=\"1\"", xml);
            Assert.Contains("type=\"aboveAverage\"", xml);

            var rules = Excel.Open(file).Worksheets[0].ConditionalFormats;
            Assert.Equal(4, rules.Count);
            Assert.Equal(ConditionalFormatType.TimePeriod, rules[0].Type);
            Assert.Equal("thisMonth", rules[0].TimePeriod);
            Assert.Equal(ConditionalFormatType.Top10, rules[1].Type);
            Assert.Equal(3, rules[1].Rank);
            Assert.True(rules[1].Percent);
            Assert.Equal(ConditionalFormatType.AboveAverage, rules[2].Type);
            Assert.Equal(ConditionalFormatType.BelowAverage, rules[3].Type);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
