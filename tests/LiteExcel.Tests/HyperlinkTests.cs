using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 5：超链接读写测试（xlsx 优先）。
/// </summary>
public class HyperlinkTests
{
    [Fact]
    public void Write_ExternalHyperlink_RoundTrip()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("GitHub");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink
        {
            Target = "https://github.com/GitHubMacrro/LiteExcel.Net",
            Tooltip = "LiteExcel repo",
        };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            var opened = Excel.Open(path);
            var cell = opened.Worksheets[0].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal("https://github.com/GitHubMacrro/LiteExcel.Net", cell.Hyperlink.Target);
            Assert.Equal("LiteExcel repo", cell.Hyperlink.Tooltip);
            Assert.False(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Write_InternalHyperlink_RoundTrip()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Name = "Data";
        wb.Worksheets.Add("Summary");
        wb.Worksheets[1].Cell("A1").SetValue("Go to Data");
        wb.Worksheets[1].Cell("A1").Hyperlink = new Hyperlink
        {
            Target = "#Data!A1",
            IsInternal = true,
        };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            var opened = Excel.Open(path);
            Assert.Equal(2, opened.Worksheets.Count);
            var cell = opened.Worksheets[1].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal("#Data!A1", cell.Hyperlink.Target);
            Assert.True(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Write_InternalHyperlink_LocationNotExternalRels()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Name = "Data";
        wb.Worksheets.Add("Summary");
        wb.Worksheets[1].Cell("A1").SetValue("Go");
        wb.Worksheets[1].Cell("A1").Hyperlink = new Hyperlink { Target = "#Data!A1", IsInternal = true };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            using var zip = System.IO.Compression.ZipFile.OpenRead(path);
            var sheetEntry = zip.GetEntry("xl/worksheets/sheet2.xml");
            Assert.NotNull(sheetEntry);
            using var sr = new StreamReader(sheetEntry.Open());
            var xml = sr.ReadToEnd();
            Assert.Contains("location=\"Data!A1\"", xml);
            Assert.DoesNotContain("r:id", xml);

            var relsEntry = zip.GetEntry("xl/worksheets/_rels/sheet2.xml.rels");
            if (relsEntry is not null)
            {
                using var rr = new StreamReader(relsEntry.Open());
                Assert.DoesNotContain("hyperlink", rr.ReadToEnd());
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("mailto:test@example.com")]
    [InlineData("file:///C:/temp/plan.pdf")]
    [InlineData(@"\\server\share\report.xlsx")]
    public void Write_NonHttpExternalHyperlink_RoundTrip(string target)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Link");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink { Target = target };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            var opened = Excel.Open(path);
            var cell = opened.Worksheets[0].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal(target, cell.Hyperlink.Target);
            Assert.False(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Write_MultipleHyperlinks_OnSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Link1");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink { Target = "https://example.com/1" };
        wb.Worksheets[0].Cell("B2").SetValue("Link2");
        wb.Worksheets[0].Cell("B2").Hyperlink = new Hyperlink { Target = "https://example.com/2" };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            var opened = Excel.Open(path);
            Assert.Equal("https://example.com/1", opened.Worksheets[0].Cell("A1").Hyperlink?.Target);
            Assert.Equal("https://example.com/2", opened.Worksheets[0].Cell("B2").Hyperlink?.Target);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpenUserSample_WithHyperlink()
    {        // 用户提供的图片样本无超链接；用真实带超链接文件验证读取（若无则跳过）
        // 构造一个带超链接的文件后验证读取路径
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("External");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink { Target = "https://www.microsoft.com" };
        wb.Worksheets[0].Cell("A2").SetValue("Internal");
        wb.Worksheets[0].Cell("A2").Hyperlink = new Hyperlink { Target = "#Sheet1!B1", IsInternal = true };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            Assert.NotNull(opened.Worksheets[0].Cell("A1").Hyperlink);
            Assert.Equal("https://www.microsoft.com", opened.Worksheets[0].Cell("A1").Hyperlink.Target);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ── Phase 5b：xlsb / xls 超链接 ──

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void Write_ExternalHyperlink_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("GitHub");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink
        {
            Target = "https://github.com/GitHubMacrro/LiteExcel.Net",
            Tooltip = "LiteExcel repo",
        };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);

            var opened = Excel.Open(path, format);
            var cell = opened.Worksheets[0].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal("https://github.com/GitHubMacrro/LiteExcel.Net", cell.Hyperlink.Target);
            Assert.Equal("LiteExcel repo", cell.Hyperlink.Tooltip);
            Assert.False(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void Write_InternalHyperlink_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Name = "Data";
        wb.Worksheets.Add("Summary");
        wb.Worksheets[1].Cell("A1").SetValue("Go to Data");
        wb.Worksheets[1].Cell("A1").Hyperlink = new Hyperlink
        {
            Target = "#Data!A1",
            IsInternal = true,
        };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);

            var opened = Excel.Open(path, format);
            Assert.Equal(2, opened.Worksheets.Count);
            var cell = opened.Worksheets[1].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal("#Data!A1", cell.Hyperlink.Target);
            Assert.True(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("mailto:test@example.com", ".xlsb", ExcelFormat.Xlsb)]
    [InlineData("file:///C:/temp/plan.pdf", ".xlsb", ExcelFormat.Xlsb)]
    [InlineData(@"\\server\share\report.xlsx", ".xlsb", ExcelFormat.Xlsb)]
    [InlineData("mailto:test@example.com", ".xls", ExcelFormat.Xls)]
    [InlineData("file:///C:/temp/plan.pdf", ".xls", ExcelFormat.Xls)]
    public void Write_NonHttpExternalHyperlink_RoundTrip_Format(string target, string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Link");
        wb.Worksheets[0].Cell("A1").Hyperlink = new Hyperlink { Target = target };

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);

            var opened = Excel.Open(path, format);
            var cell = opened.Worksheets[0].Cell("A1");
            Assert.NotNull(cell.Hyperlink);
            Assert.Equal(target, cell.Hyperlink.Target);
            Assert.False(cell.Hyperlink.IsInternal);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
