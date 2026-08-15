using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// xls（BIFF8）写入测试。真实 Excel 读回验证见开发记录（Excel COM 打开确认值一致）。
/// </summary>
public class XlsWriteTests
{
    private static string GetTempFile() =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_xlsw_{Guid.NewGuid():N}.xls");

    private static void Delete(params string[] files)
    {
        foreach (var f in files)
            if (f is not null && File.Exists(f)) File.Delete(f);
    }

    private static Workbook BuildBasic()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.Name = "数据";
        ws.SetValue("A1", "姓名");
        ws.SetValue("A2", "张三");
        ws.SetValue("B2", 25);
        ws.SetValue("C2", new DateTime(2024, 5, 10));
        ws.SetValue("D2", true);
        ws.Merge("A1:B1");
        ws.ColumnWidths = new Dictionary<int, double> { [0] = 15 };
        ws.FreezeHeader = true;
        return wb;
    }

    [Fact]
    public void SaveAs_Xls_RoundTrips_Basic()
    {
        var file = GetTempFile();
        try
        {
            BuildBasic().SaveAs(file, ExcelFormat.Xls);
            Assert.True(File.Exists(file));
            Assert.True(new FileInfo(file).Length > 512, "xls 文件不应为空");

            var rb = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xls, rb.Format);
            var s = rb.Worksheets[0];
            Assert.Equal("数据", s.Name);
            Assert.Equal("姓名", s.Cell("A1").GetString());
            Assert.Equal("张三", s.Cell("A2").GetString());
            Assert.Equal(25.0, s.Cell("B2").GetDouble());
            Assert.Equal(new DateTime(2024, 5, 10), s.Cell("C2").GetDateTime());
            Assert.True(s.Cell("D2").GetBoolean());
            Assert.True(s.FreezeHeader);
            var merge = Assert.Single(s.MergedRanges);
            Assert.Equal(0, merge.FirstRow);
            Assert.Equal(1, merge.LastCol);
            Assert.Equal(15.0, s.ColumnWidths![0], 2);
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xls_MultiSheet_ChineseNames()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].Name = "数据";
            wb.Worksheets[0].SetValue("A1", "中文内容");
            var ws2 = wb.Worksheets.Add("汇总表");
            ws2.SetValue("B2", 42);
            wb.SaveAs(file, ExcelFormat.Xls);

            var rb = Excel.Open(file);
            Assert.Equal(2, rb.Worksheets.Count);
            Assert.Equal("数据", rb.Worksheets[0].Name);
            Assert.Equal("汇总表", rb.Worksheets[1].Name);
            Assert.Equal("中文内容", rb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(42.0, rb.Worksheets[1].Cell("B2").GetDouble());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xls_LargeSharedStrings_RoundTrips()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            const int rows = 3000, cols = 3;
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    ws.SetValue(r, c, (r + c - 2) % 3 == 0 ? $"文本{r}-{c}-中文内容" : $"text-{r}-{c}-ABC-xyz");
            wb.SaveAs(file, ExcelFormat.Xls);

            var s = Excel.Open(file).Worksheets[0];
            Assert.Equal(rows, s.RowCount);
            Assert.Equal("文本1-1-中文内容", s.Cell("A1").GetString());
            Assert.Equal("text-3000-3-ABC-xyz", s.Cell(rows, cols).GetString());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xlsb_ThrowsNotSupported()
    {
        var file = GetTempFile();
        var file2 = GetTempFile().Replace(".xls", ".xlsb");
        try
        {
            var wb = Excel.Create();
            wb.SaveAs(file, ExcelFormat.Xls);
            var rb = Excel.Open(file);
            Assert.Throws<NotSupportedException>(() => rb.SaveAs(file2, ExcelFormat.Xlsb));
        }
        finally { Delete(file, file2); }
    }

    [Fact]
    public void SaveAs_Xls_FromRealXlsFixture_RoundTrips()
    {
        // 打开真实 Excel 生成的 .xls，另存为 .xls 后重开，关键数据不丢
        var fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored.xls");
        Assert.True(File.Exists(fixture), $"Required fixture is missing: {fixture}");
        var file = GetTempFile();
        try
        {
            var wb = Excel.Open(fixture);
            wb.SaveAs(file, ExcelFormat.Xls);

            var rb = Excel.Open(file);
            var s = rb.Worksheets[0];
            Assert.Equal("数据", s.Name);
            Assert.Equal("姓名", s.Cell("A1").GetString());
            Assert.Equal("张三", s.Cell("A2").GetString());
            Assert.Equal(25.0, s.Cell("B2").GetDouble());
            Assert.Equal(new DateTime(2024, 5, 10), s.Cell("C2").GetDateTime());
            Assert.True(s.Cell("D2").GetBoolean());
            Assert.Equal(50.0, s.Cell("E2").GetDouble());
            Assert.True(s.FreezeHeader);
            var big = rb.Worksheets[1];
            Assert.Equal(3000, big.RowCount);
            Assert.Equal("text-3000-3-ABC-xyz", big.Cell(3000, 3).GetString());
        }
        finally { Delete(file); }
    }
}
