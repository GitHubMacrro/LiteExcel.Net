using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// xlsb（BIFF12）写入测试。真实 Excel 打开验证见开发记录（Excel COM 打开确认值一致、无修复提示）。
/// </summary>
public class XlsbWriteTests
{
    private static string GetTempFile() =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_xlsbw_{Guid.NewGuid():N}.xlsb");

    private static void Delete(params string[] files)
    {
        foreach (var f in files)
            if (f is not null && File.Exists(f)) File.Delete(f);
    }

    private static Workbook BuildBasic()
    {
        var wb = Excel.Create(ExcelFormat.Xlsb);
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
    public void SaveAs_Xlsb_RoundTrips_Basic()
    {
        var file = GetTempFile();
        try
        {
            BuildBasic().SaveAs(file, ExcelFormat.Xlsb);
            Assert.True(File.Exists(file));
            Assert.True(new FileInfo(file).Length > 512, "xlsb 文件不应为空");

            var rb = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xlsb, rb.Format);
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
    public void SaveAs_Xlsb_MultiSheet_ChineseNames()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets["Sheet1"].SetValue("A1", "第一张");
            wb.Worksheets.Add("第二张表");
            wb.Worksheets["第二张表"].SetValue("B1", 99);
            wb.Worksheets.Add("三");
            wb.Worksheets["三"].SetValue("A2", "第三张");

            wb.SaveAs(file, ExcelFormat.Xlsb);

            var rb = Excel.Open(file);
            Assert.Equal(3, rb.Worksheets.Count);
            Assert.Equal(new[] { "Sheet1", "第二张表", "三" }, rb.Worksheets.Names);
            Assert.Equal("第一张", rb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(99.0, rb.Worksheets[1].Cell("B1").GetDouble());
            Assert.Equal("第三张", rb.Worksheets[2].Cell("A2").GetString());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xlsb_TextNumberBoolDate()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "text");
            ws.SetValue("B1", 3.14);
            ws.SetValue("C1", true);
            ws.SetValue("D1", new DateTime(2023, 12, 31));
            ws.SetValue("E1", -42);
            wb.SaveAs(file, ExcelFormat.Xlsb);

            var rb = Excel.Open(file);
            var s = rb.Worksheets[0];
            Assert.Equal("text", s.Cell("A1").GetString());
            Assert.Equal(3.14, s.Cell("B1").GetDouble(), 6);
            Assert.True(s.Cell("C1").GetBoolean());
            Assert.Equal(new DateTime(2023, 12, 31), s.Cell("D1").GetDateTime());
            Assert.Equal(-42.0, s.Cell("E1").GetDouble());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xlsb_ColumnWidthsRowHeights()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.SetValue("B1", "y"); // 让 B 列有数据，读回时保留宽度
            ws.SetValue("B2", "z"); // 让第 2 行有数据，读回时保留行高
            ws.ColumnWidths = new Dictionary<int, double> { [0] = 20, [1] = 9.5 };
            ws.RowHeights = new Dictionary<int, double> { [0] = 30, [1] = 25 };
            wb.SaveAs(file, ExcelFormat.Xlsb);

            var rb = Excel.Open(file);
            var s = rb.Worksheets[0];
            Assert.Equal(20.0, s.ColumnWidths![0], 1);
            Assert.Equal(9.5, s.ColumnWidths[1], 1);
            Assert.Equal(30.0, s.RowHeights![0], 1);
            Assert.Equal(25.0, s.RowHeights[1], 1);
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xlsb_EmptyCell_MiddleOfRow()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "left");
            ws.SetValue("C1", "right"); // B1 留空，验证 Short 单元格跳跃
            wb.SaveAs(file, ExcelFormat.Xlsb);

            var rb = Excel.Open(file);
            var s = rb.Worksheets[0];
            Assert.Equal("left", s.Cell("A1").GetString());
            Assert.Null(s.Cell("B1").GetValue());
            Assert.Equal("right", s.Cell("C1").GetString());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void SaveAs_Xlsb_LargeSharedStrings_RoundTrips()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            for (int i = 0; i < 500; i++)
                ws.SetValue(i + 1, 1, $"字符串{i:0000}_中文");
            wb.SaveAs(file, ExcelFormat.Xlsb);

            var rb = Excel.Open(file);
            Assert.Equal("字符串0000_中文", rb.Worksheets[0].Cell(1, 1).GetString());
            Assert.Equal("字符串0499_中文", rb.Worksheets[0].Cell(500, 1).GetString());
            Assert.Equal(500, rb.Worksheets[0].RowCount);
        }
        finally { Delete(file); }
    }
}
