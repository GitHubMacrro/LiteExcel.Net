using LiteExcel;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>传统 .xls（BIFF8）读取测试：用程序化构造的最小 xls 文件验证 XlsBackend 与 Excel.Open </summary>
public class XlsReadTests
{
    private static string GetTempFile() =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_xls_{Guid.NewGuid():N}.xls");

    private static string WriteTemp(byte[] bytes)
    {
        var file = GetTempFile();
        File.WriteAllBytes(file, bytes);
        return file;
    }

    [Fact]
    public void Open_ReadsBasicCells()
    {
        var sheet = new XlsTestFile.SheetSpec { Name = "数据" };
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "姓名" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 1, Kind = CellType.Text, Text = "年龄" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Text, Text = "张三" });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 1, Kind = CellType.Number, Number = 25 });

        var file = WriteTemp(XlsTestFile.Build(sheet));
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xls, wb.Format);
            var s = wb.Worksheets[0];
            Assert.Equal("数据", s.Name);
            Assert.Equal("姓名", s.Cell("A1").GetString());
            Assert.Equal("年龄", s.Cell("B1").GetString());
            Assert.Equal("张三", s.Cell("A2").GetString());
            Assert.Equal(25.0, s.Cell("B2").GetDouble());
            Assert.Equal(2, s.RowCount);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_ReadsDateAndBoolean()
    {
        var date = new DateTime(2024, 5, 10);
        var sheet = new XlsTestFile.SheetSpec();
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Date, Number = date.ToOADate() });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 1, Kind = CellType.Boolean, Bool = true });

        var file = WriteTemp(XlsTestFile.Build(sheet));
        try
        {
            var wb = Excel.Open(file);
            var s = wb.Worksheets[0];
            Assert.Equal(date, s.Cell("A1").GetDateTime());
            Assert.True(s.Cell("B1").GetBoolean());
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_ReadsMergedCellsColumnWidthsFreeze()
    {
        var sheet = new XlsTestFile.SheetSpec { FreezeHeader = true };
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "h" });
        sheet.Merges.Add((0, 1, 0, 1));
        sheet.ColWidths.Add((0, 20.0));

        var file = WriteTemp(XlsTestFile.Build(sheet));
        try
        {
            var wb = Excel.Open(file);
            var s = wb.Worksheets[0];
            Assert.True(s.FreezeHeader);
            var merge = Assert.Single(s.MergedRanges);
            Assert.Equal(0, merge.FirstRow);
            Assert.Equal(1, merge.LastRow);
            Assert.NotNull(s.ColumnWidths);
            Assert.Equal(20.0, s.ColumnWidths![0], 1);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_MultiSheet_ReadsSheetNames()
    {
        var s1 = new XlsTestFile.SheetSpec { Name = "第一张" };
        s1.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, Text = "a" });
        var s2 = new XlsTestFile.SheetSpec { Name = "Second" };
        s2.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Number, Number = 1 });

        var file = WriteTemp(XlsTestFile.Build(s1, s2));
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(2, wb.Worksheets.Count);
            Assert.Equal("第一张", wb.Worksheets[0].Name);
            Assert.Equal("Second", wb.Worksheets[1].Name);
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_InvalidBytes_ThrowsLiteExcelException()
    {
        var file = GetTempFile();
        File.WriteAllBytes(file, new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 });
        try
        {
            Assert.Throws<LiteExcelException>(() => Excel.Open(file));
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_ReadsSharedStrings()
    {
        var sheet = new XlsTestFile.SheetSpec();
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 0, Kind = CellType.Text, UseSst = true, SstIndex = 0 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 0, Col = 1, Kind = CellType.Text, UseSst = true, SstIndex = 1 });
        sheet.Cells.Add(new XlsTestFile.CellSpec { Row = 1, Col = 0, Kind = CellType.Text, UseSst = true, SstIndex = 0 });

        var file = WriteTemp(XlsTestFile.Build(new[] { "名称", "阿尔法" }, sheet));
        try
        {
            var wb = Excel.Open(file);
            var s = wb.Worksheets[0];
            Assert.Equal("名称", s.Cell("A1").GetString());
            Assert.Equal("阿尔法", s.Cell("B1").GetString());
            Assert.Equal("名称", s.Cell("A2").GetString());
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Open_ReadsSharedStringSplitAcrossContinue()
    {
        var workbook = XlsTestFile.BuildSstSplitWorkbook();
        var file = WriteTemp(XlsTestFile.BuildCfbFromWorkbook(workbook));
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal("你好世界", wb.Worksheets[0].Cell("A1").GetString());
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void SaveAs_Xls_ThrowsNotSupported()
    {
        // xls 写入暂不支持
        var file = WriteTemp(XlsTestFile.Build(new XlsTestFile.SheetSpec()));
        try
        {
            var wb = Excel.Open(file);
            Assert.Throws<NotSupportedException>(() => wb.SaveAs(GetTempFile()));
        }
        finally { File.Delete(file); }
    }

    // ── 真实文件测试：由 Microsoft Excel 生成（Excel COM SaveAs xlExcel8） ──

    private static string FixturePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored.xls");

    private static string GetFixturePath()
    {
        Assert.True(File.Exists(FixturePath), $"Required Excel xls fixture is missing: {FixturePath}");
        return FixturePath;
    }

    [Fact]
    public void Open_ExcelAuthored_ReadsSheetNamesAndCells()
    {
        var wb = Excel.Open(GetFixturePath());
        Assert.Equal(ExcelFormat.Xls, wb.Format);
        Assert.Equal(2, wb.Worksheets.Count);
        Assert.Equal("数据", wb.Worksheets[0].Name);
        Assert.Equal("大数据", wb.Worksheets[1].Name);

        var s = wb.Worksheets[0];
        Assert.Equal("姓名", s.Cell("A1").GetString());
        Assert.Equal("张三", s.Cell("A2").GetString());
        Assert.Equal(25.0, s.Cell("B2").GetDouble());
        Assert.Equal(new DateTime(2024, 5, 10), s.Cell("C2").GetDateTime());
        Assert.True(s.Cell("D2").GetBoolean());
        Assert.Equal(50.0, s.Cell("E2").GetDouble());
        Assert.Equal(CellType.Empty, s.Cell("B1").Type); // 合并 A1:B1 后 Excel 清空 B1
    }

    [Fact]
    public void Open_ExcelAuthored_ReadsMergeFreezeColumnWidth()
    {
        var s = Excel.Open(GetFixturePath()).Worksheets[0];
        Assert.True(s.FreezeHeader);
        var merge = Assert.Single(s.MergedRanges);
        Assert.Equal(0, merge.FirstRow);
        Assert.Equal(0, merge.LastRow);
        Assert.NotNull(s.ColumnWidths);
        Assert.InRange(s.ColumnWidths![0], 14.0, 16.0);
    }

    [Fact]
    public void Open_ExcelAuthored_ReadsLargeSharedStringSheet()
    {
        // 9000 个唯一字符串，Excel 写入时 SST 跨多个 CONTINUE 记录
        var s = Excel.Open(GetFixturePath()).Worksheets[1];
        Assert.Equal(3000, s.RowCount);
        Assert.Equal(3, s.MaxColumn);
        Assert.Equal("文本1-1-中文内容", s.Cell("A1").GetString());
        Assert.Equal("text-1-2-ABC-xyz", s.Cell("B1").GetString());
        Assert.Equal("text-2-1-ABC-xyz", s.Cell("A2").GetString());
        Assert.Equal("text-3000-3-ABC-xyz", s.Cell(3000, 3).GetString());
    }
}
