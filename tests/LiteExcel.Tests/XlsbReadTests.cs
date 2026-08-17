using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// xlsb（BIFF12）读取测试。合成文件覆盖基础路径，真实 Excel 文件见
/// <see cref="Open_ExcelAuthored_*"/> 系列（Fixtures/excel-authored.xlsb）。
/// </summary>
public class XlsbReadTests
{
    private static string GetTempFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsb_test_{Guid.NewGuid():N}.xlsb");
        File.WriteAllText(path, "");
        return path;
    }

    private static void Delete(params string[] files)
    {
        foreach (var f in files)
            if (f is not null && File.Exists(f)) File.Delete(f);
    }

    private static XlsbTestFile.WorkbookSpec BasicSpec()
    {
        var spec = new XlsbTestFile.WorkbookSpec();
        var sheet = new XlsbTestFile.SheetSpec { Name = "数据" };
        sheet.Rows.Add(new XlsbTestFile.RowSpec
        {
            Cells =
            {
                new XlsbTestFile.CellSpec { Col = 0, Text = "姓名" },
                new XlsbTestFile.CellSpec { Col = 1, Number = 25 },
                new XlsbTestFile.CellSpec { Col = 2, Bool = true },
                new XlsbTestFile.CellSpec { Col = 3, InlineText = true, Text = "内联" },
            },
        });
        sheet.Rows.Add(new XlsbTestFile.RowSpec
        {
            Cells =
            {
                new XlsbTestFile.CellSpec { Col = 0, Text = "张三" },
                new XlsbTestFile.CellSpec { Col = 1, Number = 3.5 },
            },
        });
        spec.Sheets.Add(sheet);
        return spec;
    }

    [Fact]
    public void Open_ReadsBasicCells()
    {
        var file = XlsbTestFile.Build(BasicSpec());
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xlsb, wb.Format);
            var s = wb.Worksheets[0];
            Assert.Equal("数据", s.Name);
            Assert.Equal("姓名", s.Cell("A1").GetString());
            Assert.Equal(25.0, s.Cell("B1").GetDouble());
            Assert.True(s.Cell("C1").GetBoolean());
            Assert.Equal("内联", s.Cell("D1").GetString());
            Assert.Equal("张三", s.Cell("A2").GetString());
            Assert.Equal(3.5, s.Cell("B2").GetDouble());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void Open_MultiSheet_ReadsSheetNames()
    {
        var spec = BasicSpec();
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "Sheet2" });
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "汇总表" });
        var file = XlsbTestFile.Build(spec);
        try
        {
            var wb = Excel.Open(file);
            Assert.Equal(3, wb.Worksheets.Count);
            Assert.Equal("数据", wb.Worksheets[0].Name);
            Assert.Equal("Sheet2", wb.Worksheets[1].Name);
            Assert.Equal("汇总表", wb.Worksheets[2].Name);
        }
        finally { Delete(file); }
    }

    [Fact]
    public void Open_ReadsDateViaStyles()
    {
        var spec = BasicSpec();
        spec.Formats[176] = "yyyy\\-mm\\-dd";
        spec.CellXfs.Clear();
        spec.CellXfs.AddRange(new[] { 0, 0, 176, 0 });
        var dateSheet = new XlsbTestFile.SheetSpec { Name = "Dates" };
        dateSheet.Rows.Add(new XlsbTestFile.RowSpec
        {
            Cells =
            {
                new XlsbTestFile.CellSpec { Col = 0, Number = 45422, Style = 2 },
                new XlsbTestFile.CellSpec { Col = 1, Number = 25, Style = 1 },
            },
        });
        spec.Sheets.Add(dateSheet);
        var file = XlsbTestFile.Build(spec);
        try
        {
            var s = Excel.Open(file).Worksheets[1];
            Assert.Equal(CellType.Date, s.Cell("A1").Type);
            Assert.Equal(new DateTime(2024, 5, 10), s.Cell("A1").GetDateTime());
            Assert.Equal(CellType.Number, s.Cell("B1").Type);
        }
        finally { Delete(file); }
    }

    [Fact]
    public void Open_ReadsMergeFreezeColWidth()
    {
        var spec = BasicSpec();
        var sheet = spec.Sheets[0];
        sheet.Merges.Add((0, 0, 0, 1));
        sheet.FrozenRows = 1;
        sheet.ColWidths[0] = 15.0;
        var file = XlsbTestFile.Build(spec);
        try
        {
            var s = Excel.Open(file).Worksheets[0];
            Assert.True(s.FreezeHeader);
            var merge = Assert.Single(s.MergedRanges);
            Assert.Equal(0, merge.FirstRow);
            Assert.Equal(1, merge.LastCol);
            Assert.Equal(15.0, s.ColumnWidths![0], 2);
        }
        finally { Delete(file); }
    }

    [Fact]
    public void Open_SharedString_ContinuationIsNotRequired()
    {
        // 合成文件：SST 含两个字符串，均经 BrtSSTItem 单记录写入
        var spec = BasicSpec();
        var file = XlsbTestFile.Build(spec);
        try
        {
            var s = Excel.Open(file).Worksheets[0];
            Assert.Equal("姓名", s.Cell("A1").GetString());
            Assert.Equal("张三", s.Cell("A2").GetString());
        }
        finally { Delete(file); }
    }

    [Fact]
    public void Open_InvalidFile_Throws()
    {
        var file = GetTempFile();
        try
        {
            File.WriteAllBytes(file, new byte[] { 0x00, 0x01, 0x02 });
            Assert.ThrowsAny<Exception>(() => Excel.Open(file));
        }
        finally { Delete(file); }
    }

    [Fact]
    public void OpenXlsb_SaveAsXlsb_RoundTrips()
    {
        var file = XlsbTestFile.Build(BasicSpec());
        var outFile = GetTempFile();
        try
        {
            var wb = Excel.Open(file);
            wb.SaveAs(outFile, ExcelFormat.Xlsb);
            Assert.True(File.Exists(outFile));

            var rb = Excel.Open(outFile);
            Assert.Equal("姓名", rb.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(25.0, rb.Worksheets[0].Cell("B1").GetDouble());
        }
        finally { Delete(file, outFile); }
    }

    // ── 真实文件测试：由 Microsoft Excel 生成（SaveAs xlExcel12） ──

    private static string FixturePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored.xlsb");

    private static string GetFixturePath()
    {
        Assert.True(File.Exists(FixturePath), $"Required Excel xlsb fixture is missing: {FixturePath}");
        return FixturePath;
    }

    [Fact]
    public void Open_ExcelAuthored_ReadsSheetNamesAndCells()
    {
        var wb = Excel.Open(GetFixturePath());
        Assert.Equal(ExcelFormat.Xlsb, wb.Format);
        Assert.Equal(2, wb.Worksheets.Count);
        Assert.Equal("数据", wb.Worksheets[0].Name);
        Assert.Equal("大数据", wb.Worksheets[1].Name);

        var s = wb.Worksheets[0];
        Assert.Equal("姓名", s.Cell("A1").GetString());
        Assert.Equal("张三", s.Cell("A2").GetString());
        Assert.Equal(25.0, s.Cell("B2").GetDouble());
        Assert.Equal(new DateTime(2024, 5, 10), s.Cell("C2").GetDateTime());
        Assert.True(s.Cell("D2").GetBoolean());
        Assert.Equal(50.0, s.Cell("E2").GetDouble()); // 公式缓存值 =B2*2
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
        // 9000 个唯一字符串
        var s = Excel.Open(GetFixturePath()).Worksheets[1];
        Assert.Equal(3000, s.RowCount);
        Assert.Equal(3, s.MaxColumn);
        Assert.Equal("文本1-1-中文内容", s.Cell("A1").GetString());
        Assert.Equal("text-1-2-ABC-xyz", s.Cell("B1").GetString());
        Assert.Equal("text-2-1-ABC-xyz", s.Cell("A2").GetString());
        Assert.Equal("text-3000-3-ABC-xyz", s.Cell(3000, 3).GetString());
    }
}
