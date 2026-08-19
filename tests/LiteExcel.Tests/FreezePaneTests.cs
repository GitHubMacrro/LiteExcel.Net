using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 6：冻结窗格增强测试（FreezeRows / FreezeColumns，xlsx/xlsb/xls）。
/// </summary>
public class FreezePaneTests
{
    [Fact]
    public void FreezeRows_RoundTrip()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeRows = 2;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            Assert.Equal(2, opened.Worksheets[0].FreezeRows);
            Assert.Equal(0, opened.Worksheets[0].FreezeColumns);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void FreezeColumns_RoundTrip()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeColumns = 3;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            Assert.Equal(3, opened.Worksheets[0].FreezeColumns);
            Assert.Equal(0, opened.Worksheets[0].FreezeRows);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void FreezeBoth_RoundTrip()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeRows = 1;
        wb.Worksheets[0].FreezeColumns = 2;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            Assert.Equal(1, opened.Worksheets[0].FreezeRows);
            Assert.Equal(2, opened.Worksheets[0].FreezeColumns);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void FreezeHeader_StillWorks_AsFreezeRows1()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeHeader = true;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            // FreezeHeader = true 兼容为 FreezeRows = 1
            Assert.Equal(1, opened.Worksheets[0].FreezeRows);
            Assert.True(opened.Worksheets[0].FreezeHeader);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void NoFreeze_NoPane()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var opened = Excel.Open(path);
            Assert.Equal(0, opened.Worksheets[0].FreezeRows);
            Assert.Equal(0, opened.Worksheets[0].FreezeColumns);
            Assert.False(opened.Worksheets[0].FreezeHeader);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ── Phase 5b/6 补充：xlsb / xls 冻结行列 ──

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void FreezeRows_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeRows = 2;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);
            var opened = Excel.Open(path, format);
            Assert.Equal(2, opened.Worksheets[0].FreezeRows);
            Assert.Equal(0, opened.Worksheets[0].FreezeColumns);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void FreezeColumns_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeColumns = 3;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);
            var opened = Excel.Open(path, format);
            Assert.Equal(3, opened.Worksheets[0].FreezeColumns);
            Assert.Equal(0, opened.Worksheets[0].FreezeRows);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void FreezeBoth_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeRows = 2;
        wb.Worksheets[0].FreezeColumns = 3;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);
            var opened = Excel.Open(path, format);
            Assert.Equal(2, opened.Worksheets[0].FreezeRows);
            Assert.Equal(3, opened.Worksheets[0].FreezeColumns);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    [InlineData(".xls", ExcelFormat.Xls)]
    public void FreezeHeader_RoundTrip_Format(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Worksheets[0].FreezeHeader = true;

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);
            var opened = Excel.Open(path, format);
            Assert.Equal(1, opened.Worksheets[0].FreezeRows);
            Assert.True(opened.Worksheets[0].FreezeHeader);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
