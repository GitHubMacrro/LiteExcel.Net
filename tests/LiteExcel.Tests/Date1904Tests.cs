using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 1904 日期系统专项测试。
/// 验证：Workbook.Date1904 属性在三格式写出侧正确写回标志，
/// 日期序列按 1904 基准（1904-01-01=0，比 1900 少 1462 天），
/// 以及读取真实 Excel 生成的 1904 文件换算正确。
/// </summary>
public class Date1904Tests
{
    private static readonly DateTime D1 = new(2024, 3, 15);
    private static readonly DateTime D2 = new(1904, 1, 1);
    private static readonly DateTime D3 = new(1904, 1, 6);
    private static readonly DateTime D4 = new(2000, 12, 31);

    private static Workbook Create1904Workbook()
    {
        var wb = Excel.Create(ExcelFormat.Xlsx);
        wb.Date1904 = true;
        wb.Worksheets.Remove("Sheet1");
        var ws = wb.Worksheets.Add("日期");
        ws.SetValue("A1", "日期");
        ws.SetValue("A2", D1);
        ws.SetValue("A3", D2);
        ws.SetValue("A4", D3);
        ws.SetValue("A5", D4);
        return wb;
    }

    private static void Assert1904RoundTrip(string path)
    {
        var wb = Excel.Open(path);
        Assert.True(wb.Date1904, $"打开 {path} 应识别 1904 日期系统");
        var ws = wb.Worksheets[0];
        Assert.Equal(D1, ws.Cell("A2").GetDateTime().Date);
        Assert.Equal(D2, ws.Cell("A3").GetDateTime().Date);
        Assert.Equal(D3, ws.Cell("A4").GetDateTime().Date);
        Assert.Equal(D4, ws.Cell("A5").GetDateTime().Date);
    }

    [Fact]
    public void WriteXlsx_1904_FlagAndDates_RoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), "liteexcel-1904.xlsx");
        var wb = Create1904Workbook();
        wb.SaveAs(path, ExcelFormat.Xlsx);
        Assert1904RoundTrip(path);
        File.Delete(path);
    }

    [Fact]
    public void WriteXlsb_1904_FlagAndDates_RoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), "liteexcel-1904.xlsb");
        var wb = Create1904Workbook();
        wb.SaveAs(path, ExcelFormat.Xlsb);
        Assert1904RoundTrip(path);
        File.Delete(path);
    }

    [Fact]
    public void WriteXls_1904_FlagAndDates_RoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), "liteexcel-1904.xls");
        var wb = Create1904Workbook();
        wb.SaveAs(path, ExcelFormat.Xls);
        Assert1904RoundTrip(path);
        File.Delete(path);
    }

    [Theory]
    [InlineData(ExcelFormat.Xlsb)]
    [InlineData(ExcelFormat.Xls)]
    public void Convert_1904Xlsx_ToOtherFormats_PreservesDates(ExcelFormat format)
    {
        // 1904 xlsx -> xlsb / xls 转换链，日期不得偏移
        var xlsxPath = Path.Combine(Path.GetTempPath(), "liteexcel-1904-src.xlsx");
        var outPath = Path.Combine(Path.GetTempPath(), "liteexcel-1904-out" + (format == ExcelFormat.Xlsb ? ".xlsb" : ".xls"));
        var wb = Create1904Workbook();
        wb.SaveAs(xlsxPath, ExcelFormat.Xlsx);

        var opened = Excel.Open(xlsxPath);
        Assert.True(opened.Date1904);
        opened.SaveAs(outPath, format);
        Assert1904RoundTrip(outPath);

        File.Delete(xlsxPath);
        File.Delete(outPath);
    }

    [Fact]
    public void Open_RealExcel1904Xlsb_ReadsCorrectDates()
    {
        // 真实 Excel 生成的 1904 xlsb（Excel 打开 1904 注入文件后另存）
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-date1904.xlsb");
        Assert.True(File.Exists(path), $"Required 1904 fixture is missing: {path}");

        var wb = Excel.Open(path);
        Assert.True(wb.Date1904);
        var ws = wb.Worksheets[0];
        Assert.Equal(D1, ws.Cell("A2").GetDateTime().Date);
        Assert.Equal(D2, ws.Cell("A3").GetDateTime().Date);
        Assert.Equal(D3, ws.Cell("A4").GetDateTime().Date);
        Assert.Equal(D4, ws.Cell("A5").GetDateTime().Date);
    }
}
