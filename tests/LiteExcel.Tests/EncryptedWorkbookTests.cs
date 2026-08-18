using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 加密工作簿识别测试。
/// fixtures 由真实 Excel COM 生成（带打开密码），验证打开时给出明确加密异常，
/// 而不是误报 zip 损坏或解析出乱数据。
/// </summary>
public class EncryptedWorkbookTests
{
    private static string FixturePath(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Theory]
    [InlineData("protected.xlsx")]
    [InlineData("protected.xlsm")]
    [InlineData("protected.xlsb")]
    public void Open_EncryptedOoxml_ThrowsClearEncryptionError(string fixture)
    {
        var path = FixturePath(fixture);
        Assert.True(File.Exists(path), $"Required encrypted fixture is missing: {path}");

        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path));

        Assert.Contains("加密", ex.Message);
        Assert.DoesNotContain("zip", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ZipArchive", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Open_EncryptedXls_ThrowsClearEncryptionError()
    {
        var path = FixturePath("protected.xls");
        Assert.True(File.Exists(path), $"Required encrypted fixture is missing: {path}");

        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path));

        Assert.Contains("加密", ex.Message);
    }

    [Theory]
    [InlineData("protected.xlsx", ExcelFormat.Xlsx)]
    [InlineData("protected.xlsm", ExcelFormat.Xlsm)]
    [InlineData("protected.xlsb", ExcelFormat.Xlsb)]
    [InlineData("protected.xls", ExcelFormat.Xls)]
    public void Open_ExplicitFormat_EncryptedFile_ThrowsClearEncryptionError(string fixture, ExcelFormat format)
    {
        var path = FixturePath(fixture);
        Assert.True(File.Exists(path), $"Required encrypted fixture is missing: {path}");

        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, format));

        Assert.Contains("加密", ex.Message);
    }

    [Theory]
    [InlineData("excel-authored-compatibility.xlsx")]
    [InlineData("excel-authored.xls")]
    [InlineData("excel-authored.xlsb")]
    public void Open_PlainExcelFiles_UnaffectedByEncryptionDetection(string fixture)
    {
        var path = FixturePath(fixture);
        Assert.True(File.Exists(path), $"Required fixture is missing: {path}");

        // 非加密文件必须照常打开，证明加密检测不误伤
        var wb = Excel.Open(path);
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
    }

    // ── 公开 path 读取入口的加密识别覆盖 ──

    [Fact]
    public void XlsxReader_Read_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() => XlsxReader.Read(path, 0));
        Assert.Contains("加密", ex.Message);
    }

    [Fact]
    public void XlsxReader_ReadAll_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() => XlsxReader.ReadAll(path));
        Assert.Contains("加密", ex.Message);
    }

    [Fact]
    public void XlsxReader_GetSheetNames_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() => XlsxReader.GetSheetNames(path));
        Assert.Contains("加密", ex.Message);
    }

    [Fact]
    public void XlsxReader_StreamRows_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() =>
            XlsxReader.StreamRows(path, "Sheet1", _ => { }));
        Assert.Contains("加密", ex.Message);
    }

    [Fact]
    public void Excel_ReadAsDataTable_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() => Excel.ReadAsDataTable(path));
        Assert.Contains("加密", ex.Message);
    }

    [Fact]
    public void Excel_GetSheetNames_EncryptedXlsx_ThrowsClearError()
    {
        var path = FixturePath("protected.xlsx");
        Assert.True(File.Exists(path));
        var ex = Assert.Throws<LiteExcelException>(() => Excel.GetSheetNames(path));
        Assert.Contains("加密", ex.Message);
    }
}
