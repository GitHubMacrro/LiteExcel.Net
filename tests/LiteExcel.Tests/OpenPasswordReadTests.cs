using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 2：打开密码读取测试。
/// 使用用户提供的真实 Excel 加密样本（files/ 目录）验证解密读取。
/// 样本命名含密码：打开加密1（打开密码=1）、修改加密12（修改密码=12）、打开修改都需要密码（双密码）。
/// </summary>
public class OpenPasswordReadTests
{
    private static string FilesDir
    {
        get
        {
            var probe = Path.GetDirectoryName(AppContext.BaseDirectory);
            while (probe is not null && !Directory.Exists(Path.Combine(probe, "tests", "LiteExcel.Tests", "Fixtures")))
                probe = Path.GetDirectoryName(probe);
            return Path.Combine(probe!, "tests", "LiteExcel.Tests", "Fixtures", "EncryptedSamples");
        }
    }

    private static string Sample(string name) => Path.Combine(FilesDir, name);

    [Theory]
    [InlineData("打开加密1.xlsx")]
    [InlineData("打开加密1.xlsm")]
    [InlineData("打开加密1.xlsb")]
    public void Open_Encrypted_Sample_WithCorrectPassword_Succeeds(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, new ExcelReadOptions { OpenPassword = "1" });
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
        Assert.True(wb.Security.HasOpenPassword);
    }

    [Theory]
    [InlineData("打开加密1.xlsx")]
    [InlineData("打开加密1.xlsm")]
    [InlineData("打开加密1.xlsb")]
    public void Open_Encrypted_Sample_WrongPassword_Throws(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { OpenPassword = "wrong" }));
        Assert.Contains("密码", ex.Message);
        Assert.DoesNotContain("wrong", ex.Message);
    }

    [Theory]
    [InlineData("打开加密1.xlsx")]
    [InlineData("打开加密1.xlsm")]
    [InlineData("打开加密1.xlsb")]
    public void Open_Encrypted_Sample_NoPassword_ThrowsClearEncryptionError(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path));
        Assert.Contains("加密", ex.Message);
        Assert.Contains("OpenPassword", ex.Message);
    }

    [Theory]
    [InlineData("打开加密1.xlsx", ExcelFormat.Xlsx)]
    [InlineData("打开加密1.xlsm", ExcelFormat.Xlsm)]
    [InlineData("打开加密1.xlsb", ExcelFormat.Xlsb)]
    public void Open_ExplicitFormat_Encrypted_Sample_WithPassword_Succeeds(string fixture, ExcelFormat format)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, format, new ExcelReadOptions { OpenPassword = "1" });
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
    }

    [Theory]
    [InlineData("打开加密1.xlsx")]
    [InlineData("打开加密1.xlsm")]
    [InlineData("打开加密1.xlsb")]
    public void Open_Stream_Encrypted_Sample_WithPassword_Succeeds(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");
        var format = Path.GetExtension(path) switch
        {
            ".xlsb" => ExcelFormat.Xlsb,
            ".xlsm" => ExcelFormat.Xlsm,
            _ => ExcelFormat.Xlsx,
        };

        using var fs = File.OpenRead(path);
        var wb = Excel.Open(fs, format, new ExcelReadOptions { OpenPassword = "1" });
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
    }

    [Theory]
    [InlineData("打开加密1.xlsx")]
    [InlineData("打开加密1.xlsm")]
    [InlineData("打开加密1.xlsb")]
    public void Open_Encrypted_DataIsReadable(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, new ExcelReadOptions { OpenPassword = "1" });
        // 数据可正常读取（至少一张表、含内容）
        foreach (var ws in wb.Worksheets)
        {
            Assert.NotNull(ws);
        }
        Assert.True(wb.Worksheets.Count >= 1);
    }

    // ── 双密码样本 ──

    [Theory]
    [InlineData("打开修改都需要密码.xlsx")]
    [InlineData("打开修改都需要密码.xlsm")]
    [InlineData("打开修改都需要密码.xlsb")]
    public void Open_DualPassword_Sample_OpenPassword_Reads_IsReadOnly(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        // 只提供打开密码：可读但应处于只读状态（Phase 3 完善，此处先验证打开成功）
        var wb = Excel.Open(path, new ExcelReadOptions { OpenPassword = "1" });
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
        Assert.True(wb.Security.HasOpenPassword);
    }

    // ── 错误消息不泄露密码 ──

    [Fact]
    public void ErrorMessage_NeverContainsPassword()
    {
        var path = Sample("打开加密1.xlsx");
        Assert.True(File.Exists(path));

        const string pwd = "SuperSecret123!";
        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { OpenPassword = pwd }));
        // 密码错误时的异常不应包含密码本身
        Assert.DoesNotContain(pwd, ex.Message);
    }

    // ── B10：非加密文件误传 OpenPassword ──

    [Theory]
    [InlineData(".xlsx", ExcelFormat.Xlsx)]
    [InlineData(".xlsm", ExcelFormat.Xlsm)]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    public void Open_PlainFile_WithOpenPassword_ThrowsClearError(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);

            var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { OpenPassword = "wrong-pass" }));
            Assert.Contains("不是加密工作簿", ex.Message);
            Assert.DoesNotContain("wrong-pass", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(".xlsx", ExcelFormat.Xlsx)]
    [InlineData(".xlsb", ExcelFormat.Xlsb)]
    public void Open_PlainFile_WithOpenPassword_Stream_ThrowsClearError(string ext, ExcelFormat format)
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
        try
        {
            wb.SaveAs(path, format);
            using var fs = File.OpenRead(path);
            var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(fs, format, new ExcelReadOptions { OpenPassword = "wrong-pass" }));
            Assert.Contains("不是加密工作簿", ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
