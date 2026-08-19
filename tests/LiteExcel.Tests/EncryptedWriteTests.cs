using LiteExcel;
using LiteExcel.Internal.Encryption;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 4：密码保存与加密写出测试。
/// 验证 SaveAs 带打开密码后能生成加密文件、能被自己解密读取、round-trip 一致。
/// </summary>
public class EncryptedWriteTests
{
    [Fact]
    public void SaveAs_WithOpenPassword_ProducesEncryptedFile()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Hello");
        wb.Security.SetOpenPassword("secret-123");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            // 文件应已是 CFB 加密容器（非普通 zip）
            var bytes = File.ReadAllBytes(path);
            Assert.Equal(0xD0, bytes[0]);
            Assert.Equal(0xCF, bytes[1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SaveAs_WithOpenPassword_CanBeOpenedWithPassword()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("RoundtripData");
        wb.Security.SetOpenPassword("secret-123");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            var opened = Excel.Open(path, new ExcelReadOptions { OpenPassword = "secret-123" });
            Assert.NotNull(opened);
            Assert.Equal("RoundtripData", opened.Worksheets[0].Cell("A1").GetValue()?.ToString());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SaveAs_WithOpenPassword_WrongPassword_FailsToOpen()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("SomeContent");
        wb.Security.SetOpenPassword("secret-123");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { OpenPassword = "wrong" }));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SaveAs_AfterRemoveOpenPassword_ProducesPlainFile()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Security.SetOpenPassword("secret-123");
        wb.Security.RemoveOpenPassword();

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);
            var bytes = File.ReadAllBytes(path);
            // 应为普通 zip（PK 头）
            Assert.Equal(0x50, bytes[0]);
            Assert.Equal(0x4B, bytes[1]);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpenEncrypted_ThenSaveAs_PreservesOpenPassword()
    {
        // 打开真实加密样本，SaveAs 后应仍是加密文件
        var src = ModifyPasswordTests.SamplePath("打开加密1.xlsx");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { OpenPassword = "1" });
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(outPath);
            var bytes = File.ReadAllBytes(outPath);
            Assert.Equal(0xD0, bytes[0]); // 仍是 CFB 加密

            // 用原密码能打开
            var reopened = Excel.Open(outPath, new ExcelReadOptions { OpenPassword = "1" });
            Assert.NotNull(reopened);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void OpenEncrypted_RemovePassword_SaveAs_ProducesPlainFile()
    {
        var src = ModifyPasswordTests.SamplePath("打开加密1.xlsx");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { OpenPassword = "1" });
        wb.Security.RemoveOpenPassword();
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(outPath);
            var bytes = File.ReadAllBytes(outPath);
            Assert.Equal(0x50, bytes[0]); // 无密码 = 普通 zip
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void SaveAs_Xlsm_WithOpenPassword_PreservesMacro()
    {
        // 用真实 xlsm 加密样本验证宏保留（打开后 SaveAs 仍加密 + 宏不丢）
        var src = ModifyPasswordTests.SamplePath("打开加密1.xlsm");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { OpenPassword = "1" });
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsm");
        try
        {
            wb.SaveAs(outPath);
            var bytes = File.ReadAllBytes(outPath);
            Assert.Equal(0xD0, bytes[0]); // 仍是加密

            var reopened = Excel.Open(outPath, new ExcelReadOptions { OpenPassword = "1" });
            Assert.NotNull(reopened);
            // 宏字节应保留
            if (wb.VbaProjectBytes is not null)
                Assert.NotNull(reopened.VbaProjectBytes);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    // ── B9：EncryptedPackage 被篡改 → 完整性校验失败 ──

    [Fact]
    public void Open_TamperedEncryptedPackage_IntegrityFails()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Security.SetOpenPassword("secret-123");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            // 篡改 EncryptedPackage 主体字节（文件中部，位于大 payload 流内，避开 CFB 头/目录尾部）
            var bytes = File.ReadAllBytes(path);
            int idx = bytes.Length / 2;
            bytes[idx] ^= 0xFF;
            File.WriteAllBytes(path, bytes);

            var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { OpenPassword = "secret-123" }));
            // 密码正确但完整性校验失败 → 明确报篡改
            Assert.True(ex.Message.Contains("完整性") || ex.Message.Contains("篡改"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
