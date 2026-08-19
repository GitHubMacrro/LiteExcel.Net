using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 3：修改密码识别与处理测试。
/// 使用真实 Excel 样本（files/）验证 fileSharing 识别、只读状态、双密码组合。
/// 样本：修改加密12.xlsx/xlsm/xlsb（修改密码=12）、打开修改都需要密码.xlsx/xlsm/xlsb（打开=1+修改=12）。
/// </summary>
public class ModifyPasswordTests
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

    internal static string SamplePath(string name) => Sample(name);

    // ── 仅修改密码（无打开密码）──

    [Theory]
    [InlineData("修改加密12.xlsx")]
    [InlineData("修改加密12.xlsm")]
    [InlineData("修改加密12.xlsb")]
    public void Open_ModifyProtected_NoPassword_IsReadOnly(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path);
        Assert.NotNull(wb);
        Assert.True(wb.Worksheets.Count > 0);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.HasModifyAccess);
        Assert.True(wb.Security.IsReadOnly);
        Assert.False(wb.Security.CanSave);
    }

    [Theory]
    [InlineData("修改加密12.xlsx")]
    [InlineData("修改加密12.xlsm")]
    [InlineData("修改加密12.xlsb")]
    public void Open_ModifyProtected_NoPassword_SaveThrows(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path);
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(path));
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(outPath));
            Assert.Contains("只读", ex.Message);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Theory]
    [InlineData("修改加密12.xlsx")]
    [InlineData("修改加密12.xlsm")]
    [InlineData("修改加密12.xlsb")]
    public void Open_ModifyProtected_WithPassword_GrantsAccess(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, new ExcelReadOptions { ModifyPassword = "12" });
        Assert.NotNull(wb);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.True(wb.Security.HasModifyAccess);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
    }

    [Theory]
    [InlineData("修改加密12.xlsx")]
    [InlineData("修改加密12.xlsm")]
    [InlineData("修改加密12.xlsb")]
    public void Open_ModifyProtected_WithPassword_SaveSucceeds(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, new ExcelReadOptions { ModifyPassword = "12" });
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + Path.GetExtension(path));
        try
        {
            // 有修改权限应能保存
            wb.SaveAs(outPath);
            Assert.True(File.Exists(outPath));
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    // ── 双密码：打开 + 修改 ──

    [Theory]
    [InlineData("打开修改都需要密码.xlsx")]
    [InlineData("打开修改都需要密码.xlsm")]
    [InlineData("打开修改都需要密码.xlsb")]
    public void Open_DualPassword_OnlyOpen_IsReadOnly(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        // 只提供打开密码：能读但只读
        var wb = Excel.Open(path, new ExcelReadOptions { OpenPassword = "1" });
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.True(wb.Security.IsReadOnly);
        Assert.False(wb.Security.CanSave);
    }

    [Theory]
    [InlineData("打开修改都需要密码.xlsx")]
    [InlineData("打开修改都需要密码.xlsm")]
    [InlineData("打开修改都需要密码.xlsb")]
    public void Open_DualPassword_Both_CanSave(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        var wb = Excel.Open(path, new ExcelReadOptions { OpenPassword = "1", ModifyPassword = "12" });
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
    }

    [Theory]
    [InlineData("打开修改都需要密码.xlsx")]
    [InlineData("打开修改都需要密码.xlsm")]
    [InlineData("打开修改都需要密码.xlsb")]
    public void Open_DualPassword_OnlyModify_NoOpenPassword_Throws(string fixture)
    {
        var path = Sample(fixture);
        Assert.True(File.Exists(path), $"Missing sample: {path}");

        // 只有修改密码但没有打开密码：解密失败，明确报错
        var ex = Assert.Throws<LiteExcelException>(() => Excel.Open(path, new ExcelReadOptions { ModifyPassword = "12" }));
        Assert.Contains("打开密码", ex.Message);
    }

    // ── 修改密码写出 ──

    [Fact]
    public void SaveAs_WithSetModifyPassword_WritesFileSharing()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Security.SetModifyPassword("modify-secret");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            // 重新打开：应识别到修改密码（只读）
            var reopened = Excel.Open(path);
            Assert.True(reopened.Security.HasModifyPassword);
            Assert.False(reopened.Security.HasModifyAccess);
            Assert.True(reopened.Security.IsReadOnly);
            Assert.False(reopened.Security.CanSave);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SaveAs_WithSetModifyPassword_ProvidePassword_CanSave()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Security.SetModifyPassword("modify-secret");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(path);

            // 提供修改密码：获得权限可保存
            var reopened = Excel.Open(path, new ExcelReadOptions { ModifyPassword = "modify-secret" });
            Assert.False(reopened.Security.IsReadOnly);
            Assert.True(reopened.Security.CanSave);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpenModifyProtected_ThenSaveAs_PreservesFileSharing()
    {
        var src = Sample("修改加密12.xlsx");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { ModifyPassword = "12" });
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(outPath);

            // 未改修改密码：透传保留原 fileSharing
            var reopened = Excel.Open(outPath);
            Assert.True(reopened.Security.HasModifyPassword);
            Assert.True(reopened.Security.IsReadOnly);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void OpenModifyProtected_RemoveModifyPassword_SaveAs_NoProtection()
    {
        var src = Sample("修改加密12.xlsx");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { ModifyPassword = "12" });
        wb.Security.RemoveModifyPassword();
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            wb.SaveAs(outPath);

            var reopened = Excel.Open(outPath);
            Assert.False(reopened.Security.HasModifyPassword);
            Assert.False(reopened.Security.IsReadOnly);
            Assert.True(reopened.Security.CanSave);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    // ── Phase 5b 补：xlsb 修改密码写出 ──

    [Fact]
    public void SaveAs_Xlsb_WithSetModifyPassword_WritesFileSharing()
    {
        var wb = Excel.Create();
        wb.Worksheets[0].Cell("A1").SetValue("Data");
        wb.Security.SetModifyPassword("modify-secret");

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try
        {
            wb.SaveAs(path, ExcelFormat.Xlsb);

            // 重新打开：应识别到修改密码（只读）
            var reopened = Excel.Open(path, ExcelFormat.Xlsb);
            Assert.True(reopened.Security.HasModifyPassword);
            Assert.False(reopened.Security.HasModifyAccess);
            Assert.True(reopened.Security.IsReadOnly);
            Assert.False(reopened.Security.CanSave);

            // 提供修改密码：获得权限可保存
            var reopened2 = Excel.Open(path, new ExcelReadOptions { ModifyPassword = "modify-secret" });
            Assert.False(reopened2.Security.IsReadOnly);
            Assert.True(reopened2.Security.CanSave);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpenModifyProtected_Xlsb_ThenSaveAs_PreservesFileSharing()
    {
        var src = Sample("修改加密12.xlsb");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { ModifyPassword = "12" });
        Assert.False(wb.Security.IsReadOnly);
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try
        {
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            // 未改修改密码：透传保留原 fileSharing
            var reopened = Excel.Open(outPath, ExcelFormat.Xlsb);
            Assert.True(reopened.Security.HasModifyPassword);
            Assert.True(reopened.Security.IsReadOnly);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void OpenModifyProtected_Xlsb_RemoveModifyPassword_SaveAs_NoProtection()
    {
        var src = Sample("修改加密12.xlsb");
        Assert.True(File.Exists(src), $"Missing sample: {src}");

        var wb = Excel.Open(src, new ExcelReadOptions { ModifyPassword = "12" });
        wb.Security.RemoveModifyPassword();
        var outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try
        {
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            var reopened = Excel.Open(outPath, ExcelFormat.Xlsb);
            Assert.False(reopened.Security.HasModifyPassword);
            Assert.False(reopened.Security.IsReadOnly);
            Assert.True(reopened.Security.CanSave);
        }
        finally
        {
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }
}
