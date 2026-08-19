using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Phase 1：文件级安全状态模型测试。
/// 覆盖 WorkbookSecurity 状态矩阵、密码管理方法、只读工作簿保存拦截。
/// </summary>
public class SecurityStateTests
{
    // ── 无密码场景 ──

    [Fact]
    public void NewWorkbook_NoPassword_CanSave()
    {
        var wb = Excel.Create();
        Assert.False(wb.Security.HasOpenPassword);
        Assert.False(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
        Assert.True(wb.Security.HasModifyAccess);
    }

    [Fact]
    public void OpenPlainWorkbook_NoPassword_CanSave()
    {
        using var ms = new MemoryStream();
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Save(ms, ExcelFormat.Xlsx);

        ms.Position = 0;
        var opened = Excel.Open(ms, ExcelFormat.Xlsx);
        Assert.False(opened.Security.HasOpenPassword);
        Assert.False(opened.Security.HasModifyPassword);
        Assert.True(opened.Security.CanSave);
        Assert.False(opened.Security.IsReadOnly);
    }

    // ── 打开密码 ──

    [Fact]
    public void SetOpenPassword_UpdatesState()
    {
        var wb = Excel.Create();
        wb.Security.SetOpenPassword("secret");
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.CanSave);
        Assert.False(wb.Security.IsReadOnly);
    }

    [Fact]
    public void SetOpenPassword_EmptyString_RemovesPassword()
    {
        var wb = Excel.Create();
        wb.Security.SetOpenPassword("secret");
        Assert.True(wb.Security.HasOpenPassword);

        wb.Security.SetOpenPassword("");
        Assert.False(wb.Security.HasOpenPassword);
    }

    [Fact]
    public void RemoveOpenPassword_UpdatesState()
    {
        var wb = Excel.Create();
        wb.Security.SetOpenPassword("secret");
        wb.Security.RemoveOpenPassword();
        Assert.False(wb.Security.HasOpenPassword);
    }

    // ── 修改密码 ──

    [Fact]
    public void SetModifyPassword_WithoutAccess_IsReadOnly_NoModifyAccess()
    {
        // 模拟打开时识别到修改密码但未授权
        var wb = Excel.Create();
        wb.Security.Initialize(null, fileHasModifyProtection: true);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.HasModifyAccess);
        Assert.True(wb.Security.IsReadOnly);
        Assert.False(wb.Security.CanSave);
    }

    [Fact]
    public void GrantModifyAccess_EnablesSave()
    {
        var wb = Excel.Create();
        wb.Security.Initialize(null, fileHasModifyProtection: true);
        wb.Security.GrantModifyAccess(readOnlyRecommended: true);
        Assert.True(wb.Security.HasModifyAccess);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
    }

    [Fact]
    public void SetModifyPassword_GrantsAccess()
    {
        var wb = Excel.Create();
        wb.Security.SetModifyPassword("modify-secret");
        Assert.True(wb.Security.HasModifyAccess);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
        Assert.True(wb.Security.ReadOnlyRecommended);
    }

    [Fact]
    public void RemoveModifyPassword_WithoutAccess_Throws()
    {
        var wb = Excel.Create();
        wb.Security.Initialize(null, fileHasModifyProtection: true);
        // 未授权时移除必须失败，防止绕过保护
        Assert.Throws<LiteExcelException>(() => wb.Security.RemoveModifyPassword());
    }

    [Fact]
    public void RemoveModifyPassword_WithAccess_Succeeds()
    {
        var wb = Excel.Create();
        wb.Security.Initialize(null, fileHasModifyProtection: true);
        wb.Security.GrantModifyAccess(readOnlyRecommended: true);
        wb.Security.RemoveModifyPassword();
        Assert.False(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
    }

    [Fact]
    public void SetModifyPassword_WithoutAccess_Throws()
    {
        var wb = Excel.Create();
        wb.Security.Initialize(null, fileHasModifyProtection: true);
        // 未授权时设置/替换修改密码必须失败，防止未授权剥离或替换写保护
        Assert.Throws<LiteExcelException>(() => wb.Security.SetModifyPassword("attacker-set"));
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.HasModifyAccess);
        Assert.False(wb.Security.CanSave);
    }

    [Fact]
    public void ClearAll_WithoutAccess_Throws()
    {
        var wb = Excel.Create();
        wb.Security.Initialize("open-secret", fileHasModifyProtection: true);
        // 未授权时清空全部密码必须失败（含剥离写保护）
        Assert.Throws<LiteExcelException>(() => wb.Security.ClearAll());
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.CanSave);
    }

    [Fact]
    public void ClearAll_WithAccess_Succeeds()
    {
        var wb = Excel.Create();
        wb.Security.Initialize("open-secret", fileHasModifyProtection: true);
        wb.Security.GrantModifyAccess(readOnlyRecommended: true);
        wb.Security.ClearAll();
        Assert.False(wb.Security.HasOpenPassword);
        Assert.False(wb.Security.HasModifyPassword);
        Assert.True(wb.Security.CanSave);
    }

    // ── 双密码 ──

    [Fact]
    public void DualPassword_OnlyOpenGranted_IsReadOnly()
    {
        var wb = Excel.Create();
        // 打开密码已提供（解密成功），但修改密码未提供 -> 只读
        wb.Security.Initialize("open-secret", fileHasModifyProtection: true);
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.HasModifyAccess);
        Assert.True(wb.Security.IsReadOnly);
        Assert.False(wb.Security.CanSave);
    }

    [Fact]
    public void DualPassword_BothGranted_CanSave()
    {
        var wb = Excel.Create();
        wb.Security.Initialize("open-secret", fileHasModifyProtection: true);
        wb.Security.GrantModifyAccess(readOnlyRecommended: true);
        Assert.True(wb.Security.HasOpenPassword);
        Assert.True(wb.Security.HasModifyPassword);
        Assert.False(wb.Security.IsReadOnly);
        Assert.True(wb.Security.CanSave);
    }

    // ── 保存拦截 ──

    [Fact]
    public void Save_ReadOnlyWorkbook_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Security.Initialize(null, fileHasModifyProtection: true); // 只读

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(path));
            Assert.Contains("只读", ex.Message);
            Assert.DoesNotContain("modify-secret", ex.Message); // 密码不出现在异常
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SaveAs_ReadOnlyWorkbook_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Security.Initialize(null, fileHasModifyProtection: true);

        var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(Path.GetTempPath() + "t.xlsx", ExcelFormat.Xlsx));
        Assert.Contains("只读", ex.Message);
    }

    [Fact]
    public void Save_ToStream_ReadOnlyWorkbook_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Security.Initialize(null, fileHasModifyProtection: true);

        using var ms = new MemoryStream();
        var ex = Assert.Throws<LiteExcelException>(() => wb.Save(ms, ExcelFormat.Xlsx));
        Assert.Contains("只读", ex.Message);
    }

    // ── 密码不出现在输出 ──

    [Fact]
    public void ErrorMessages_NeverContainPassword()
    {
        const string openPwd = "OpenSecret123!";
        const string modifyPwd = "ModifySecret456!";

        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        // 打开密码从文件解密得到；修改密码由用户显式设置
        wb.Security.Initialize(openPwd);
        wb.Security.SetModifyPassword(modifyPwd);

        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(path, ExcelFormat.Csv));
            Assert.DoesNotContain(openPwd, ex.Message);
            Assert.DoesNotContain(modifyPwd, ex.Message);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    // ── csv/xls 不支持密码写出 ──

    [Fact]
    public void SaveAs_Csv_WithPassword_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Security.SetOpenPassword("secret");

        var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(Path.GetTempPath() + "t.csv", ExcelFormat.Csv));
        Assert.Contains("不支持", ex.Message);
    }

    [Fact]
    public void SaveAs_Xls_WithPassword_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("S1");
        wb.Security.SetModifyPassword("modify-secret");

        var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(Path.GetTempPath() + "t.xls", ExcelFormat.Xls));
        Assert.Contains("不支持", ex.Message);
    }
}
