using System;

namespace LiteExcel;

/// <summary>
/// 工作表保护（<c>sheetProtection</c>）。控制受保护工作表中允许/禁止的操作，可选密码。
/// 密码不对外暴露明文，仅以哈希形式写出（SHA-512 + salt + spinCount，与 fileSharing 同机制）。
/// </summary>
public sealed class SheetProtection
{
    private string? _password;

    // 从文件读取的密码哈希（用于 VerifyPassword；用户新设密码时为空）
    internal byte[]? HashValue { get; private set; }
    internal byte[]? SaltValue { get; private set; }
    internal string? AlgorithmName { get; private set; }
    internal int SpinCount { get; private set; } = 100000;

    /// <summary>是否启用保护（写出 <c>sheetProtection</c> 的前提） </summary>
    public bool Enabled { get; set; }

    /// <summary>是否允许选定被锁定的单元格（默认 true） </summary>
    public bool SelectLockedCells { get; set; } = true;

    /// <summary>是否允许选定未锁定的单元格（默认 true） </summary>
    public bool SelectUnlockedCells { get; set; } = true;

    public bool FormatCells { get; set; }
    public bool FormatColumns { get; set; }
    public bool FormatRows { get; set; }
    public bool InsertColumns { get; set; }
    public bool InsertRows { get; set; }
    public bool InsertHyperlinks { get; set; }
    public bool DeleteColumns { get; set; }
    public bool DeleteRows { get; set; }
    public bool Sort { get; set; }
    public bool AutoFilter { get; set; }
    public bool PivotTables { get; set; }

    /// <summary>是否允许编辑对象（图形/图片），默认 true（与 Excel 默认一致） </summary>
    public bool Objects { get; set; } = true;

    /// <summary>是否允许编辑方案，默认 true（与 Excel 默认一致） </summary>
    public bool Scenarios { get; set; } = true;

    /// <summary>是否已设置密码 </summary>
    public bool HasPassword => !string.IsNullOrEmpty(_password);

    /// <summary>设置保护密码。null/空白视为移除 </summary>
    public void SetPassword(string? password)
    {
        _password = string.IsNullOrEmpty(password) ? null : password;
        HashValue = null;
        SaltValue = null;
    }

    /// <summary>移除保护密码 </summary>
    public void RemovePassword() => SetPassword(null);

    /// <summary>验证密码是否正确。仅对从文件读取的哈希有效；新设密码在写出前恒 false </summary>
    public bool VerifyPassword(string password)
    {
        if (string.IsNullOrEmpty(password) || HashValue is null || SaltValue is null
            || string.IsNullOrEmpty(AlgorithmName))
            return false;
        return new Internal.Encryption.FileSharingInfo(HashValue, SaltValue, AlgorithmName, SpinCount, false)
            .VerifyPassword(password);
    }

    /// <summary>是否应写出 sheetProtection 元素 </summary>
    internal bool IsActive => Enabled || HasPassword;

    internal void LoadHash(byte[]? hash, byte[]? salt, string? algo, int spin)
    {
        HashValue = hash;
        SaltValue = salt;
        AlgorithmName = algo;
        SpinCount = spin;
    }

    /// <summary>生成写入 sheetProtection 的密码属性片段（新设密码时现算 SHA-512 + 新 salt） </summary>
    internal string WriteHashAttributes()
    {
        if (string.IsNullOrEmpty(_password)) return "";
        var salt = new byte[16];
        using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
            rng.GetBytes(salt);
        var hash = Internal.Encryption.FileSharingInfo.ComputeHash(_password, salt, "SHA-512", 100000);
        return $" algorithmName=\"SHA-512\" hashValue=\"{Convert.ToBase64String(hash)}\" " +
               $"saltValue=\"{Convert.ToBase64String(salt)}\" spinCount=\"100000\"";
    }
}
