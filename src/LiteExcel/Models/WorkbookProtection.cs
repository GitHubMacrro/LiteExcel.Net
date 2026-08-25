using System;

namespace LiteExcel;

/// <summary>
/// 工作簿保护（<c>workbookProtection</c>）。锁定工作簿结构/窗口，可选密码。
/// 密码不对外暴露明文，仅以哈希形式写出（SHA-512 + salt + spinCount，与 fileSharing 同机制）。
/// </summary>
public sealed class WorkbookProtection
{
    private string? _password;

    // 从文件读取的密码哈希（用于 VerifyPassword；用户新设密码时为空）
    internal byte[]? HashValue { get; private set; }
    internal byte[]? SaltValue { get; private set; }
    internal string? AlgorithmName { get; private set; }
    internal int SpinCount { get; private set; } = 100000;

    /// <summary>是否启用保护（写出 <c>workbookProtection</c> 的前提） </summary>
    public bool Enabled { get; set; }

    /// <summary>是否锁定结构（禁止插入/删除/移动/隐藏/重命名工作表），默认 true </summary>
    public bool LockStructure { get; set; } = true;

    /// <summary>是否锁定窗口（禁止移动/调整工作簿窗口） </summary>
    public bool LockWindows { get; set; }

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

    /// <summary>是否应写出 workbookProtection 元素 </summary>
    internal bool IsActive => Enabled || HasPassword;

    internal void LoadHash(byte[]? hash, byte[]? salt, string? algo, int spin)
    {
        HashValue = hash;
        SaltValue = salt;
        AlgorithmName = algo;
        SpinCount = spin;
    }

    /// <summary>生成写入 workbookProtection 的密码属性片段（新设密码时现算 SHA-512 + 新 salt） </summary>
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
