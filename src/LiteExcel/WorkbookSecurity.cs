using System;

namespace LiteExcel;

/// <summary>
/// 工作簿文件级安全状态。
/// 管理两类文件级密码：打开密码（文件加密）与修改密码（写保护）。
/// 密码本体仅存储于本对象内部，不对外暴露；错误消息与序列化均不含密码明文。
/// </summary>
public sealed class WorkbookSecurity
{
    private string? _openPassword;
    private string? _modifyPassword;
    private bool _fileHasModifyProtection;

    /// <summary>用户是否主动设置/移除了修改密码（决定保存时是否透传原 fileSharing） </summary>
    internal bool ModifyPasswordTouched { get; private set; }

    internal WorkbookSecurity()
    {
        // 无密码的工作簿天然拥有修改权限
        HasModifyAccess = true;
    }

    /// <summary>文件是否有打开密码（文件加密） </summary>
    public bool HasOpenPassword => !string.IsNullOrEmpty(_openPassword);

    /// <summary>文件是否有修改密码（写保护） </summary>
    public bool HasModifyPassword => _fileHasModifyProtection || !string.IsNullOrEmpty(_modifyPassword);

    /// <summary>当前是否已获得修改权限（提供了正确的修改密码） </summary>
    public bool HasModifyAccess { get; internal set; }

    /// <summary>当前工作簿是否只读（存在修改密码但未获得修改权限） </summary>
    public bool IsReadOnly => HasModifyPassword && !HasModifyAccess;

    /// <summary>当前是否允许保存（只读工作簿不可保存，除非获得修改权限） </summary>
    public bool CanSave => !IsReadOnly;

    /// <summary>设置打开密码。覆盖旧值；空/空白字符串视为移除 </summary>
    public void SetOpenPassword(string password)
    {
        if (string.IsNullOrEmpty(password))
        {
            _openPassword = null;
            return;
        }
        _openPassword = password;
    }

    /// <summary>移除打开密码（下次保存为无打开密码文件） </summary>
    public void RemoveOpenPassword() => _openPassword = null;

    /// <summary>
    /// 设置修改密码。覆盖旧值；空/空白字符串视为移除。
    /// 修改保护存在时要求已获得修改权限（防止未授权剥离/替换写保护）。
    /// </summary>
    public void SetModifyPassword(string password, bool readOnlyRecommended = true)
    {
        EnsureModifyAccess();
        ModifyPasswordTouched = true;
        _fileHasModifyProtection = false;
        if (string.IsNullOrEmpty(password))
        {
            _modifyPassword = null;
            HasModifyAccess = true; // 无修改密码则不再受限
            return;
        }
        _modifyPassword = password;
        HasModifyAccess = true;
        ReadOnlyRecommended = readOnlyRecommended;
    }

    /// <summary>移除修改密码（下次保存为无修改保护文件）。要求已获得修改权限 </summary>
    public void RemoveModifyPassword()
    {
        EnsureModifyAccess();
        ModifyPasswordTouched = true;
        _modifyPassword = null;
        _fileHasModifyProtection = false;
        HasModifyAccess = true;
        ReadOnlyRecommended = false;
    }

    private void EnsureModifyAccess()
    {
        if (!HasModifyAccess)
            throw new LiteExcelException(
                "当前工作簿未获得修改权限（未提供正确的修改密码），无法修改写保护设置。");
    }

    /// <summary>是否建议只读打开（随修改密码设置，Excel 打开时提示） </summary>
    public bool ReadOnlyRecommended { get; private set; }

    /// <summary>清空全部文件级密码。修改保护存在时要求已获得修改权限 </summary>
    public void ClearAll()
    {
        EnsureModifyAccess();
        ModifyPasswordTouched = true;
        _openPassword = null;
        _modifyPassword = null;
        _fileHasModifyProtection = false;
        HasModifyAccess = true;
        ReadOnlyRecommended = false;
    }

    // ── 内部访问（仅供保存管线使用，不对外暴露密码明文） ──

    internal string? GetOpenPassword() => _openPassword;

    internal string? GetModifyPassword() => _modifyPassword;

    /// <summary>打开文件时用读取到的真实密码状态初始化（内部调用，不暴露明文） </summary>
    /// <param name="openPassword">文件的实际打开密码（解密成功即已提供正确密码）；无则 null</param>
    /// <param name="fileHasModifyProtection">文件是否设置了修改密码（写保护）</param>
    /// <param name="readOnlyRecommended">是否建议只读打开</param>
    internal void Initialize(string? openPassword, bool fileHasModifyProtection = false, bool readOnlyRecommended = false)
    {
        _openPassword = openPassword;
        _modifyPassword = null;
        _fileHasModifyProtection = fileHasModifyProtection;
        ModifyPasswordTouched = false;
        // 打开时默认无修改权限：文件有修改保护则只读；无保护则拥有全部权限
        HasModifyAccess = !fileHasModifyProtection;
        ReadOnlyRecommended = readOnlyRecommended;
    }

    /// <summary>授予修改权限（打开时提供了正确修改密码后调用） </summary>
    internal void GrantModifyAccess(bool readOnlyRecommended)
    {
        HasModifyAccess = true;
        ReadOnlyRecommended = readOnlyRecommended;
    }
}
