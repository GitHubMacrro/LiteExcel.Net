using System;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace LiteExcel.Internal.Encryption;

/// <summary>
/// workbook.xml 中 &lt;fileSharing&gt; 元素的信息（修改密码 / 写保护）。
/// 记录哈希参数用于验证用户提供的修改密码；密码明文绝不保存在此。
/// </summary>
internal sealed class FileSharingInfo
{
    public FileSharingInfo(byte[] hashValue, byte[]? saltValue, string? algorithmName,
        int spinCount, bool readOnlyRecommended)
    {
        HashValue = hashValue;
        SaltValue = saltValue;
        AlgorithmName = algorithmName;
        SpinCount = spinCount;
        ReadOnlyRecommended = readOnlyRecommended;
    }

    public byte[] HashValue { get; }
    public byte[]? SaltValue { get; }
    public string? AlgorithmName { get; }
    public int SpinCount { get; }
    public bool ReadOnlyRecommended { get; }

    /// <summary>
    /// 计算 fileSharing 的密码哈希（ISO/IEC 29500 盐化哈希）。
    /// 与 <see cref="VerifyPassword"/> 对称：UTF-16LE 密码 + salt 迭代哈希，截取 hashValue 长度。
    /// </summary>
    public static byte[] ComputeHash(string password, byte[] salt, string algorithmName = "SHA-512", int spinCount = 100000)
    {
        HashAlgorithm NewHash() => algorithmName switch
        {
            "SHA-512" => SHA512.Create(),
            "SHA-256" => SHA256.Create(),
            "SHA-1" => SHA1.Create(),
            "SHA-384" => SHA384.Create(),
            _ => throw new LiteExcelException($"不支持的哈希算法：{algorithmName}"),
        };

        using var h = NewHash();
        var pwd = Encoding.Unicode.GetBytes(password);
        var input = new byte[salt.Length + pwd.Length];
        Array.Copy(salt, input, salt.Length);
        Array.Copy(pwd, 0, input, salt.Length, pwd.Length);
        var cur = h.ComputeHash(input);
        for (int i = 0; i < spinCount; i++)
        {
            var it = BitConverter.GetBytes(i);
            var combined = new byte[it.Length + cur.Length];
            Array.Copy(it, combined, it.Length);
            Array.Copy(cur, 0, combined, it.Length, cur.Length);
            cur = h.ComputeHash(combined);
        }
        return cur;
    }

    /// <summary>
    /// 校验密码是否正确。
    /// 采用 ISO/IEC 29500 fileSharing 的盐化哈希：UTF-16LE 密码 + salt 迭代哈希。
    /// 若参数缺失或算法未知则返回 false（不误授权）。
    /// </summary>
    public bool VerifyPassword(string password)
    {
        if (SaltValue is null || string.IsNullOrEmpty(AlgorithmName) || SpinCount <= 0)
            return false;

        try
        {
            var computed = ComputeHash(password, SaltValue, AlgorithmName, SpinCount);
            var expected = computed.Take(HashValue.Length).ToArray();
            return expected.SequenceEqual(HashValue);
        }
        catch (LiteExcelException)
        {
            return false;
        }
    }
}
