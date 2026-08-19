using LiteExcel.Internal.Cfb;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel.Internal.Encryption;

/// <summary>
/// OOXML Agile Encryption 解密器（MS-OFFCRYPTO / ECMA-376）。
/// 用于解密带打开密码的 xlsx/xlsm/xlsb（OLE CFB 容器 + EncryptionInfo + EncryptedPackage）。
/// 算法：自定义迭代哈希派生密钥（SHA-1/256/384/512 + AES-CBC），payload 按 4096 字节分段解密。
/// 仅依赖 BCL（SHA*/Aes），net48 与 net8.0 均可用。
/// </summary>
internal static class AgileDecryptor
{
    // MS-OFFCRYPTO block keys
    private static readonly byte[] BlkVerifierHashInput = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
    private static readonly byte[] BlkEncVerifierHashValue = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
    private static readonly byte[] BlkEncKeyValue = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
    private static readonly byte[] BlkDataIntegrity1 = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
    private static readonly byte[] BlkDataIntegrity2 = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };

    /// <summary>
    /// 尝试用密码解密 CFB 加密工作簿。
    /// 成功返回解密后的 zip 流；密码错误抛 <see cref="LiteExcelException"/>。
    /// </summary>
    public static MemoryStream Decrypt(Stream cfbStream, string password)
    {
        var cfb = CfbFile.Open(cfbStream);
        var encInfo = cfb.GetStream("EncryptionInfo")
            ?? throw new LiteExcelException("文件已加密（带打开密码），但缺少 EncryptionInfo 流，无法解密。");
        var encPackage = cfb.GetStream("EncryptedPackage")
            ?? throw new LiteExcelException("文件已加密（带打开密码），但缺少 EncryptedPackage 流，无法解密。");

        var xml = ParseEncryptionInfo(encInfo);
        if (!TryParseAgile(xml, out var spinCount, out var encKeySalt, out var encVerifierInput,
                out var encVerifierHash, out var encKeyValue, out var hashAlgo, out var keyBits,
                out var keyDataSalt, out var keyDataHash, out var hashSize, out var encHmacKey, out var encHmacValue))
        {
            throw new LiteExcelException(
                "文件使用不支持的加密方式。当前仅支持 OOXML Agile 加密（Excel 2013 及以上默认）。");
        }

        var iterated = IteratedHash(encKeySalt, password, hashAlgo, spinCount);

        // 验证密码
        var k1 = DeriveKey(iterated, BlkVerifierHashInput, hashAlgo, keyBits);
        var k2 = DeriveKey(iterated, BlkEncVerifierHashValue, hashAlgo, keyBits);
        var verifierInput = DecryptCbc(encVerifierInput, k1, encKeySalt);
        // verifier 输入/哈希为 hashSize 字节，补零到 16 字节倍数 → 按 hashSize 截断再比对
        if (verifierInput.Length > hashSize)
            verifierInput = verifierInput.Take(hashSize).ToArray();
        var verifierHash = DecryptCbc(encVerifierHash, k2, encKeySalt);
        using (var h = CreateHash(hashAlgo))
        {
            var computed = h.ComputeHash(verifierInput);
            var expected = verifierHash.Take(computed.Length).ToArray();
            if (!computed.SequenceEqual(expected))
                throw new LiteExcelException("打开密码不正确，无法解密文件。");
        }

        // 解密密钥
        var k3 = DeriveKey(iterated, BlkEncKeyValue, hashAlgo, keyBits);
        var secretKey = DecryptCbc(encKeyValue, k3, encKeySalt);

        // 分段解密 payload
        long dataSize = BitConverter.ToInt64(encPackage, 0);
        var result = new MemoryStream();
        int seg = 0;
        int offset = 8;
        using (var h = CreateHash(keyDataHash))
        {
            while (offset < encPackage.Length)
            {
                int take = Math.Min(4096, encPackage.Length - offset);
                var iv = h.ComputeHash(keyDataSalt.Concat(BitConverter.GetBytes(seg)).ToArray());
                iv = iv.Take(16).ToArray();
                var dec = DecryptCbcSegment(encPackage, offset, take, secretKey, iv);
                result.Write(dec, 0, dec.Length);
                offset += take;
                seg++;
            }
        }

        // data integrity：解密 HMAC 密钥并校验密文（防止被篡改）
        if (encHmacKey is not null && encHmacValue is not null)
        {
            var iv1 = HashForKey(keyDataSalt, BlkDataIntegrity1);
            var iv2 = HashForKey(keyDataSalt, BlkDataIntegrity2);
            var hmacKey = DecryptCbc(encHmacKey, secretKey, iv1);
            var hmacExpected = DecryptCbc(encHmacValue, secretKey, iv2);
            using (var hmac = new HMACSHA512(hmacKey))
            {
                var computed = hmac.ComputeHash(encPackage);
                var expected = hmacExpected.Take(computed.Length).ToArray();
                if (!computed.SequenceEqual(expected))
                    throw new LiteExcelException("文件完整性校验失败（EncryptedPackage 可能被篡改）。");
            }
        }

        // 截断到真实明文长度（去掉末段 AES 零填充）
        if (dataSize > 0 && dataSize < result.Length)
        {
            var trimmed = new byte[dataSize];
            result.Position = 0;
            result.Read(trimmed, 0, (int)dataSize);
            result.SetLength(dataSize);
            result.Position = 0;
        }
        else
        {
            result.Position = 0;
        }
        return result;
    }

    // ── 解析 ──

    private static byte[] ParseEncryptionInfo(byte[] encInfo)
    {
        // 头 8 字节：major(2) minor(2) flags(4)；XML 从 offset 8 开始
        ushort major = BitConverter.ToUInt16(encInfo, 0);
        uint flags = BitConverter.ToUInt32(encInfo, 4);
        if ((major != 4 && major != 3) || (flags & 0x40) == 0)
            throw new LiteExcelException("文件使用不支持的加密方式（非 Agile Encryption）。");
        return encInfo.Skip(8).ToArray();
    }

    private static bool TryParseAgile(byte[] xmlBytes,
        out int spinCount, out byte[] encKeySalt, out byte[] encVerifierInput,
        out byte[] encVerifierHash, out byte[] encKeyValue, out string hashAlgo,
        out int keyBits, out byte[] keyDataSalt, out string keyDataHash,
        out int hashSize, out byte[]? encHmacKey, out byte[]? encHmacValue)
    {
        spinCount = 0;
        encKeySalt = Array.Empty<byte>();
        encVerifierInput = Array.Empty<byte>();
        encVerifierHash = Array.Empty<byte>();
        encKeyValue = Array.Empty<byte>();
        hashAlgo = "SHA512";
        keyBits = 256;
        keyDataSalt = Array.Empty<byte>();
        keyDataHash = "SHA512";
        hashSize = 64;
        encHmacKey = null;
        encHmacValue = null;

        try
        {
            var xml = Encoding.UTF8.GetString(xmlBytes).TrimStart('\uFEFF');
            var doc = XDocument.Parse(xml);
            XNamespace p = "http://schemas.microsoft.com/office/2006/keyEncryptor/password";
            XNamespace ns = "http://schemas.microsoft.com/office/2006/encryption";

            var keyEnc = doc.Descendants(p + "encryptedKey").FirstOrDefault();
            if (keyEnc is null) return false;

            spinCount = int.Parse((string)keyEnc.Attribute("spinCount") ?? "100000");
            encKeySalt = Convert.FromBase64String((string)keyEnc.Attribute("saltValue")!);
            encVerifierInput = Convert.FromBase64String((string)keyEnc.Attribute("encryptedVerifierHashInput")!);
            encVerifierHash = Convert.FromBase64String((string)keyEnc.Attribute("encryptedVerifierHashValue")!);
            encKeyValue = Convert.FromBase64String((string)keyEnc.Attribute("encryptedKeyValue")!);
            hashAlgo = (string)keyEnc.Attribute("hashAlgorithm") ?? "SHA512";
            keyBits = int.Parse((string)keyEnc.Attribute("keyBits") ?? "256");
            hashSize = int.Parse((string)keyEnc.Attribute("hashSize") ?? "64");

            var keyData = doc.Descendants(ns + "keyData").FirstOrDefault();
            keyDataSalt = keyData is null
                ? Array.Empty<byte>()
                : Convert.FromBase64String((string)keyData.Attribute("saltValue")!);
            keyDataHash = (string?)keyData?.Attribute("hashAlgorithm") ?? "SHA512";

            var dataIntegrity = doc.Descendants(ns + "dataIntegrity").FirstOrDefault();
            if (dataIntegrity is not null)
            {
                var hmacKeyAttr = dataIntegrity.Attribute("encryptedHmacKey");
                var hmacValueAttr = dataIntegrity.Attribute("encryptedHmacValue");
                if (hmacKeyAttr is not null && hmacValueAttr is not null)
                {
                    encHmacKey = Convert.FromBase64String(hmacKeyAttr.Value);
                    encHmacValue = Convert.FromBase64String(hmacValueAttr.Value);
                }
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>HMAC IV 派生：iv = SHA512(keyDataSalt + blockKey).Take(16)（与加密侧一致） </summary>
    private static byte[] HashForKey(byte[] salt, byte[] blockKey)
    {
        using var h = SHA512.Create();
        var combined = new byte[salt.Length + blockKey.Length];
        Array.Copy(salt, combined, salt.Length);
        Array.Copy(blockKey, 0, combined, salt.Length, blockKey.Length);
        return h.ComputeHash(combined).Take(16).ToArray();
    }

    // ── 密码派生 ──

    /// <summary>自定义迭代哈希：H = SHA(salt + pwd_UTF16LE)；for i: H = SHA(le32(i) + H) </summary>
    private static byte[] IteratedHash(byte[] salt, string password, string algo, int spin)
    {
        using var h = CreateHash(algo);
        var pwd = Encoding.Unicode.GetBytes(password);
        var input = new byte[salt.Length + pwd.Length];
        Array.Copy(salt, input, salt.Length);
        Array.Copy(pwd, 0, input, salt.Length, pwd.Length);
        var h0 = h.ComputeHash(input);
        for (int i = 0; i < spin; i++)
        {
            var it = BitConverter.GetBytes(i);
            var combined = new byte[it.Length + h0.Length];
            Array.Copy(it, combined, it.Length);
            Array.Copy(h0, 0, combined, it.Length, h0.Length);
            h0 = h.ComputeHash(combined);
        }
        return h0;
    }

    /// <summary>派生密钥：key = truncate(SHA(h + blockKey), keyBits/8) </summary>
    private static byte[] DeriveKey(byte[] iterated, byte[] blockKey, string algo, int keyBits)
    {
        using var h = CreateHash(algo);
        var combined = new byte[iterated.Length + blockKey.Length];
        Array.Copy(iterated, combined, iterated.Length);
        Array.Copy(blockKey, 0, combined, iterated.Length, blockKey.Length);
        var full = h.ComputeHash(combined);
        return full.Take(keyBits / 8).ToArray();
    }

    // ── AES-CBC ──

    private static byte[] DecryptCbc(byte[] data, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        using var dec = aes.CreateDecryptor();
        return dec.TransformFinalBlock(data, 0, data.Length);
    }

    private static byte[] DecryptCbcSegment(byte[] package, int offset, int length, byte[] key, byte[] iv)
    {
        using var aes = Aes.Create();
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        using var dec = aes.CreateDecryptor();
        return dec.TransformFinalBlock(package, offset, length);
    }

    private static HashAlgorithm CreateHash(string algo) => algo switch
    {
        "SHA512" => SHA512.Create(),
        "SHA384" => SHA384.Create(),
        "SHA256" => SHA256.Create(),
        "SHA1" => SHA1.Create(),
        _ => throw new LiteExcelException($"不支持的哈希算法：{algo}"),
    };
}
