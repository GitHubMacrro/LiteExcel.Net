using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace LiteExcel.Internal.Encryption;

/// <summary>
/// OOXML Agile Encryption 加密器（与 <see cref="AgileDecryptor"/> 对称）。
/// 把普通 zip 包封装为带打开密码的 OLE CFB 工作簿（EncryptionInfo + EncryptedPackage + Version 流）。
/// 算法：自定义迭代哈希派生密钥 + AES-256-CBC 分段加密 payload。
/// </summary>
internal static class OoxmlEncryptor
{
    private static readonly byte[] BlkVerifierHashInput = { 0xFE, 0xA7, 0xD2, 0x76, 0x3B, 0x4B, 0x9E, 0x79 };
    private static readonly byte[] BlkEncVerifierHashValue = { 0xD7, 0xAA, 0x0F, 0x6D, 0x30, 0x61, 0x34, 0x4E };
    private static readonly byte[] BlkEncKeyValue = { 0x14, 0x6E, 0x0B, 0xE7, 0xAB, 0xAC, 0xD0, 0xD6 };
    private static readonly byte[] BlkDataIntegrity1 = { 0x5F, 0xB2, 0xAD, 0x01, 0x0C, 0xB9, 0xE1, 0xF6 };
    private static readonly byte[] BlkDataIntegrity2 = { 0xA0, 0x67, 0x7F, 0x02, 0xB2, 0x2C, 0x84, 0x33 };

    private const int SaltSize = 16;
    private const int BlockSize = 16;
    private const int KeyBits = 256;
    private const int HashSize = 64;
    private const int SpinCount = 100000;
    private const int SegmentLength = 4096;

    /// <summary>加密 zip 包，返回完整 CFB 文件字节（含 EncryptionInfo/EncryptedPackage/Version 流） </summary>
    public static byte[] Encrypt(byte[] plainZip, string password)
    {
        // 1. 派生密钥
        var encKeySalt = RandomBytes(SaltSize);
        var iterated = IteratedHash(encKeySalt, password);

        var k1 = DeriveKey(iterated, BlkVerifierHashInput);
        var k2 = DeriveKey(iterated, BlkEncVerifierHashValue);
        var k3 = DeriveKey(iterated, BlkEncKeyValue);

        // 2. verifier + secret key
        var verifierInput = RandomBytes(SaltSize);
        using (var h = SHA512.Create())
        {
            var verifierHash = h.ComputeHash(verifierInput);

            // 3. 加密各段
            var encVerifierInput = EncryptCbc(verifierInput, k1, encKeySalt);
            var encVerifierHash = EncryptCbc(verifierHash, k2, encKeySalt);
            var secretKey = RandomBytes(KeyBits / 8);
            var encKeyValue = EncryptCbc(secretKey, k3, encKeySalt);

            // 4. 加密 payload（分段）
            var keyDataSalt = RandomBytes(SaltSize);
            var encryptedPayload = EncryptPayload(plainZip, secretKey, keyDataSalt);

            // 5. data integrity（HMAC）
            var hmacSalt = RandomBytes(HashSize);
            var iv1 = HashForKey(keyDataSalt, BlkDataIntegrity1);
            var iv2 = HashForKey(keyDataSalt, BlkDataIntegrity2);
            var encHmacKey = EncryptCbc(hmacSalt, secretKey, iv1);
            byte[] hmacValue;
            using (var hmac = new HMACSHA512(hmacSalt))
            {
                hmacValue = hmac.ComputeHash(encryptedPayload);
            }
            var encHmacValue = EncryptCbc(hmacValue, secretKey, iv2);

            // 6. 组装 EncryptionInfo XML（带 Agile 版本头：major=4 minor=4 flags=0x40）
            var encInfoXml = BuildEncryptionInfoXml(encKeySalt, keyDataSalt, encVerifierInput,
                encVerifierHash, encKeyValue, encHmacKey, encHmacValue);
            var encInfoBytes = new byte[8 + encInfoXml.Length];
            encInfoBytes[0] = 0x04; encInfoBytes[1] = 0x00;
            encInfoBytes[2] = 0x04; encInfoBytes[3] = 0x00;
            encInfoBytes[4] = 0x40; encInfoBytes[5] = 0x00;
            encInfoBytes[6] = 0x00; encInfoBytes[7] = 0x00;
            Encoding.UTF8.GetBytes(encInfoXml, 0, encInfoXml.Length, encInfoBytes, 8);

            // 7. 组装 CFB：EncryptedPackage + EncryptionInfo + DataSpaces 骨架
            return Cfb.EncryptedCfbWriter.Build(encryptedPayload, encInfoBytes);
        }
    }

    private static byte[] EncryptPayload(byte[] plain, byte[] secretKey, byte[] keyDataSalt)
    {
        // 结构：前 8 字节 = 明文长度（小端 int64），然后分段加密
        using var outMs = new MemoryStream();
        var sizeBytes = BitConverter.GetBytes((long)plain.Length);
        outMs.Write(sizeBytes, 0, 8);

        int offset = 0;
        int seg = 0;
        using (var h = SHA512.Create())
        {
            while (offset < plain.Length)
            {
                int take = Math.Min(SegmentLength, plain.Length - offset);
                var segment = new byte[take];
                Array.Copy(plain, offset, segment, 0, take);
                var iv = h.ComputeHash(keyDataSalt.Concat(BitConverter.GetBytes(seg)).ToArray());
                iv = iv.Take(16).ToArray();
                var enc = EncryptCbc(segment, secretKey, iv);
                outMs.Write(enc, 0, enc.Length);
                offset += take;
                seg++;
            }
        }
        return outMs.ToArray();
    }

    private static string BuildEncryptionInfoXml(byte[] encKeySalt, byte[] keyDataSalt,
        byte[] encVerifierInput, byte[] encVerifierHash, byte[] encKeyValue,
        byte[] encHmacKey, byte[] encHmacValue)
    {
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               $"<encryption xmlns=\"http://schemas.microsoft.com/office/2006/encryption\" " +
               $"xmlns:p=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\" " +
               $"xmlns:c=\"http://schemas.microsoft.com/office/2006/keyEncryptor/certificate\">" +
               $"<keyData saltSize=\"{SaltSize}\" blockSize=\"{BlockSize}\" keyBits=\"{KeyBits}\" hashSize=\"{HashSize}\" " +
               $"cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" " +
               $"saltValue=\"{Convert.ToBase64String(keyDataSalt)}\" />" +
               $"<dataIntegrity encryptedHmacKey=\"{Convert.ToBase64String(encHmacKey)}\" " +
               $"encryptedHmacValue=\"{Convert.ToBase64String(encHmacValue)}\" />" +
               $"<keyEncryptors><keyEncryptor uri=\"http://schemas.microsoft.com/office/2006/keyEncryptor/password\">" +
               $"<p:encryptedKey spinCount=\"{SpinCount}\" saltSize=\"{SaltSize}\" blockSize=\"{BlockSize}\" keyBits=\"{KeyBits}\" hashSize=\"{HashSize}\" " +
               $"cipherAlgorithm=\"AES\" cipherChaining=\"ChainingModeCBC\" hashAlgorithm=\"SHA512\" " +
               $"saltValue=\"{Convert.ToBase64String(encKeySalt)}\" " +
               $"encryptedVerifierHashInput=\"{Convert.ToBase64String(encVerifierInput)}\" " +
               $"encryptedVerifierHashValue=\"{Convert.ToBase64String(encVerifierHash)}\" " +
               $"encryptedKeyValue=\"{Convert.ToBase64String(encKeyValue)}\" /></keyEncryptor></keyEncryptors></encryption>";
    }

    // ── 密码派生（与 AgileDecryptor 对称） ──

    private static byte[] IteratedHash(byte[] salt, string password)
    {
        using var h = SHA512.Create();
        var pwd = Encoding.Unicode.GetBytes(password);
        var input = new byte[salt.Length + pwd.Length];
        Array.Copy(salt, input, salt.Length);
        Array.Copy(pwd, 0, input, salt.Length, pwd.Length);
        var cur = h.ComputeHash(input);
        for (int i = 0; i < SpinCount; i++)
        {
            var it = BitConverter.GetBytes(i);
            var combined = new byte[it.Length + cur.Length];
            Array.Copy(it, combined, it.Length);
            Array.Copy(cur, 0, combined, it.Length, cur.Length);
            cur = h.ComputeHash(combined);
        }
        return cur;
    }

    private static byte[] DeriveKey(byte[] iterated, byte[] blockKey)
    {
        using var h = SHA512.Create();
        var combined = new byte[iterated.Length + blockKey.Length];
        Array.Copy(iterated, combined, iterated.Length);
        Array.Copy(blockKey, 0, combined, iterated.Length, blockKey.Length);
        return h.ComputeHash(combined).Take(KeyBits / 8).ToArray();
    }

    private static byte[] HashForKey(byte[] salt, byte[] blockKey)
    {
        using var h = SHA512.Create();
        var combined = new byte[salt.Length + blockKey.Length];
        Array.Copy(salt, combined, salt.Length);
        Array.Copy(blockKey, 0, combined, salt.Length, blockKey.Length);
        return h.ComputeHash(combined).Take(16).ToArray();
    }

    private static byte[] EncryptCbc(byte[] data, byte[] key, byte[] iv)
    {
        // OOXML Agile 用零填充（补 0x00 到 blockSize 倍数），非 PKCS7
        int paddedLen = ((data.Length + BlockSize - 1) / BlockSize) * BlockSize;
        var padded = new byte[paddedLen];
        Array.Copy(data, padded, data.Length);
        using var aes = Aes.Create();
        aes.Key = key;
        aes.IV = iv;
        aes.Mode = CipherMode.CBC;
        aes.Padding = PaddingMode.None;
        using var enc = aes.CreateEncryptor();
        return enc.TransformFinalBlock(padded, 0, padded.Length);
    }

    private static byte[] RandomBytes(int count)
    {
        var buf = new byte[count];
        using var rng = RandomNumberGenerator.Create();
        rng.GetBytes(buf);
        return buf;
    }
}
