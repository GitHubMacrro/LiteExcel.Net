using System;

namespace LiteExcel.Internal;

/// <summary>
/// BIFF 系列（xls / xlsb）共用的位级解码助手。
/// </summary>
internal static class BiffShared
{
    /// <summary>解码 BIFF 的 RK 压缩数值（整数或舍入 double） </summary>
    public static double DecodeRk(int rk)
    {
        bool fInt = (rk & 0x02) != 0;
        bool fX100 = (rk & 0x01) != 0;
        double val;
        if (fInt)
        {
            val = rk >> 2;
        }
        else
        {
            long bits = ((long)(rk & 0xFFFFFFFC)) << 32;
            val = BitConverter.Int64BitsToDouble(bits);
        }
        if (fX100) val /= 100.0;
        return val;
    }

    /// <summary>BIFF 错误码 → 文本表示 </summary>
    public static string ErrorCode(byte code) => code switch
    {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#ERR",
    };
}
