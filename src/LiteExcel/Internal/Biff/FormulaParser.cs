using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// BIFF8（xls）与 BIFF12（xlsb）公式 RPN → A1 文本解析器。
/// 采用后序求值：维护操作数栈，运算符/函数弹出操作数并压回表达式文本。
/// 解析不了（数组公式、3D 引用、名称等）返回 null，调用方降级为仅缓存结果值。
/// </summary>
internal static class FormulaParser
{
    /// <summary>解析 RPN 为 A1 文本（不带前导 '='）。失败返回 null。</summary>
    public static string? Parse(byte[] rpn, bool biff12)
    {
        var stack = new List<string>();
        int pos = 0;
        int len = rpn.Length;

        while (pos < len)
        {
            int ptg = rpn[pos++];
            int t = Token(ptg);

            switch (t)
            {
                // 二元/一元运算符（无操作数）
                case 0x03: if (!BinOp(stack, "+")) return null; break;
                case 0x04: if (!BinOp(stack, "-")) return null; break;
                case 0x05: if (!BinOp(stack, "*")) return null; break;
                case 0x06: if (!BinOp(stack, "/")) return null; break;
                case 0x07: if (!BinOp(stack, "^")) return null; break;
                case 0x08: if (!BinOp(stack, "&")) return null; break;
                case 0x09: if (!BinOp(stack, "<")) return null; break;
                case 0x0A: if (!BinOp(stack, "<=")) return null; break;
                case 0x0B: if (!BinOp(stack, "=")) return null; break;
                case 0x0C: if (!BinOp(stack, ">=")) return null; break;
                case 0x0D: if (!BinOp(stack, ">")) return null; break;
                case 0x0E: if (!BinOp(stack, "<>")) return null; break;
                case 0x0F: if (!BinOp(stack, " ")) return null; break;
                case 0x10: if (!BinOp(stack, ",")) return null; break;
                case 0x11: if (!BinOp(stack, ":")) return null; break;
                case 0x12: if (!UnOp(stack, "+")) return null; break;
                case 0x13: if (!UnOp(stack, "-")) return null; break;
                case 0x14:
                    if (stack.Count < 1) return null;
                    stack[stack.Count - 1] = stack[stack.Count - 1] + "%";
                    break;
                case 0x15:
                    if (stack.Count < 1) return null;
                    stack[stack.Count - 1] = "(" + stack[stack.Count - 1] + ")";
                    break;
                case 0x16: // MissArg
                    stack.Add("");
                    break;

                case 0x17: // Str
                {
                    var s = ReadString(rpn, ref pos, biff12);
                    if (s is null) return null;
                    stack.Add("\"" + s + "\"");
                    break;
                }
                case 0x1C: // Err
                    if (pos >= len) return null;
                    stack.Add(ErrorText(rpn[pos++]));
                    break;
                case 0x1D: // Bool
                    if (pos >= len) return null;
                    stack.Add(rpn[pos++] != 0 ? "TRUE" : "FALSE");
                    break;
                case 0x1E: // Int
                    if (pos + 2 > len) return null;
                    stack.Add(((int)ReadU16(rpn, pos)).ToString(CultureInfo.InvariantCulture));
                    pos += 2;
                    break;
                case 0x1F: // Num
                    if (pos + 8 > len) return null;
                    double num = BitConverter.ToDouble(rpn, pos);
                    stack.Add(num.ToString("G15", CultureInfo.InvariantCulture));
                    pos += 8;
                    break;

                case 0x21: // Func（参数个数取自内置表）
                {
                    if (pos + 2 > len) return null;
                    int iftab = ReadU16(rpn, pos);
                    pos += 2;
                    var name = FormulaFtab.Names[iftab];
                    if (name is null) return null;
                    int argc = FormulaFtab.ArgcFor(iftab);
                    if (!ApplyFunction(stack, name, argc)) return null;
                    break;
                }
                case 0x22: // FuncVar
                {
                    if (pos + 3 > len) return null;
                    int argc = rpn[pos++];
                    int iftab = ReadU16(rpn, pos);
                    pos += 2;
                    var name = FormulaFtab.Names[iftab];
                    if (name is null) return null;
                    if (!ApplyFunction(stack, name, argc & 0x7F)) return null;
                    break;
                }

                case 0x24: // Ref
                case 0x2C: // RefN
                {
                    int rw = ReadRow(rpn, ref pos, biff12);
                    if (rw < 0) return null;
                    int col = ReadCol(rpn, ref pos);
                    if (col < 0) return null;
                    stack.Add(RefText(rw, col));
                    break;
                }
                case 0x25: // Area
                case 0x2D: // AreaN
                {
                    int rw1 = ReadRow(rpn, ref pos, biff12);
                    int rw2 = ReadRow(rpn, ref pos, biff12);
                    int c1 = ReadCol(rpn, ref pos);
                    int c2 = ReadCol(rpn, ref pos);
                    if (rw1 < 0 || rw2 < 0 || c1 < 0 || c2 < 0) return null;
                    stack.Add(RefText(rw1, c1) + ":" + RefText(rw2, c2));
                    break;
                }

                case 0x29: // MemFunc
                    if (pos + 4 > len) return null;
                    pos += 4;
                    break;

                case 0x19: // PtgAttr（2 字节）
                {
                    if (pos >= len) return null;
                    int grbit = rpn[pos++];
                    switch (grbit)
                    {
                        case 0x00: // Noop
                        case 0x02: // If
                        case 0x08: // Goto
                        case 0x80: // IfError
                            if (pos + 2 > len) return null;
                            pos += 2;
                            break;
                        case 0x01: // Semi
                            break;
                        case 0x04: // Choose
                            if (pos + 2 > len) return null;
                            int cce = ReadU16(rpn, pos);
                            pos += 2 + 2 * cce;
                            if (pos > len) return null;
                            break;
                        case 0x10: // Sum
                            if (pos + 2 > len) return null;
                            pos += 2;
                            if (stack.Count < 1) return null;
                            stack[stack.Count - 1] = "SUM(" + stack[stack.Count - 1] + ")";
                            break;
                        case 0x40: // Space
                        case 0x41: // SpaceSemi
                            if (pos + 2 > len) return null;
                            pos += 2;
                            break;
                        default:
                            return null; // Baxcel 等不支持
                    }
                    break;
                }

                default:
                    return null;
            }
        }

        return stack.Count == 1 ? stack[0] : null;
    }

    // ── 辅助 ──

    /// <summary>ptg → 规范化 token 类（对齐 SheetJS PtgDupes）。</summary>
    private static int Token(int ptg)
    {
        if (ptg <= 0x3D) return ptg;
        switch (ptg)
        {
            case 0x58: case 0x78: return 0x22;
            case 0x59: case 0x79: return 0x39;
            case 0x5A: case 0x7A: return 0x3A;
            case 0x5B: case 0x7B: return 0x3B;
            case 0x5C: case 0x7C: return 0x3C;
            case 0x5D: case 0x7D: return 0x3D;
            case 0x60: case 0x61: case 0x62: case 0x63:
            case 0x64: case 0x65: case 0x66: case 0x67:
            case 0x68: case 0x69: case 0x6A: case 0x6B:
            case 0x6C: case 0x6D: case 0x6E: case 0x6F:
                return ptg - 0x40;
            default:
                return ptg - 0x20;
        }
    }

    private static bool BinOp(List<string> stack, string op)
    {
        if (stack.Count < 2) return false;
        string e1 = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        string e2 = stack[stack.Count - 1];
        stack.RemoveAt(stack.Count - 1);
        stack.Add(e2 + op + e1);
        return true;
    }

    private static bool UnOp(List<string> stack, string op)
    {
        if (stack.Count < 1) return false;
        stack[stack.Count - 1] = op + stack[stack.Count - 1];
        return true;
    }

    private static bool ApplyFunction(List<string> stack, string name, int argc)
    {
        if (stack.Count < argc) return false;
        int start = stack.Count - argc;
        var sb = new StringBuilder(name);
        sb.Append('(');
        for (int i = start; i < stack.Count; i++)
        {
            if (i > start) sb.Append(',');
            sb.Append(stack[i]);
        }
        sb.Append(')');
        stack.RemoveRange(start, argc);
        stack.Add(sb.ToString());
        return true;
    }

    private static int ReadRow(byte[] rpn, ref int pos, bool biff12)
    {
        if (biff12)
        {
            if (pos + 4 > rpn.Length) return -1;
            int rw = ReadU32(rpn, pos);
            pos += 4;
            return rw;
        }
        if (pos + 2 > rpn.Length) return -1;
        int r = ReadU16(rpn, pos);
        pos += 2;
        return r;
    }

    private static int ReadCol(byte[] rpn, ref int pos)
    {
        if (pos + 2 > rpn.Length) return -1;
        int col = ReadU16(rpn, pos);
        pos += 2;
        return col & 0x3FFF;
    }

    /// <summary>公式字符串（PtgStr）：BIFF8 = cch(1) grbit(1) chars；BIFF12 = cch(2) UTF-16。</summary>
    private static string? ReadString(byte[] rpn, ref int pos, bool biff12)
    {
        if (biff12)
        {
            if (pos + 2 > rpn.Length) return null;
            int cch = ReadU16(rpn, pos);
            pos += 2;
            if (pos + cch * 2 > rpn.Length) return null;
            var s = Encoding.Unicode.GetString(rpn, pos, cch * 2);
            pos += cch * 2;
            return s;
        }
        if (pos + 2 > rpn.Length) return null;
        int cch2 = rpn[pos++];
        byte grbit = rpn[pos++];
        bool highByte = (grbit & 0x01) != 0;
        if (highByte)
        {
            if (pos + cch2 * 2 > rpn.Length) return null;
            var s = Encoding.Unicode.GetString(rpn, pos, cch2 * 2);
            pos += cch2 * 2;
            return s;
        }
        if (pos + cch2 > rpn.Length) return null;
        var s2 = Encoding.GetEncoding(28591).GetString(rpn, pos, cch2);
        pos += cch2;
        return s2;
    }

    private static string RefText(int rw, int col)
    {
        var sb = new StringBuilder();
        int c = col;
        while (c >= 0)
        {
            sb.Insert(0, (char)('A' + (c % 26)));
            c = c / 26 - 1;
            if (c < 0) break;
        }
        sb.Append(rw + 1);
        return sb.ToString();
    }

    private static string ErrorText(byte code) => code switch
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

    private static ushort ReadU16(byte[] d, int off) =>
        (ushort)(d[off] | (d[off + 1] << 8));

    private static int ReadU32(byte[] d, int off) =>
        d[off] | (d[off + 1] << 8) | (d[off + 2] << 16) | (d[off + 3] << 24);
}
