using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace LiteExcel.Internal.Biff;

/// <summary>
/// A1 公式文本 → BIFF8/BIFF12 RPN 编码器。
/// 支持常量（数字/字符串/布尔）、单元格引用（含绝对）、区域引用、
/// 基础运算符（+ - * / ^ &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;=）、一元负号、
/// 内置函数（SUM/AVERAGE/MIN/MAX/COUNT/IF 等）、括号和嵌套。
/// 不支持的公式返回 false，调用方降级为缓存值写出。
/// </summary>
internal static class FormulaEncoder
{
    /// <summary>尝试将 A1 公式文本编码为 RPN 字节。失败返回 null。</summary>
    public static byte[]? TryEncode(string formula, bool biff12)
    {
        if (string.IsNullOrEmpty(formula)) return null;
        var f = formula.Trim();
        if (f.StartsWith("=")) f = f.Substring(1);
        if (string.IsNullOrEmpty(f)) return null;

        var tokens = Tokenize(f);
        if (tokens is null || tokens.Count == 0) return null;

        var rpn = InfixToRpn(tokens);
        if (rpn is null || rpn.Count == 0) return null;

        var result = EncodeRpn(rpn, biff12);
        return result;
    }

    // ── 词法分析 ──

    private readonly struct Token
    {
        public readonly TokenType Type;
        public readonly string Text;
        public readonly double Num;
        public Token(TokenType type, string text, double num = 0) { Type = type; Text = text; Num = num; }
    }

    private enum TokenType { Number, String, Bool, Ref, Area, Func, Op, LParen, RParen, Comma }

    private static readonly Dictionary<string, int> OpPrec = new()
    {
        ["^"] = 5, ["*"] = 4, ["/"] = 4, ["+"] = 3, ["-"] = 3, ["&"] = 2,
        ["="] = 1, ["<>"] = 1, ["<"] = 1, [">"] = 1, ["<="] = 1, [">="] = 1,
    };

    private static List<Token>? Tokenize(string s)
    {
        var tokens = new List<Token>();
        int i = 0;
        while (i < s.Length)
        {
            char c = s[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }

            if (c == '(') { tokens.Add(new Token(TokenType.LParen, "(")); i++; continue; }
            if (c == ')') { tokens.Add(new Token(TokenType.RParen, ")")); i++; continue; }
            if (c == ',') { tokens.Add(new Token(TokenType.Comma, ",")); i++; continue; }

            if (c == '"')
            {
                var sb = new StringBuilder();
                i++;
                while (i < s.Length)
                {
                    if (s[i] == '"')
                    {
                        if (i + 1 < s.Length && s[i + 1] == '"') { sb.Append('"'); i += 2; }
                        else { i++; break; }
                    }
                    else { sb.Append(s[i]); i++; }
                }
                tokens.Add(new Token(TokenType.String, sb.ToString()));
                continue;
            }

            if (char.IsDigit(c) || (c == '.' && i + 1 < s.Length && char.IsDigit(s[i + 1])))
            {
                int start = i;
                while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '.' || s[i] == 'e' || s[i] == 'E'))
                {
                    if ((s[i] == 'e' || s[i] == 'E') && i + 1 < s.Length && (s[i + 1] == '+' || s[i + 1] == '-'))
                        i++;
                    i++;
                }
                var numStr = s.Substring(start, i - start);
                if (!double.TryParse(numStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                    return null;
                tokens.Add(new Token(TokenType.Number, numStr, num));
                continue;
            }

            if (c == '-' || c == '+')
            {
                if (tokens.Count == 0 || tokens[tokens.Count - 1].Type == TokenType.Op || tokens[tokens.Count - 1].Type == TokenType.LParen || tokens[tokens.Count - 1].Type == TokenType.Comma)
                {
                    tokens.Add(new Token(TokenType.Op, c == '-' ? "u-" : "u+"));
                    i++;
                    continue;
                }
                tokens.Add(new Token(TokenType.Op, c.ToString()));
                i++;
                continue;
            }

            if (c == '<' || c == '>' || c == '=')
            {
                string op = c.ToString();
                if (i + 1 < s.Length)
                {
                    if (c == '<' && s[i + 1] == '>') { op = "<>"; i += 2; continue; }
                    if ((c == '<' || c == '>') && s[i + 1] == '=') { op = c + "="; i += 2; continue; }
                }
                tokens.Add(new Token(TokenType.Op, op));
                i++;
                continue;
            }

            if (c == '^' || c == '*' || c == '/' || c == '&')
            {
                tokens.Add(new Token(TokenType.Op, c.ToString()));
                i++;
                continue;
            }

            if (c == '$' || char.IsLetter(c) || c == '_' || c == '\'')
            {
                int start = i;
                bool hasSheet = false;
                while (i < s.Length && s[i] == '\'') { i++; hasSheet = true; while (i < s.Length && s[i] != '\'') i++; if (i < s.Length) i++; }
                while (i < s.Length && !IsTokenBreak(s[i]))
                {
                    if (s[i] == '!')
                    {
                        hasSheet = true;
                        i++;
                        while (i < s.Length && s[i] == '$') i++;
                        start = i;
                    }
                    else if (s[i] == ':') break;
                    else i++;
                }
                var word = s.Substring(hasSheet ? start : start, i - start);
                if (i < s.Length && s[i] == ':')
                {
                    i++;
                    int s2 = i;
                    while (i < s.Length && !IsTokenBreak(s[i]) && s[i] != ':') i++;
                    var end = s.Substring(s2, i - s2);
                    tokens.Add(new Token(TokenType.Area, word + ":" + end));
                    continue;
                }

                if (string.Equals(word, "TRUE", StringComparison.OrdinalIgnoreCase))
                    tokens.Add(new Token(TokenType.Bool, "TRUE", 1));
                else if (string.Equals(word, "FALSE", StringComparison.OrdinalIgnoreCase))
                    tokens.Add(new Token(TokenType.Bool, "FALSE", 0));
                else if (i < s.Length && s[i] == '(')
                {
                    tokens.Add(new Token(TokenType.Func, word.ToUpperInvariant()));
                    i++; // consume '(' so it doesn't become a separate LParen
                    continue;
                }
                else if (IsCellRef(word))
                    tokens.Add(new Token(TokenType.Ref, word));
                else
                    return null;
                continue;
            }

            return null;
        }
        return tokens;
    }

    private static bool IsTokenBreak(char c) =>
        char.IsWhiteSpace(c) || c == '(' || c == ')' || c == ',' ||
        c == '+' || c == '-' || c == '*' || c == '/' || c == '^' ||
        c == '&' || c == '=' || c == '<' || c == '>' || c == '"';

    private static bool IsCellRef(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        int i = 0;
        if (s[0] == '$') i++;
        if (i >= s.Length || !char.IsLetter(s[i])) return false;
        while (i < s.Length && (char.IsLetter(s[i]) || s[i] == '$')) i++;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        return i == s.Length;
    }

    // ── Shunting Yard ──

    private static List<Token>? InfixToRpn(List<Token> tokens)
    {
        var output = new List<Token>();
        var stack = new Stack<Token>();
        int funcArgCount = 0;
        var argCounts = new Stack<int>();

        foreach (var tok in tokens)
        {
            switch (tok.Type)
            {
                case TokenType.Number:
                case TokenType.String:
                case TokenType.Bool:
                case TokenType.Ref:
                case TokenType.Area:
                    output.Add(tok);
                    break;

                case TokenType.Func:
                    stack.Push(tok);
                    argCounts.Push(1);
                    break;

                case TokenType.Comma:
                    while (stack.Count > 0 && stack.Peek().Type != TokenType.LParen && stack.Peek().Type != TokenType.Func)
                        output.Add(stack.Pop());
                    if (argCounts.Count > 0)
                        argCounts.Push(argCounts.Pop() + 1);
                    break;

                case TokenType.Op:
                    string op = tok.Text;
                    if (op == "u-" || op == "u+")
                    {
                        stack.Push(tok);
                    }
                    else
                    {
                        while (stack.Count > 0 && stack.Peek().Type == TokenType.Op)
                        {
                            var top = stack.Peek().Text;
                            if (top == "u-" || top == "u+") break;
                            if (OpPrec.TryGetValue(top, out var prec) && prec >= OpPrec[op])
                                output.Add(stack.Pop());
                            else break;
                        }
                        stack.Push(tok);
                    }
                    break;

                case TokenType.LParen:
                    stack.Push(tok);
                    break;

                case TokenType.RParen:
                    while (stack.Count > 0 && stack.Peek().Type != TokenType.LParen && stack.Peek().Type != TokenType.Func)
                        output.Add(stack.Pop());
                    if (stack.Count == 0) return null;
                    if (stack.Peek().Type == TokenType.Func)
                    {
                        var func = stack.Pop();
                        int argc = argCounts.Pop();
                        output.Add(new Token(TokenType.Func, func.Text, argc));
                    }
                    else
                        stack.Pop(); // LParen
                    break;
            }
        }

        while (stack.Count > 0)
        {
            var t = stack.Pop();
            if (t.Type == TokenType.LParen || t.Type == TokenType.RParen) return null;
            output.Add(t);
        }

        return output;
    }

    // ── RPN → Ptg 编码 ──

    private static byte[]? EncodeRpn(List<Token> rpn, bool biff12)
    {
        var ms = new MemoryStream();

        foreach (var tok in rpn)
        {
            switch (tok.Type)
            {
                case TokenType.Number:
                    if (EncodeNumber(ms, tok.Num) is false) return null;
                    break;

                case TokenType.String:
                    if (EncodeString(ms, tok.Text, biff12) is false) return null;
                    break;

                case TokenType.Bool:
                    ms.WriteByte(0x1D); // PtgBool
                    ms.WriteByte((byte)tok.Num);
                    break;

                case TokenType.Ref:
                    if (EncodeRef(ms, tok.Text, biff12) is false) return null;
                    break;

                case TokenType.Area:
                    if (EncodeArea(ms, tok.Text, biff12) is false) return null;
                    break;

                case TokenType.Op:
                    byte opPtg = tok.Text switch
                    {
                        "+" => 0x03, "-" => 0x04, "*" => 0x05, "/" => 0x06,
                        "^" => 0x07, "&" => 0x08, "=" => 0x0B, "<>" => 0x0E,
                        "<" => 0x09, ">" => 0x0D, "<=" => 0x0A, ">=" => 0x0C,
                        "u-" => 0x13, "u+" => 0x12,
                        _ => (byte)0,
                    };
                    if (opPtg == 0) return null;
                    ms.WriteByte(opPtg);
                    break;

                case TokenType.Func:
                    var iftab = FormulaFtab.LookupName(tok.Text);
                    if (iftab < 0) return null;
                    int argc = (int)tok.Num;
                    bool isVar = FormulaFtab.IsVarArg(iftab);
                    if (isVar)
                    {
                        ms.WriteByte(0x22); // PtgFuncVar
                        ms.WriteByte((byte)argc);
                        WriteU16(ms, (ushort)iftab);
                    }
                    else
                    {
                        ms.WriteByte(0x21); // PtgFunc
                        WriteU16(ms, (ushort)iftab);
                    }
                    break;

                default:
                    return null;
            }
        }

        return ms.ToArray();
    }

    private static bool EncodeNumber(MemoryStream ms, double num)
    {
        if (num == Math.Floor(num) && num >= -32768 && num <= 32767)
        {
            ms.WriteByte(0x1E); // PtgInt
            WriteU16(ms, (ushort)(int)num);
        }
        else
        {
            ms.WriteByte(0x1F); // PtgNum
            var b = BitConverter.GetBytes(num);
            ms.Write(b, 0, 8);
        }
        return true;
    }

    private static bool EncodeString(MemoryStream ms, string s, bool biff12)
    {
        ms.WriteByte(0x17); // PtgStr
        if (biff12)
        {
            WriteU16(ms, (ushort)s.Length);
            var b = Encoding.Unicode.GetBytes(s);
            ms.Write(b, 0, b.Length);
        }
        else
        {
            ms.WriteByte((byte)s.Length);
            ms.WriteByte(0x01); // grbit: high byte (UTF-16LE)
            var b = Encoding.Unicode.GetBytes(s);
            ms.Write(b, 0, b.Length);
        }
        return true;
    }

    private static bool EncodeRef(MemoryStream ms, string refStr, bool biff12)
    {
        var (row, col, rowAbs, colAbs) = ParseCellRef(refStr);
        if (row < 0 || col < 0) return false;

        ms.WriteByte(0x24); // PtgRef

        if (biff12)
        {
            WriteU32(ms, (uint)row);
        }
        else
        {
            WriteU16(ms, (ushort)row);
        }

        ushort colWord = (ushort)(col & 0x3FFF);
        if (colAbs) colWord |= 0x4000;
        if (rowAbs) colWord |= 0x8000;
        WriteU16(ms, colWord);

        return true;
    }

    private static bool EncodeArea(MemoryStream ms, string areaStr, bool biff12)
    {
        int colon = areaStr.IndexOf(':');
        if (colon < 0) return false;
        var first = areaStr.Substring(0, colon);
        var last = areaStr.Substring(colon + 1);

        var (r1, c1, r1a, c1a) = ParseCellRef(first);
        var (r2, c2, r2a, c2a) = ParseCellRef(last);
        if (r1 < 0 || c1 < 0 || r2 < 0 || c2 < 0) return false;

        ms.WriteByte(0x25); // PtgArea

        if (biff12)
        {
            WriteU32(ms, (uint)r1);
            WriteU32(ms, (uint)r2);
        }
        else
        {
            WriteU16(ms, (ushort)r1);
            WriteU16(ms, (ushort)r2);
        }

        ushort c1w = (ushort)(c1 & 0x3FFF);
        ushort c2w = (ushort)(c2 & 0x3FFF);
        if (c1a) c1w |= 0x4000;
        if (r1a) c1w |= 0x8000;
        if (c2a) c2w |= 0x4000;
        if (r2a) c2w |= 0x8000;
        WriteU16(ms, c1w);
        WriteU16(ms, c2w);

        return true;
    }

    private static (int row, int col, bool rowAbs, bool colAbs) ParseCellRef(string s)
    {
        int i = 0;
        bool colAbs = false, rowAbs = false;

        int sheetEnd = s.IndexOf('!');
        if (sheetEnd >= 0) i = sheetEnd + 1;

        if (i < s.Length && s[i] == '$') { colAbs = true; i++; }
        int colStart = i;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == colStart) return (-1, -1, false, false);
        string colStr = s.Substring(colStart, i - colStart);
        int col = 0;
        for (int k = 0; k < colStr.Length; k++)
            col = col * 26 + (char.ToUpperInvariant(colStr[k]) - 'A' + 1);
        col--;

        if (i < s.Length && s[i] == '$') { rowAbs = true; i++; }
        int rowStart = i;
        while (i < s.Length && char.IsDigit(s[i])) i++;
        if (i == rowStart) return (-1, -1, false, false);
        int row = int.Parse(s.Substring(rowStart, i - rowStart)) - 1;

        return (row, col, rowAbs, colAbs);
    }

    private static void WriteU16(MemoryStream ms, ushort v)
    {
        ms.WriteByte((byte)v);
        ms.WriteByte((byte)(v >> 8));
    }

    private static void WriteU32(MemoryStream ms, uint v)
    {
        ms.WriteByte((byte)v);
        ms.WriteByte((byte)(v >> 8));
        ms.WriteByte((byte)(v >> 16));
        ms.WriteByte((byte)(v >> 24));
    }
}
