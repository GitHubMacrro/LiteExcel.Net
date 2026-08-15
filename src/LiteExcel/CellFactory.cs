using System.Data;

namespace LiteExcel;

/// <summary>
/// 内部工具：把 CLR 值转换为 Cell
/// </summary>
internal static class CellFactory
{
    public static Cell FromObject(object? value)
    {
        if (value is null || value == DBNull.Value) return Cell.Empty;

        return value switch
        {
            bool b => Cell.FromBoolean(b),
            DateTime dt => Cell.FromDate(dt),
            sbyte n => Cell.FromNumber(n),
            byte n => Cell.FromNumber(n),
            short n => Cell.FromNumber(n),
            ushort n => Cell.FromNumber(n),
            int n => Cell.FromNumber(n),
            uint n => Cell.FromNumber(n),
            long n => Cell.FromNumber(n),
            ulong n => Cell.FromNumber(n),
            float n => Cell.FromNumber(n),
            double n => Cell.FromNumber(n),
            decimal n => Cell.FromNumber((double)n),
            Cell c => c,
            _ => Cell.FromText(value.ToString()),
        };
    }
}
