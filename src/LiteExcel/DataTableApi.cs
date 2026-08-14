using System.Collections;
using System.Data;

namespace LiteExcel;

/// <summary>
/// DataTable 便利 API（AOT 安全，无反射） 
/// </summary>
public static partial class XlsxReader
{
    /// <summary>
    /// 将指定工作表读为 DataTable 第一行作为列头 
    /// </summary>
    public static DataTable ReadAsDataTable(string path, int sheetIndex = 0, bool firstRowIsHeader = true)
    {
        var sheet = Read(path, sheetIndex, firstRowIsHeader);
        return SheetToDataTable(sheet);
    }

    /// <summary>
    /// 将指定工作表读为 DataTable 第一行作为列头 
    /// </summary>
    public static DataTable ReadAsDataTable(string path, string sheetName, bool firstRowIsHeader = true)
    {
        var sheet = Read(path, sheetName, firstRowIsHeader);
        return SheetToDataTable(sheet);
    }

    /// <summary>
    /// 将指定工作表读为 DataTable 第一行作为列头 
    /// </summary>
    public static DataTable ReadAsDataTable(Stream stream, int sheetIndex = 0, bool firstRowIsHeader = true)
    {
        var sheet = Read(stream, sheetIndex, firstRowIsHeader);
        return SheetToDataTable(sheet);
    }

    /// <summary>
    /// 将指定工作表读为 DataTable 第一行作为列头 
    /// </summary>
    public static DataTable ReadAsDataTable(Stream stream, string sheetName, bool firstRowIsHeader = true)
    {
        var sheet = Read(stream, sheetName, firstRowIsHeader);
        return SheetToDataTable(sheet);
    }

    private static DataTable SheetToDataTable(SheetData sheet)
    {
        var dt = new DataTable(sheet.SheetName);

        int colCount = sheet.Headers.Count;
        if (colCount == 0 && sheet.Rows.Count > 0)
        {
            colCount = sheet.Rows[0].Count;
            for (int i = 0; i < colCount; i++)
                sheet.Headers.Add($"Column{i + 1}");
        }

        foreach (var header in sheet.Headers)
        {
            dt.Columns.Add(header, typeof(object));
        }

        foreach (var row in sheet.Rows)
        {
            var values = new object?[colCount];
            for (int i = 0; i < colCount && i < row.Count; i++)
            {
                values[i] = CellToObject(row[i]);
            }
            dt.Rows.Add(values);
        }

        return dt;
    }

    private static object? CellToObject(Cell cell)
    {
        return cell.Type switch
        {
            CellType.Empty => DBNull.Value,
            CellType.Text => cell.Text,
            CellType.Number => cell.Number,
            CellType.Date => cell.Date,
            CellType.Boolean => cell.Boolean,
            _ => cell.Text,
        };
    }
}

/// <summary>
/// DataTable 便利 API（AOT 安全，无反射） 
/// </summary>
public static partial class XlsxWriter
{
    /// <summary>
    /// 将 DataTable 写入 xlsx 文件 
    /// </summary>
    public static void Write(string path, DataTable table, string sheetName = "Sheet1")
    {
        var sheet = DataTableToSheet(table, sheetName);
        Write(path, sheet);
    }

    private static SheetData DataTableToSheet(DataTable table, string sheetName)
    {
        var sheet = new SheetData { SheetName = sheetName };

        foreach (DataColumn col in table.Columns)
        {
            sheet.Headers.Add(col.ColumnName);
        }

        foreach (DataRow dataRow in table.Rows)
        {
            var cells = new List<Cell>(table.Columns.Count);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                cells.Add(ObjectToCell(dataRow[i]));
            }
            sheet.Rows.Add(cells);
        }

        return sheet;
    }

    private static Cell ObjectToCell(object? value)
    {
        if (value == null || value == DBNull.Value) return Cell.Empty;

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
            _ => Cell.FromText(value.ToString()),
        };
    }
}
