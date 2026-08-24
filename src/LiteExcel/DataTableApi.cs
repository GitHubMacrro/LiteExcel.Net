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

        var headers = sheet.Headers;
        int colCount = headers.Count;
        if (colCount == 0 && sheet.Rows.Count > 0)
        {
            colCount = sheet.Rows[0].Count;
            headers = new List<string>(colCount);
            for (int i = 0; i < colCount; i++)
                headers.Add($"Column{i + 1}");
        }

        foreach (var header in headers)
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

    internal static SheetData DataTableToSheet(DataTable table, string sheetName)
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
                cells.Add(CellFactory.FromObject(dataRow[i]));
            }
            sheet.Rows.Add(cells);
        }

        return sheet;
    }
}
