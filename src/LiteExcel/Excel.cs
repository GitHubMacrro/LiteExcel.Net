using LiteExcel.Internal;
using LiteExcel.Internal.Biff;
using System.Data;
using System.IO;
using System.IO.Compression;

namespace LiteExcel;

/// <summary>
/// 高层统一入口（门面）。
/// 用户只感知本类，不感知底层 Reader/Writer 与格式后端差异。
/// </summary>
public static class Excel
{
    // ── 打开 / 新建 ──

    /// <summary>打开工作簿，按扩展名自动识别格式。已支持 xlsx/xlsm/xls/xlsb/csv </summary>
    public static Workbook Open(string path, ExcelReadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        if (!File.Exists(path))
            throw new FileNotFoundException("文件不存在", path);

        var format = DetectFormat(path);
        return OpenCore(path, format, options);
    }

    /// <summary>以指定格式打开工作簿 </summary>
    public static Workbook Open(string path, ExcelFormat format, ExcelReadOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        if (!File.Exists(path))
            throw new FileNotFoundException("文件不存在", path);
        return OpenCore(path, format, options);
    }

    private static Workbook OpenCore(string path, ExcelFormat format, ExcelReadOptions? options)
    {
        options ??= new ExcelReadOptions();

        switch (format)
        {
            case ExcelFormat.Xlsx:
            case ExcelFormat.Xlsm:
                break;
            case ExcelFormat.Csv:
            {
                var csvSheet = CsvBackend.Read(path);
                return Workbook.FromSheetData(new[] { csvSheet }, null, ExcelFormat.Csv, path);
            }
            case ExcelFormat.Xls:
            {
                var sheets = XlsBackend.ReadAll(path);
                return Workbook.FromSheetData(sheets, null, ExcelFormat.Xls, path);
            }
            case ExcelFormat.Xlsb:
            {
                var sheets = XlsbBackend.ReadAll(path);
                return Workbook.FromSheetData(sheets, null, ExcelFormat.Xlsb, path);
            }
            default:
                throw new NotSupportedException($"{format} 读取后端尚未实现，当前仅支持 xlsx/xlsm/csv/xls/xlsb");
        }

        // 单次解压内完成读表/读属性/捕获保留部件，保证三者来自同一文件快照
        Workbook wb;
        OoxmlPreservedParts? preserved;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: true))
        {
            var sheets = XlsxReader.ReadAllRaw(zip);
            var props = XlsxReader.ReadProperties(zip);
            preserved = OoxmlPreservedParts.Capture(zip, sheets.Count);
            preserved.WorkbookCodeName = XlsxReader.WorkbookCodeNameSnapshot; // ReadWorkbook 刚捕获
            wb = Workbook.FromSheetData(sheets, props, format, path);
        }
        wb.PreservedParts = preserved;

        if (options.FillMergedCells)
        {
            foreach (var ws in wb.Worksheets)
                ws.FillMergedValues();
        }

        return wb;
    }

    /// <summary>新建工作簿（默认 Xlsx）。新建后需调用 SaveAs 指定路径 </summary>
    public static Workbook Create(ExcelFormat format = ExcelFormat.Xlsx)
    {
        if (format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsm && format != ExcelFormat.Csv
            && format != ExcelFormat.Xls && format != ExcelFormat.Xlsb)
            throw new NotSupportedException($"{format} 写入后端尚未实现，当前仅支持 xlsx/xlsm/csv/xls/xlsb");
        var wb = Workbook.CreateEmpty(format);
        wb.Worksheets.Add("Sheet1");
        return wb;
    }

    /// <summary>新建工作簿并指定首个工作表名 </summary>
    public static Workbook Create(string sheetName, ExcelFormat format = ExcelFormat.Xlsx)
    {
        var wb = Create(format);
        if (string.IsNullOrWhiteSpace(sheetName)) return wb;
        wb.Worksheets.Remove("Sheet1");
        wb.Worksheets.Add(sheetName);
        return wb;
    }

    /// <summary>新建工作簿并批量添加工作表。传 null 或空数组时保留默认 Sheet1 </summary>
    public static Workbook Create(string[] sheetNames, ExcelFormat format = ExcelFormat.Xlsx)
    {
        var wb = Create(format);
        if (sheetNames is null || sheetNames.Length == 0) return wb;
        wb.Worksheets.Remove("Sheet1");
        foreach (var name in sheetNames)
            wb.Worksheets.Add(name);
        return wb;
    }

    /// <summary>根据扩展名识别格式 </summary>
    public static ExcelFormat DetectFormat(string path)
    {
        var ext = Path.GetExtension(path)?.ToLowerInvariant() ?? "";
        return ext switch
        {
            ".xlsm" => ExcelFormat.Xlsm,
            ".xlsb" => ExcelFormat.Xlsb,
            ".xls" => ExcelFormat.Xls,
            ".csv" => ExcelFormat.Csv,
            _ => ExcelFormat.Xlsx,
        };
    }

    // ── 写出 ──

    /// <summary>写出工作簿（按工作簿当前格式写 .xlsx/.xlsm；或按 options 指定） </summary>
    public static void Write(string path, Workbook workbook, ExcelWriteOptions? options = null)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));

        options ??= new ExcelWriteOptions();
        var format = workbook.Format;

        // 根据扩展名推断目标格式（若与工作簿格式冲突，以扩展名为准）
        var extFormat = DetectFormat(path);
        if (extFormat is ExcelFormat.Xlsx or ExcelFormat.Xlsm or ExcelFormat.Csv)
            format = extFormat;

        ApplyWriteOptions(workbook, options);
        workbook.SaveAs(path, format);
    }

    /// <summary>写出单个工作表（SheetData 低层模型，AOT 安全） </summary>
    public static void Write(string path, SheetData sheet, ExcelWriteOptions? options = null)
    {
        if (sheet is null) throw new ArgumentNullException(nameof(sheet));
        var wb = Workbook.CreateEmpty(ExcelFormat.Xlsx);
        var ws = Worksheet.FromSheetData(sheet);
        wb.Worksheets.AddInternal(ws);
        wb.OnWorksheetAdded(ws);
        Write(path, wb, options);
    }

    /// <summary>写出 DataTable（AOT 安全） </summary>
    public static void Write(string path, DataTable table, string sheetName = "Sheet1", ExcelWriteOptions? options = null)
    {
        if (table is null) throw new ArgumentNullException(nameof(table));
        var wb = Workbook.CreateEmpty(ExcelFormat.Xlsx);
        var ws = Worksheet.FromSheetData(DataTableToSheet(table, sheetName));
        wb.Worksheets.AddInternal(ws);
        wb.OnWorksheetAdded(ws);
        Write(path, wb, options);
    }

    /// <summary>写出 List&lt;T&gt;（反射映射，不兼容 AOT/裁剪） </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("List<T> 映射依赖反射，不兼容 AOT/裁剪。AOT 项目请用 DataTable 或 SheetData 重载")]
#endif
    public static void Write<T>(string path, IEnumerable<T> data, string sheetName = "Sheet1", Action<WriteOptions<T>>? configure = null)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        XlsxWriter.Write(path, data, opt =>
        {
            opt.SheetName = sheetName;
            configure?.Invoke(opt);
        });
    }

    // ── 读取便利 ──

    /// <summary>读取指定工作表为 List&lt;T&gt;（反射映射，不兼容 AOT/裁剪）。默认第一张表，首行作为表头 </summary>
#if NET5_0_OR_GREATER
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("List<T> 映射依赖反射，不兼容 AOT/裁剪。AOT 项目请用 ReadAsDataTable 或 Open 重载")]
#endif
    public static List<T> Read<T>(string path, string? sheetName = null, Action<ReadOptions<T>>? configure = null) where T : new()
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        return sheetName is null
            ? XlsxReader.Read<T>(path, 0, configure)
            : XlsxReader.Read<T>(path, sheetName, configure);
    }

    /// <summary>读取指定工作表为 DataTable（AOT 安全）。默认第一张表 </summary>
    public static DataTable ReadAsDataTable(string path, string? sheetName = null, bool firstRowIsHeader = true)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        return sheetName is null
            ? XlsxReader.ReadAsDataTable(path, 0, firstRowIsHeader)
            : XlsxReader.ReadAsDataTable(path, sheetName, firstRowIsHeader);
    }

    /// <summary>列出所有工作表名 </summary>
    public static List<string> GetSheetNames(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        return XlsxReader.GetSheetNames(path);
    }

    /// <summary>流式读取指定工作表，逐行回调，不驻留内存（自动跳过首行） </summary>
    public static void StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        if (onRow is null) throw new ArgumentNullException(nameof(onRow));
        XlsxReader.StreamRows(path, sheetName, onRow);
    }

    /// <summary>创建流式写入器（逐行写大文件，不驻留内存）。使用后调用 Dispose/Close 完成文件 </summary>
    public static XlsxStreamWriter CreateWriter(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        return XlsxStreamWriter.Create(path);
    }

    /// <summary>创建流式写入器（写入流，LeaveOpen 由调用方管理） </summary>
    public static XlsxStreamWriter CreateWriter(Stream stream)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        return XlsxStreamWriter.Create(stream);
    }

    // ── 内部辅助 ──

    private static void ApplyWriteOptions(Workbook workbook, ExcelWriteOptions options)
    {
        if (options.AutoFitColumns)
        {
            foreach (var ws in workbook.Worksheets)
            {
                var sd = ws.ToSheetData();
                XlsxWriter.AutoColumnWidths(sd);
                if (sd.ColumnWidths is not null)
                    ws.ColumnWidths = sd.ColumnWidths.Select((w, i) => (i, w)).ToDictionary(x => x.i, x => x.w);
            }
        }

        if (options.FreezeHeader)
        {
            foreach (var ws in workbook.Worksheets)
                ws.FreezeHeader = true;
        }

        if (options.Properties is not null)
        {
            workbook.Properties.Creator = options.Properties.Creator ?? workbook.Properties.Creator;
            workbook.Properties.LastModifiedBy = options.Properties.LastModifiedBy ?? workbook.Properties.LastModifiedBy;
            workbook.Properties.Title = options.Properties.Title ?? workbook.Properties.Title;
            workbook.Properties.Subject = options.Properties.Subject ?? workbook.Properties.Subject;
            workbook.Properties.Application = options.Properties.Application ?? workbook.Properties.Application;
            if (options.Properties.Created is not null) workbook.Properties.Created = options.Properties.Created;
            if (options.Properties.Modified is not null) workbook.Properties.Modified = options.Properties.Modified;
        }
    }

    private static SheetData DataTableToSheet(DataTable table, string sheetName)
    {
        var sheet = new SheetData { SheetName = sheetName };
        foreach (DataColumn col in table.Columns)
            sheet.Headers.Add(col.ColumnName);
        foreach (DataRow dataRow in table.Rows)
        {
            var cells = new List<Cell>(table.Columns.Count);
            for (int i = 0; i < table.Columns.Count; i++)
                cells.Add(CellFactory.FromObject(dataRow[i]));
            sheet.Rows.Add(cells);
        }
        return sheet;
    }
}
