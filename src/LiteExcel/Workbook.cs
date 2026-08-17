using LiteExcel.Internal;
using LiteExcel.Internal.Biff;
using System.IO;
using System.Linq;

namespace LiteExcel;

/// <summary>
/// 高层工作簿模型（文件级）。
/// 负责工作表集合、文档属性、保存/另存为。
/// 打开时加载到内存，不长期持有文件流，因此无需 IDisposable。
/// </summary>
public sealed class Workbook
{
    private string? _currentPath;
    private List<string>? _openedSheetNames;

    /// <summary>工作表集合 </summary>
    public WorksheetCollection Worksheets { get; }

    /// <summary>文档属性（作者/时间/标题等） </summary>
    public WorkbookProperties Properties { get; }

    /// <summary>当前工作簿格式 </summary>
    public ExcelFormat Format { get; private set; }

    /// <summary>
    /// 打开时捕获的、写入器不重建的 OOXML 部件（宏/主题/绘图/图表等）。
    /// 保存时按二进制透传，避免未映射部件被静默删除。新建工作簿为 null。
    /// </summary>
    internal OoxmlPreservedParts? PreservedParts { get; set; }

    /// <summary>
    /// 打开时捕获的 VBA 宏工程原始字节（xl/vbaProject.bin）。写入 xlsb 时透传保留。
    /// 新建工作簿或源文件无宏时为 null。
    /// </summary>
    internal byte[]? VbaProjectBytes { get; set; }

    /// <summary>打开时捕获的工作簿宿主 VBA 代码名（workbookPr@codeName / BrtWbProp codeName） </summary>
    internal string? WorkbookCodeName { get; set; }

    /// <summary>
    /// 当前目标路径。
    /// <see cref="Open"/> 后指向源文件；<see cref="SaveAs"/> 后更新为新路径；
    /// <see cref="Create"/> 后为 null（此时只能 SaveAs）。
    /// </summary>
    public string? CurrentPath => _currentPath;

    private Workbook()
    {
        Properties = new WorkbookProperties();
        Worksheets = new WorksheetCollection(this);
    }

    internal static Workbook CreateEmpty(ExcelFormat format)
    {
        var wb = new Workbook { Format = format };
        return wb;
    }

    internal static Workbook FromSheetData(IReadOnlyList<SheetData> sheets, WorkbookProperties? properties, ExcelFormat format, string? path)
    {
        var wb = new Workbook
        {
            Format = format,
            _currentPath = path,
            _openedSheetNames = sheets.Select(s => s.SheetName).ToList(),
        };
        if (properties is not null)
        {
            wb.Properties.Creator = properties.Creator;
            wb.Properties.LastModifiedBy = properties.LastModifiedBy;
            wb.Properties.Created = properties.Created;
            wb.Properties.Modified = properties.Modified;
            wb.Properties.Title = properties.Title;
            wb.Properties.Subject = properties.Subject;
            wb.Properties.Application = properties.Application;
        }
        foreach (var sheet in sheets)
        {
            var ws = Worksheet.FromSheetData(sheet);
            wb.Worksheets.AddInternal(ws);
            wb.OnWorksheetAdded(ws);
        }
        return wb;
    }

    // ── 保存 ──

    /// <summary>保存到当前目标路径。若当前无路径（新建），抛出 <see cref="LiteExcelException"/> </summary>
    public void Save()
    {
        if (string.IsNullOrEmpty(_currentPath))
            throw new LiteExcelException("当前工作簿没有目标路径，请使用 SaveAs 指定保存位置");
        SaveCore(_currentPath, Format);
    }

    /// <summary>另存为指定路径。格式沿用当前格式 </summary>
    public void SaveAs(string path)
    {
        SaveAs(path, Format);
    }

    /// <summary>另存为指定路径并指定格式（格式必须为已支持的可写格式）。路径扩展名必须与 format 匹配，否则抛 <see cref="LiteExcelException"/> </summary>
    public void SaveAs(string path, ExcelFormat format)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));

        ValidateExtension(path, format);
        SaveCore(path, format);
        _currentPath = path;
        Format = format;
    }

    /// <summary>保存到流并指定格式。不更新 <see cref="CurrentPath"/> </summary>
    public void Save(Stream stream, ExcelFormat format)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("流不可写", nameof(stream));

        SaveCore(stream, format);
    }

    /// <summary>校验保存路径的扩展名与目标格式一致，避免写出内容与扩展名不匹配、Excel 无法打开的文件 </summary>
    internal static void ValidateExtension(string path, ExcelFormat format)
    {
        string ext = System.IO.Path.GetExtension(path);
        string expected = format switch
        {
            ExcelFormat.Xlsx => ".xlsx",
            ExcelFormat.Xlsm => ".xlsm",
            ExcelFormat.Csv => ".csv",
            ExcelFormat.Xls => ".xls",
            ExcelFormat.Xlsb => ".xlsb",
            _ => null,
        };
        if (expected is not null && !string.Equals(ext, expected, System.StringComparison.OrdinalIgnoreCase))
            throw new LiteExcelException($"保存路径扩展名 '{ext}' 与目标格式 {format}（应为 '{expected}'）不匹配，Excel 将无法按预期打开该文件。请使用匹配的扩展名，例如 SaveAs(\"out{expected}\", ExcelFormat.{format})。");
    }

    private void SaveCore(string path, ExcelFormat format)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        SaveCore(fs, format);
    }

    private void SaveCore(Stream stream, ExcelFormat format)
    {
        switch (format)
        {
            case ExcelFormat.Xlsx:
            case ExcelFormat.Xlsm:
                var sheets = BuildSheetDataList();
                bool structureUnchanged = StructureUnchanged(sheets);
                XlsxWriter.Write(stream, sheets, Properties, PreservedParts, mergeSheetRels: structureUnchanged,
                    macroEnabled: format == ExcelFormat.Xlsm);
                break;
            case ExcelFormat.Csv:
                if (Worksheets.Count != 1)
                    throw new NotSupportedException("CSV 仅支持单工作表工作簿");
                CsvBackend.Write(stream, Worksheets[0].ToSheetData());
                break;
            case ExcelFormat.Xls:
                var xlsSheets = BuildSheetDataList();
                XlsWriter.Write(stream, xlsSheets);
                break;
            case ExcelFormat.Xlsb:
                var xlsbSheets = BuildSheetDataList();
                XlsbWriter.Write(stream, xlsbSheets, VbaProjectBytes, WorkbookCodeName);
                break;
            default:
                throw new NotSupportedException($"未知格式：{format}");
        }
    }

    internal List<SheetData> BuildSheetDataList()
    {
        var list = new List<SheetData>(Worksheets.Count);
        foreach (var ws in Worksheets)
            list.Add(ws.ToSheetData());
        return list;
    }

    /// <summary>工作表数量与顺序相对打开时是否未变（决定能否复用工作表级保留 rels） </summary>
    private bool StructureUnchanged(List<SheetData> sheets)
    {
        if (_openedSheetNames is null || sheets.Count != _openedSheetNames.Count)
            return false;
        for (int i = 0; i < sheets.Count; i++)
        {
            if (sheets[i].SheetName != _openedSheetNames[i])
                return false;
        }
        return true;
    }

    // ── 集合回调 ──

    internal void OnWorksheetAdded(Worksheet ws)
    {
        // 预留：工作簿级联动（如记录 Modified）
        Properties.Modified = DateTime.Now;
    }

    internal void OnWorksheetRemoved(Worksheet ws)
    {
        Properties.Modified = DateTime.Now;
    }
}
