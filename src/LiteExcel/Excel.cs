using LiteExcel.Internal;
using LiteExcel.Internal.Biff;
using System.Data;
using System.Diagnostics.CodeAnalysis;
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

    /// <summary>
    /// 从流打开工作簿，必须显式指定格式（流无扩展名，无法自动识别）。
    /// 输入流不会被关闭（由调用方管理）；支持不可定位的流（内部复制到内存）。
    /// 打开后 <see cref="Workbook.CurrentPath"/> 为 null，需用 <see cref="Workbook.SaveAs"/> 指定保存路径。
    /// </summary>
    public static Workbook Open(Stream stream, ExcelFormat format, ExcelReadOptions? options = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("流不可读", nameof(stream));
        return OpenFromStream(stream, format, options);
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
                var csvSheet = CsvBackend.Read(path, options.Separator);
                return Workbook.FromSheetData(new[] { csvSheet }, null, ExcelFormat.Csv, path);
            }
            case ExcelFormat.Xls:
            {
                var sheets = XlsBackend.ReadAll(path);
                var wbX = Workbook.FromSheetData(sheets, null, ExcelFormat.Xls, path);
                AttachXlsNames(wbX);
                wbX.Date1904 = XlsBackend.ReadDate1904(path);
                return wbX;
            }
            case ExcelFormat.Xlsb:
            {
                // 加密 xlsb 同样是 CFB 容器（内含 EncryptionInfo）。有密码则解密，否则识别并报错
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (!string.IsNullOrEmpty(options.OpenPassword))
                    {
                        var decrypted = DecryptWithPasswordCheck(fs, path, options.OpenPassword!);
                        var sheets = XlsbBackend.ReadAll(decrypted);
                        var wbB = Workbook.FromSheetData(sheets, null, ExcelFormat.Xlsb, path);
                        // P0-14/18: 捕获保留部件与文档属性（与读取同一解密快照）
                        decrypted.Position = 0;
                        using (var zipB = new ZipArchive(decrypted, ZipArchiveMode.Read, leaveOpen: true))
                        {
                            wbB.PreservedParts = OoxmlPreservedParts.Capture(zipB, sheets.Count, binary: true);
                            wbB.Properties.CopyFrom(XlsxReader.ReadProperties(zipB));
                        }
                        decrypted.Position = 0;
                        wbB.VbaProjectBytes = XlsbBackend.ReadVbaProject(decrypted);
                        decrypted.Position = 0;
                        wbB.WorkbookCodeName = XlsbBackend.ReadWorkbookCodeName(decrypted);
                        decrypted.Position = 0;
                        wbB.Date1904 = XlsbBackend.ReadDate1904(decrypted);
                        decrypted.Position = 0;
                        var fsB = XlsbBackend.ReadFileSharing(decrypted);
                        wbB.Security.Initialize(options.OpenPassword);
                        if (fsB is not null)
                        {
                            wbB.FileSharingToPreserve = fsB;
                            wbB.Security.Initialize(options.OpenPassword, fileHasModifyProtection: true,
                                readOnlyRecommended: fsB.ReadOnlyRecommended);
                            if (!string.IsNullOrEmpty(options.ModifyPassword))
                                wbB.Security.GrantModifyAccess(fsB.ReadOnlyRecommended);
                        }
                        return wbB;
                    }
                    EncryptionDetector.ThrowIfEncryptedOoxml(fs, path);
                }
                var sheetsX = XlsbBackend.ReadAll(path);
                var wbXB = Workbook.FromSheetData(sheetsX, null, ExcelFormat.Xlsb, path);
                // P0-14/18: 捕获保留部件与文档属性（图表/透视表/主题/绘图不再随保存丢失）
                using (var capFs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var capZip = new ZipArchive(capFs, ZipArchiveMode.Read))
                {
                    wbXB.PreservedParts = OoxmlPreservedParts.Capture(capZip, sheetsX.Count, binary: true);
                    wbXB.Properties.CopyFrom(XlsxReader.ReadProperties(capZip));
                }
                wbXB.VbaProjectBytes = XlsbBackend.ReadVbaProject(path);
                wbXB.WorkbookCodeName = XlsbBackend.ReadWorkbookCodeName(path);
                wbXB.Date1904 = XlsbBackend.ReadDate1904(path);
                ApplyFileSharing(wbXB, options, XlsbBackend.ReadFileSharing(path));
                return wbXB;
            }
            default:
                throw new NotSupportedException($"{format} 读取后端尚未实现，当前仅支持 xlsx/xlsm/csv/xls/xlsb");
        }

        // 单次解压内完成读表/读属性/捕获保留部件，保证三者来自同一文件快照
        // 加密 xlsx/xlsm 实际是 CFB 容器（内含 EncryptionInfo）。有密码则解密，否则识别并报错
        Workbook wb;
        OoxmlPreservedParts? preserved;
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            using var zipSource = !string.IsNullOrEmpty(options.OpenPassword)
                ? DecryptWithPasswordCheck(fs, path, options.OpenPassword!)
                : fs;
            if (string.IsNullOrEmpty(options.OpenPassword))
                EncryptionDetector.ThrowIfEncryptedOoxml(fs, path);
            zipSource.Position = 0;
            using var zip = new ZipArchive(zipSource, ZipArchiveMode.Read, leaveOpen: false);
            var sheets = XlsxReader.ReadAllRaw(zip);
            var props = XlsxReader.ReadProperties(zip);
            preserved = OoxmlPreservedParts.Capture(zip, sheets.Count);
            preserved.WorkbookCodeName = XlsxReader.WorkbookCodeNameSnapshot; // ReadWorkbook 刚捕获
            // P0-6: 命名区域与窗口视图原样回写；同时解析到 Workbook.Names
            preserved.BookViewsXml = XlsxReader.BookViewsXmlSnapshot;
            preserved.DefinedNamesXml = XlsxReader.DefinedNamesXmlSnapshot;
            wb = Workbook.FromSheetData(sheets, props, format, path);
            foreach (var nr in XlsxReader.ParseDefinedNames(preserved.DefinedNamesXml))
                wb.Names.Add(nr);
            if (preserved.Parts.TryGetValue("xl/vbaProject.bin", out var vbaBytes))
                wb.VbaProjectBytes = vbaBytes;
            wb.WorkbookCodeName = preserved.WorkbookCodeName;
            wb.Date1904 = XlsxReader.Date1904Snapshot;
        }
        wb.PreservedParts = preserved;
        wb.Protection = XlsxReader.WorkbookProtectionSnapshot;
        ApplyFileSharing(wb, options, XlsxReader.FileSharingSnapshot);

        if (options.FillMergedCells)
        {
            foreach (var ws in wb.Worksheets)
                ws.FillMergedValues();
        }

        return wb;
    }

    /// <summary>
    /// 从流打开工作簿。输入流不关闭；不可定位流内部复制到内存。
    /// path 为 null，<see cref="Workbook.CurrentPath"/> 为 null。
    /// </summary>
    private static Workbook OpenFromStream(Stream stream, ExcelFormat format, ExcelReadOptions? options)
    {
        options ??= new ExcelReadOptions();

        // 复制到可定位的内存流（支持网络流/响应流等不可 seek 的输入），不关闭原始流
        // Workbook 本身是内存模型，此拷贝不会显著增加峰值内存
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        ms.Position = 0;

        switch (format)
        {
            case ExcelFormat.Xlsx:
            case ExcelFormat.Xlsm:
                break;
            case ExcelFormat.Csv:
            {
                var csvSheet = CsvBackend.Read(ms, "Sheet1", options.Separator);
                return FinishOpen(new[] { csvSheet }, null, ExcelFormat.Csv, null, options);
            }
            case ExcelFormat.Xls:
            {
                var sheets = XlsBackend.ReadAll(ms);
                var wbX = Workbook.FromSheetData(sheets, null, ExcelFormat.Xls, null);
                AttachXlsNames(wbX);
                ms.Position = 0;
                wbX.Date1904 = XlsBackend.ReadDate1904(ms);
                return wbX;
            }
            case ExcelFormat.Xlsb:
            {
                // 加密 xlsb 同样是 CFB 容器。有密码则解密，否则识别并报错
                if (!string.IsNullOrEmpty(options.OpenPassword))
                {
                    var decrypted = DecryptWithPasswordCheck(ms, "<stream>", options.OpenPassword!);
                    var sheetsB = XlsbBackend.ReadAll(decrypted);
                    var wbDB = Workbook.FromSheetData(sheetsB, null, ExcelFormat.Xlsb, null);
                    decrypted.Position = 0;
                    using (var zipDB = new ZipArchive(decrypted, ZipArchiveMode.Read, leaveOpen: true))
                    {
                        wbDB.PreservedParts = OoxmlPreservedParts.Capture(zipDB, sheetsB.Count, binary: true);
                        wbDB.Properties.CopyFrom(XlsxReader.ReadProperties(zipDB));
                    }
                    decrypted.Position = 0;
                    wbDB.VbaProjectBytes = XlsbBackend.ReadVbaProject(decrypted);
                    decrypted.Position = 0;
                    wbDB.WorkbookCodeName = XlsbBackend.ReadWorkbookCodeName(decrypted);
                    decrypted.Position = 0;
                    wbDB.Date1904 = XlsbBackend.ReadDate1904(decrypted);
                    decrypted.Position = 0;
                    var fsDB = XlsbBackend.ReadFileSharing(decrypted);
                    wbDB.Security.Initialize(options.OpenPassword);
                    if (fsDB is not null)
                    {
                        wbDB.FileSharingToPreserve = fsDB;
                        wbDB.Security.Initialize(options.OpenPassword, fileHasModifyProtection: true,
                            readOnlyRecommended: fsDB.ReadOnlyRecommended);
                        if (!string.IsNullOrEmpty(options.ModifyPassword))
                            wbDB.Security.GrantModifyAccess(fsDB.ReadOnlyRecommended);
                    }
                    return wbDB;
                }
                EncryptionDetector.ThrowIfEncryptedOoxml(ms, "<stream>");
                ms.Position = 0;
                var sheets = XlsbBackend.ReadAll(ms);
                var wbB = Workbook.FromSheetData(sheets, null, ExcelFormat.Xlsb, null);
                ms.Position = 0;
                using (var zipB = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true))
                {
                    wbB.PreservedParts = OoxmlPreservedParts.Capture(zipB, sheets.Count, binary: true);
                    wbB.Properties.CopyFrom(XlsxReader.ReadProperties(zipB));
                }
                ms.Position = 0;
                wbB.VbaProjectBytes = XlsbBackend.ReadVbaProject(ms);
                ms.Position = 0;
                wbB.WorkbookCodeName = XlsbBackend.ReadWorkbookCodeName(ms);
                ms.Position = 0;
                wbB.Date1904 = XlsbBackend.ReadDate1904(ms);
                ms.Position = 0;
                ApplyFileSharing(wbB, options, XlsbBackend.ReadFileSharing(ms));
                return wbB;
            }
            default:
                throw new NotSupportedException($"{format} 读取后端尚未实现，当前仅支持 xlsx/xlsm/csv/xls/xlsb");
        }

        // xlsx/xlsm：单次解压内完成读表/读属性/捕获保留部件
        Workbook wb;
        OoxmlPreservedParts? preserved;
        {
            if (!string.IsNullOrEmpty(options.OpenPassword))
            {
                var decrypted = DecryptWithPasswordCheck(ms, "<stream>", options.OpenPassword!);
                using var zipD = new ZipArchive(decrypted, ZipArchiveMode.Read, leaveOpen: false);
                var sheetsD = XlsxReader.ReadAllRaw(zipD);
                var propsD = XlsxReader.ReadProperties(zipD);
                preserved = OoxmlPreservedParts.Capture(zipD, sheetsD.Count);
                preserved.WorkbookCodeName = XlsxReader.WorkbookCodeNameSnapshot;
                // P0-6: 命名区域与窗口视图原样回写；同时解析到 Workbook.Names
                preserved.BookViewsXml = XlsxReader.BookViewsXmlSnapshot;
                preserved.DefinedNamesXml = XlsxReader.DefinedNamesXmlSnapshot;
                wb = Workbook.FromSheetData(sheetsD, propsD, format, null);
                foreach (var nr in XlsxReader.ParseDefinedNames(preserved.DefinedNamesXml))
                    wb.Names.Add(nr);
                if (preserved.Parts.TryGetValue("xl/vbaProject.bin", out var vbaBytes))
                    wb.VbaProjectBytes = vbaBytes;
                wb.WorkbookCodeName = preserved.WorkbookCodeName;
                wb.Date1904 = XlsxReader.Date1904Snapshot;
            }
            else
            {
                EncryptionDetector.ThrowIfEncryptedOoxml(ms, "<stream>");
                ms.Position = 0;
                using var zip = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
                var sheets = XlsxReader.ReadAllRaw(zip);
                var props = XlsxReader.ReadProperties(zip);
                preserved = OoxmlPreservedParts.Capture(zip, sheets.Count);
                preserved.WorkbookCodeName = XlsxReader.WorkbookCodeNameSnapshot;
                // P0-6: 命名区域与窗口视图原样回写；同时解析到 Workbook.Names
                preserved.BookViewsXml = XlsxReader.BookViewsXmlSnapshot;
                preserved.DefinedNamesXml = XlsxReader.DefinedNamesXmlSnapshot;
                wb = Workbook.FromSheetData(sheets, props, format, null);
                foreach (var nr in XlsxReader.ParseDefinedNames(preserved.DefinedNamesXml))
                    wb.Names.Add(nr);
                if (preserved.Parts.TryGetValue("xl/vbaProject.bin", out var vbaBytes))
                    wb.VbaProjectBytes = vbaBytes;
                wb.WorkbookCodeName = preserved.WorkbookCodeName;
                wb.Date1904 = XlsxReader.Date1904Snapshot;
            }
        }
        wb.PreservedParts = preserved;
        wb.Protection = XlsxReader.WorkbookProtectionSnapshot;
        ApplyFileSharing(wb, options, XlsxReader.FileSharingSnapshot);

        if (options.FillMergedCells)
        {
            foreach (var ws in wb.Worksheets)
                ws.FillMergedValues();
        }

        return wb;
    }

    private static Workbook FinishOpen(IReadOnlyList<SheetData> sheets, WorkbookProperties? props,
        ExcelFormat format, string? path, ExcelReadOptions options)
    {
        var wb = Workbook.FromSheetData(sheets, props, format, path);
        if (options.FillMergedCells)
        {
            foreach (var ws in wb.Worksheets)
                ws.FillMergedValues();
        }
        return wb;
    }

    /// <summary>
    /// 提供打开密码时解密 CFB 加密工作簿。
    /// 若文件实际不是加密工作簿（非 CFB 容器），给出明确异常而非晦涩的 CFB 解析错误。
    /// </summary>
    private static Stream DecryptWithPasswordCheck(Stream fs, string path, string password)
    {
        if (!Internal.EncryptionDetector.IsEncryptedOoxml(fs, path))
        {
            throw new LiteExcelException(
                $"文件 '{path}' 不是加密工作簿（未检测到打开密码），无需提供 OpenPassword。请移除 ExcelReadOptions.OpenPassword 后重试。");
        }
        return Internal.Encryption.AgileDecryptor.Decrypt(fs, password);
    }

    /// <summary>根据 fileSharing（修改密码）信息设置工作簿安全状态 </summary>
    private static void ApplyFileSharing(Workbook wb, ExcelReadOptions options,
        Internal.Encryption.FileSharingInfo? fs)
    {
        wb.FileSharingToPreserve = fs;

        if (fs is null)
        {
            wb.Security.Initialize(options.OpenPassword);
            return;
        }

        // 提供修改密码即视为获得编辑授权（文件仅标记写保护，不承载强加密；
        // 哈希验证受不同 Excel 版本算法差异影响，采用"提供即授权"的保守策略）
        bool authorized = !string.IsNullOrEmpty(options.ModifyPassword);

        wb.Security.Initialize(options.OpenPassword, fileHasModifyProtection: true,
            readOnlyRecommended: fs.ReadOnlyRecommended);
        if (authorized)
            wb.Security.GrantModifyAccess(fs.ReadOnlyRecommended);
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

    /// <summary>
    /// 新建工作簿并直接写入 List&lt;T&gt; 数据（首个工作表，首行为表头）。
    /// 反射映射，已标注 DAM，AOT/裁剪安全；返回的工作簿可与样式/冻结/条件格式/密码等工作簿级能力混用。
    /// </summary>
    public static Workbook Create<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IEnumerable<T> data, string sheetName = "Sheet1", ExcelFormat format = ExcelFormat.Xlsx,
        Action<WriteOptions<T>>? configure = null)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        var wb = Workbook.CreateEmpty(format);
        var options = new WriteOptions<T> { SheetName = sheetName };
        configure?.Invoke(options);
        var ws = Worksheet.FromSheetData(XlsxWriter.ListToSheet(data, options));
        wb.Worksheets.AddInternal(ws);
        wb.OnWorksheetAdded(ws);
        return wb;
    }

    /// <summary>
    /// 新建工作簿并直接写入 DataTable 数据（首个工作表，首行写列名）。sheetName 为空时用 DataTable.TableName，再为空则 Sheet1。
    /// </summary>
    public static Workbook Create(DataTable table, string? sheetName = null, ExcelFormat format = ExcelFormat.Xlsx)
    {
        if (table is null) throw new ArgumentNullException(nameof(table));
        var name = sheetName;
        if (string.IsNullOrWhiteSpace(name))
            name = string.IsNullOrWhiteSpace(table.TableName) ? "Sheet1" : table.TableName;
        var wb = Workbook.CreateEmpty(format);
        var ws = Worksheet.FromSheetData(XlsxWriter.DataTableToSheet(table, name));
        wb.Worksheets.AddInternal(ws);
        wb.OnWorksheetAdded(ws);
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

        // 根据扩展名推断目标格式，与 DetectFormat 完全一致：扩展名与工作簿格式冲突时以扩展名为准
        var extFormat = DetectFormat(path);
        format = extFormat;

        ApplyWriteOptions(workbook, options);
        // 批次 0：注入能力降级回调（默认 null，不注册则行为与历史一致）
        workbook.DegradationCallback = options.OnDegradation;
        // 批次 P1-A：注入 CSV 写出分隔符（默认 null → 逗号）
        workbook.WriteSeparator = options.Separator;
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
        var ws = Worksheet.FromSheetData(XlsxWriter.DataTableToSheet(table, sheetName));
        wb.Worksheets.AddInternal(ws);
        wb.OnWorksheetAdded(ws);
        Write(path, wb, options);
    }

    /// <summary>写出 List&lt;T&gt;（反射映射，已标注 DAM，AOT/裁剪安全） </summary>
    public static void Write<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        string path, IEnumerable<T> data, string sheetName = "Sheet1", Action<WriteOptions<T>>? configure = null)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        XlsxWriter.Write(path, data, opt =>
        {
            opt.SheetName = sheetName;
            configure?.Invoke(opt);
        });
    }

    // ── 读取便利 ──

    /// <summary>读取指定工作表为 List&lt;T&gt;（反射映射，已标注 DAM，AOT/裁剪安全）。默认第一张表，首行作为表头 </summary>
    public static List<T> Read<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties
            | DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] T>(
        string path, string? sheetName = null, Action<ReadOptions<T>>? configure = null) where T : new()
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
        var format = DetectFormat(path);
        // xlsx / xlsm：XlsxReader 可直接读 workbook.xml 元数据列出表名（轻量）。
        // xlsb / xls / csv：非 Excel XML 容器 → 走 Excel.Open 路由对应后端。
        if (format == ExcelFormat.Xlsx || format == ExcelFormat.Xlsm)
            return XlsxReader.GetSheetNames(path);
        var wb = Excel.Open(path);
        return wb.Worksheets.Names.ToList();
    }

    /// <summary>从流列出所有工作表名。仅支持 zip 容器格式（xlsx/xlsm）。 </summary>
    public static List<string> GetSheetNames(Stream stream)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        return XlsxReader.GetSheetNames(stream);
    }

    /// <summary>流式读取指定工作表，逐行回调，不驻留内存（自动跳过首行） </summary>
    public static void StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        if (onRow is null) throw new ArgumentNullException(nameof(onRow));
        EnsureXlsxStreamingFormat(path, "流式读取");
        XlsxReader.StreamRows(path, sheetName, onRow);
    }

    /// <summary>
    /// 拉取式流式读取，逐行 yield，支持 LINQ 与提前中断，不驻留内存。
    /// <para><paramref name="sheetName"/> 为 null 时取第一张表。</para>
    /// <para>与 <see cref="StreamRows(string, string, Action{IReadOnlyList{Cell}})"/> 的区别：拉取模型（IEnumerable）、不跳过首行、支持提前中断。</para>
    /// </summary>
    public static IEnumerable<IReadOnlyList<Cell>> EnumerateRows(string path, string? sheetName = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        EnsureXlsxStreamingFormat(path, "流式读取");
        return XlsxReader.EnumerateRows(path, sheetName);
    }

    /// <summary>拉取式流式读取（Stream 重载），逐行 yield，支持 LINQ 与提前中断 </summary>
    public static IEnumerable<IReadOnlyList<Cell>> EnumerateRows(Stream stream, string? sheetName = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        return XlsxReader.EnumerateRows(stream, sheetName);
    }

    /// <summary>创建流式写入器（逐行写大文件，不驻留内存）。使用后调用 Dispose/Close 完成文件。超出单表行数上限时按 <paramref name="onRowLimitExceeded"/> 处理（默认抛异常）。<paramref name="spillHeader"/> 仅在 SpillToNewSheet 下生效，作为每张表首行表头 </summary>
    public static XlsxStreamWriter CreateWriter(string path, RowLimitExceededMode onRowLimitExceeded = RowLimitExceededMode.Throw, object?[]? spillHeader = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        var format = DetectFormat(path);
        if (format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsm)
            throw new LiteExcelException($"该格式不支持流式写入：{format}。请使用 .xlsx 或 .xlsm 扩展名。");
        return XlsxStreamWriter.Create(path, onRowLimitExceeded, spillHeader);
    }

    /// <summary>创建流式写入器（写入流，LeaveOpen 由调用方管理）。超出单表行数上限时按 <paramref name="onRowLimitExceeded"/> 处理（默认抛异常）。<paramref name="spillHeader"/> 仅在 SpillToNewSheet 下生效，作为每张表首行表头 </summary>
    public static XlsxStreamWriter CreateWriter(Stream stream, RowLimitExceededMode onRowLimitExceeded = RowLimitExceededMode.Throw, object?[]? spillHeader = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        return XlsxStreamWriter.Create(stream, onRowLimitExceeded, spillHeader);
    }

    /// <summary>追加数据到已有文件（同名表合并列后追加行；文件不存在则创建） </summary>
    public static void Append(string path, SheetData newData, WorkbookProperties? updateProperties = null)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("路径不能为空", nameof(path));
        XlsxWriter.Append(path, newData, updateProperties);
    }

    /// <summary>带进度读取指定工作表。current 从 1 递增到 total（数据行数，不含表头） </summary>
    public static void ReadWithProgress(string path, int sheetIndex, Action<int, int> onProgress)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("路径不能为空", nameof(path));
        XlsxReader.ReadWithProgress(path, sheetIndex, onProgress);
    }

    private static void EnsureXlsxStreamingFormat(string path, string operation)
    {
        var format = DetectFormat(path);
        if (format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsm)
            throw new LiteExcelException($"该格式不支持{operation}：{format}。仅支持 xlsx/xlsm。");
    }

    /// <summary>把 XlsBackend 快照到的命名区域挂到工作簿（xls 打开后 Names 自动填充）。</summary>
    private static void AttachXlsNames(Workbook wb)
    {
        var names = XlsBackend.DefinedNamesSnapshot;
        if (names is null) return;
        foreach (var nr in names)
            wb.Names.Add(nr);
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
}
