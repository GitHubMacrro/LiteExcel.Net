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

    /// <summary>能力降级回调（由 Excel.Write 注入 options.OnDegradation；直接 SaveAs 时为 null，行为与历史一致） </summary>
    internal Action<DegradationInfo>? DegradationCallback { get; set; }

    /// <summary>CSV 写出分隔符（由 Excel.Write 注入；为 null 时使用逗号） </summary>
    internal char? WriteSeparator { get; set; }

    /// <summary>CSV 写出编码（由 Excel.Write 注入；为 null 时使用 UTF-8 带 BOM） </summary>
    internal System.Text.Encoding? WriteEncoding { get; set; }

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

    /// <summary>是否使用 1904 日期系统（Excel 序列值基准为 1904-01-01）。打开时捕获；保存时写回对应格式标志。 </summary>
    internal bool Date1904 { get; set; }

    /// <summary>文件级安全状态（打开密码 / 修改密码 / 只读与保存权限） </summary>
    public WorkbookSecurity Security { get; }

    /// <summary>工作簿保护（workbookProtection：锁结构/窗口）。默认 null（无保护） </summary>
    public WorkbookProtection? Protection { get; set; }

    /// <summary>命名区域（definedNames，全局 + sheet-local）。读取打开文件时自动填充。 </summary>
    public List<NamedRange> Names { get; } = new();

    /// <summary>打开时捕获的原 fileSharing（修改密码哈希），保存时透传保留。用户显式设置新修改密码时失效 </summary>
    internal Internal.Encryption.FileSharingInfo? FileSharingToPreserve { get; set; }

    /// <summary>源文件是否含透视表（XLS 检测 SXVIEW 记录）。含透视表时默认阻止保存，因为当前模型无法保真写回 BIFF8 透视表。
    /// 用户可通过 <see cref="AllowFeatureLossOnSave"/> 显式允许降级写出。 </summary>
    internal bool SourceHasPivotTables { get; set; }

    internal bool SourceHasAdvancedXlsbParts { get; set; }

    internal HashSet<int> AdvancedXlsbSheetIndexes { get; } = new();

    /// <summary>是否允许保存时丢失不支持的高级功能（如 BIFF8 透视表）。默认 false：含透视表的 XLS 保存被阻止。
    /// 用户显式设为 true 后允许保存，但透视表等不可保真能力会被丢弃，并通过降级回调上报。 </summary>
    public bool AllowFeatureLossOnSave { get; set; }

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
        Security = new WorkbookSecurity();
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
        ThrowIfReadOnly();
        if (string.IsNullOrEmpty(_currentPath))
            throw new LiteExcelException("当前工作簿没有目标路径，请使用 SaveAs 指定保存位置");
        SaveCore(_currentPath, Format);
    }

    /// <summary>另存为指定路径。格式沿用当前格式 </summary>
    public void SaveAs(string path)
    {
        ThrowIfReadOnly();
        SaveAs(path, Format);
    }

    /// <summary>另存为指定路径并指定格式（格式必须为已支持的可写格式）。路径扩展名必须与 format 匹配，否则抛 <see cref="LiteExcelException"/> </summary>
    public void SaveAs(string path, ExcelFormat format)
    {
        ThrowIfReadOnly();
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
        ThrowIfReadOnly();
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("流不可写", nameof(stream));

        SaveCore(stream, format);
    }

    /// <summary>只读工作簿（有修改密码但未授权）禁止保存 </summary>
    private void ThrowIfReadOnly()
    {
        if (!Security.CanSave)
            throw new LiteExcelException(
                "当前工作簿以只读方式打开（文件设置了修改密码，但未提供正确的修改密码），不能保存。" +
                "请通过 Excel.Open 的 ExcelReadOptions.ModifyPassword 提供正确的修改密码，或使用无修改保护的工作簿。");
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
        // 在创建目标文件前完成格式能力校验。
        ThrowIfMacroNotSupported(format);
        // 在创建目标文件前阻止无法保真的 XLS 透视表保存。
        ThrowIfPivotTablesNotPreservable(format);
        ThrowIfAdvancedXlsbPartsNotPreservable(format);

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        SaveCore(fs, format);
    }

    private void SaveCore(Stream stream, ExcelFormat format)
    {
        // 流写出路径同样执行格式能力校验。
        ThrowIfMacroNotSupported(format);
        // 文件级密码仅支持 xlsx/xlsm/xlsb；csv/xls 不支持加密写出
        ThrowIfPasswordNotSupported(format);
        // 阻止无法保真的 BIFF8 透视表保存。
        ThrowIfPivotTablesNotPreservable(format);
        ThrowIfAdvancedXlsbPartsNotPreservable(format);

        switch (format)
        {
            case ExcelFormat.Xlsx:
            case ExcelFormat.Xlsm:
            {
                var sheets = BuildSheetDataList();
                bool structureUnchanged = StructureUnchanged(sheets);
                bool verbatimX = CanVerbatimXlsx(sheets);
                var openPwd = Security.GetOpenPassword();
                var (fsHash, fsSalt, fsSpin, fsRo) = BuildFileSharingParams();
                if (!string.IsNullOrEmpty(openPwd))
                {
                    // 先写入 ZIP，再封装为加密 CFB。
                    using var zipMs = new MemoryStream();
                    XlsxWriter.Write(zipMs, sheets, Properties, PreservedParts, mergeSheetRels: structureUnchanged,
                        macroEnabled: format == ExcelFormat.Xlsm, date1904: Date1904,
                        fileSharingHash: fsHash, fileSharingSalt: fsSalt, fileSharingSpin: fsSpin, fileSharingReadOnlyRecommended: fsRo,
                        workbookProtection: Protection, degradationCallback: DegradationCallback, verbatim: verbatimX);
                    zipMs.Position = 0;
                    var encrypted = Internal.Encryption.OoxmlEncryptor.Encrypt(zipMs.ToArray(), openPwd);
                    stream.Write(encrypted, 0, encrypted.Length);
                }
                else
                {
                    XlsxWriter.Write(stream, sheets, Properties, PreservedParts, mergeSheetRels: structureUnchanged,
                        macroEnabled: format == ExcelFormat.Xlsm, date1904: Date1904,
                        fileSharingHash: fsHash, fileSharingSalt: fsSalt, fileSharingSpin: fsSpin, fileSharingReadOnlyRecommended: fsRo,
                        workbookProtection: Protection, degradationCallback: DegradationCallback, verbatim: verbatimX);
                }
                break;
            }
            case ExcelFormat.Csv:
                if (Worksheets.Count != 1)
                    throw new NotSupportedException("CSV 仅支持单工作表工作簿");
                CsvBackend.Write(stream, Worksheets[0].ToSheetData(), DegradationCallback, ExcelFormat.Csv, WriteSeparator, WriteEncoding);
                break;
            case ExcelFormat.Xls:
            {
                var xlsSheets = BuildSheetDataList();
                XlsWriter.Write(stream, xlsSheets, Date1904, DegradationCallback, ExcelFormat.Xls,
                    Names, Properties);
                break;
            }
            case ExcelFormat.Xlsb:
            {
                var xlsbSheets = BuildSheetDataList();
                var openPwdB = Security.GetOpenPassword();
                var (fsHashB, fsSaltB, fsSpinB, fsRoB) = BuildFileSharingParams();
                bool verbatimB = CanVerbatimXlsb(xlsbSheets);
                if (!string.IsNullOrEmpty(openPwdB))
                {
                    using var zipMs = new MemoryStream();
                    XlsbWriter.Write(zipMs, xlsbSheets, VbaProjectBytes, WorkbookCodeName, Date1904,
                        fsHashB, fsSaltB, fsSpinB, fsRoB, DegradationCallback, ExcelFormat.Xlsb,
                        PreservedParts, Properties, Names, verbatim: verbatimB);
                    zipMs.Position = 0;
                    var encrypted = Internal.Encryption.OoxmlEncryptor.Encrypt(zipMs.ToArray(), openPwdB);
                    stream.Write(encrypted, 0, encrypted.Length);
                }
                else
                {
                    XlsbWriter.Write(stream, xlsbSheets, VbaProjectBytes, WorkbookCodeName, Date1904,
                        fsHashB, fsSaltB, fsSpinB, fsRoB, DegradationCallback, ExcelFormat.Xlsb,
                        PreservedParts, Properties, Names, verbatim: verbatimB);
                }
                break;
            }
            default:
                throw new NotSupportedException($"未知格式：{format}");
        }
    }

    /// <summary>含 VBA 宏的工作簿不允许保存为不支持宏的格式（xlsx/xls），防止宏静默丢失或生成不一致文件 </summary>
    private void ThrowIfMacroNotSupported(ExcelFormat format)
    {
        if (VbaProjectBytes is not null && (format == ExcelFormat.Xls || format == ExcelFormat.Xlsx))
            throw new LiteExcelException(
                $"无法写出 {format}：当前工作簿包含 VBA 宏，而 {format} 格式不支持宏。" +
                "请另存为 .xlsm 或 .xlsb 以保留宏。");
    }

    /// <summary>文件级密码（打开/修改）仅支持 xlsx/xlsm/xlsb；csv/xls 不支持加密写出 </summary>
    private void ThrowIfPasswordNotSupported(ExcelFormat format)
    {
        if (!Security.HasOpenPassword && !Security.HasModifyPassword)
            return;
        if (format == ExcelFormat.Csv || format == ExcelFormat.Xls)
            throw new LiteExcelException(
                $"无法写出 {format}：{format} 格式不支持文件级密码（打开密码/修改密码）。" +
                "请使用 xlsx/xlsm/xlsb 保存，或先移除密码。");
    }

    /// <summary>源 XLS 含透视表时默认阻止保存（BIFF8 透视表无法保真写回或转换到其他格式）。
    /// 用户可设 <see cref="AllowFeatureLossOnSave"/> = true 显式允许降级，此时透视表会被丢弃并经降级回调上报。 </summary>
    private void ThrowIfPivotTablesNotPreservable(ExcelFormat format)
    {
        if (!SourceHasPivotTables) return;
        if (AllowFeatureLossOnSave)
        {
            // 显式允许降级时上报并继续写出。
            DegradationCallback?.Invoke(new DegradationInfo
            {
                Capability = DegradationCapability.PivotTables,
                TargetFormat = format,
                Message = $"源 XLS 文件包含透视表，当前版本无法保真写回或转换 BIFF8 透视表到 {format} 格式。" +
                    "透视表将被丢弃（数据保留，透视视图丢失）。",
            });
            return;
        }
        throw new LiteExcelException(
            $"源 XLS 文件包含透视表，当前版本无法保真写回或转换 BIFF8 透视表，保存会永久删除透视表。默认已阻止本次保存。\n" +
            "如确认接受功能丢失，请设 workbook.AllowFeatureLossOnSave = true 后重试。");
    }

    private void ThrowIfAdvancedXlsbPartsNotPreservable(ExcelFormat format)
    {
        if (Format != ExcelFormat.Xlsb || !SourceHasAdvancedXlsbParts)
            return;
        bool modified = AdvancedXlsbSheetIndexes.Count == 0
            ? Worksheets.Any(ws => ws.IsModified)
            : AdvancedXlsbSheetIndexes.Any(index => index >= 0 && index < Worksheets.Count && Worksheets[index].IsModified);
        if (!modified)
            return;
        if (AllowFeatureLossOnSave)
        {
            DegradationCallback?.Invoke(new DegradationInfo
            {
                Capability = DegradationCapability.PivotTables,
                TargetFormat = format,
                Message = $"源 XLSB 文件包含当前模型无法安全合并的高级部件，保存到 {format} 时这些部件可能丢失。"
            });
            return;
        }
        throw new LiteExcelException(
            "源 XLSB 文件包含透视表、切片器、图表或其他高级部件，当前版本无法在编辑后安全合并这些部件。默认已阻止本次保存。\n" +
            "如确认接受功能丢失，请设 workbook.AllowFeatureLossOnSave = true 后重试。");
    }

    /// <summary>
    /// 生成 fileSharing（修改密码）写出参数。
    /// 优先透传打开时捕获的原 fileSharing（未改动修改密码时）；否则从 Security 的修改密码重新生成。
    /// 返回 (hash, salt, spin, readOnlyRecommended)；无修改密码时 hash 为 null。
    /// </summary>
    private (string? hash, string? salt, int? spin, bool readOnlyRecommended) BuildFileSharingParams()
    {
        // 用户显式设置了修改密码：重新生成
        var modifyPwd = Security.GetModifyPassword();
        if (!string.IsNullOrEmpty(modifyPwd))
        {
            var salt = new byte[16];
            using (var rng = System.Security.Cryptography.RandomNumberGenerator.Create())
                rng.GetBytes(salt);
            var hash = Internal.Encryption.FileSharingInfo.ComputeHash(modifyPwd, salt);
            return (Convert.ToBase64String(hash), Convert.ToBase64String(salt), 100000, Security.ReadOnlyRecommended);
        }

        // 修改密码被主动变更时不保留原 fileSharing。
        if (Security.ModifyPasswordTouched)
            return (null, null, null, false);

        // 保留打开时捕获的 fileSharing。
        var preserved = FileSharingToPreserve;
        if (preserved is not null)
            return (Convert.ToBase64String(preserved.HashValue),
                preserved.SaltValue is null ? null : Convert.ToBase64String(preserved.SaltValue),
                preserved.SpinCount > 0 ? preserved.SpinCount : null,
                preserved.ReadOnlyRecommended);

        return (null, null, null, false);
    }

    internal List<SheetData> BuildSheetDataList()
    {
        var list = new List<SheetData>(Worksheets.Count);
        foreach (var ws in Worksheets)
            list.Add(ws.ToSheetData());
        return list;
    }

    /// <summary>
    /// 工作表数量相对打开时是否未变（决定能否复用工作表级保留 rels）。
    /// 只比较数量而非表名：表名仅存在于 workbook.xml，不影响 sheet{i}.xml 与其 rels 的绑定；
    /// 改表名不应导致 drawing/图表关联被丢弃。
    /// 注意：重排（Move）后同位置 sheet rels 可能错配，属低频场景，保留比删除更安全。
    /// </summary>
    private bool StructureUnchanged(List<SheetData> sheets)
    {
        return _openedSheetNames is not null && sheets.Count == _openedSheetNames.Count;
    }

    /// <summary>
    /// 判断是否可对 XLSB 做 verbatim 保留（原样写出原始二进制部件而非重建）。
    /// 条件：源格式为 XLSB + 结构不变（表数+表名）+ 无工作表修改 + 原始二进制部件已捕获。
    /// </summary>
    private bool CanVerbatimXlsb(List<SheetData> sheets)
    {
        if (Format != ExcelFormat.Xlsb) return false;
        if (_openedSheetNames is null) return false;
        if (sheets.Count != _openedSheetNames.Count) return false;
        for (int i = 0; i < sheets.Count; i++)
            if (!string.Equals(sheets[i].SheetName, _openedSheetNames[i], System.StringComparison.Ordinal))
                return false;
        foreach (var ws in Worksheets)
            if (ws.IsModified)
                return false;
        if (PreservedParts?.VerbatimBinaries is null) return false;
        if (!PreservedParts.VerbatimBinaries.ContainsKey("xl/workbook.bin")) return false;
        if (!PreservedParts.VerbatimBinaries.ContainsKey("xl/styles.bin")) return false;
        // 修改密码变动时需重建包含 BrtFileSharingIso 记录的 workbook.bin。
        if (Security.ModifyPasswordTouched) return false;
        if (Security.HasModifyPassword) return false;
        return true;
    }

    /// <summary>
    /// 判断是否可对 XLSX/XLSM 做 verbatim 保留（原样写出原始 styles.xml / sharedStrings.xml / sheetN.xml）。
    /// 条件：源格式为 xlsx/xlsm + 结构不变（表数+表名）+ 无工作表修改 + 原始 XML 部件已捕获
    ///     + 原始 styles 含扩展内容（否则重建路径更干净，避免改变既有简化文件的行为）。
    /// 用于保留 slicerStyles / timelineStyles / pivotButton XF 等扩展样式，避免重建时样式索引变化导致透视表/切片器渲染失败。
    /// </summary>
    private bool CanVerbatimXlsx(List<SheetData> sheets)
    {
        if (Format != ExcelFormat.Xlsx && Format != ExcelFormat.Xlsm) return false;
        if (_openedSheetNames is null) return false;
        if (sheets.Count != _openedSheetNames.Count) return false;
        for (int i = 0; i < sheets.Count; i++)
            if (!string.Equals(sheets[i].SheetName, _openedSheetNames[i], System.StringComparison.Ordinal))
                return false;
        foreach (var ws in Worksheets)
            if (ws.IsModified)
                return false;
        if (PreservedParts?.VerbatimXmlParts is null) return false;
        if (!PreservedParts.VerbatimXmlParts.TryGetValue("xl/styles.xml", out var rawStyles) || rawStyles is null)
            return false;
        if (!HasExtendedStyles(rawStyles)) return false;
        for (int i = 1; i <= sheets.Count; i++)
            if (!PreservedParts.VerbatimXmlParts.ContainsKey($"xl/worksheets/sheet{i}.xml"))
                return false;
        // 修改密码变动时需重建包含 fileSharing 的 workbook.xml。
        if (Security.ModifyPasswordTouched) return false;
        if (Security.HasModifyPassword) return false;
        return true;
    }

    /// <summary>原始 styles.xml 是否含扩展样式内容（extLst 切片器/时间线样式、pivotButton XF、自定义 XF/numFmt 等），
    /// 是则值得 verbatim 保留；否则重建路径（支持稀疏写出等简化）更合适。 </summary>
    private static bool HasExtendedStyles(byte[] stylesXml)
    {
        var text = System.Text.Encoding.UTF8.GetString(stylesXml);
        // extLst（slicerStyles/timelineStyles）、pivotButton XF、自定义 cellStyleXfs 都表明文件有扩展样式
        if (text.IndexOf("extLst", StringComparison.Ordinal) >= 0) return true;
        if (text.IndexOf("pivotButton", StringComparison.Ordinal) >= 0) return true;
        if (text.IndexOf("slicerStyles", StringComparison.Ordinal) >= 0) return true;
        if (text.IndexOf("timelineStyles", StringComparison.Ordinal) >= 0) return true;
        return false;
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
