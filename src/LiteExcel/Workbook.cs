using LiteExcel.Internal;
using LiteExcel.Internal.Biff;
using System.IO;
using System.Linq;

namespace LiteExcel;

/// <summary>
/// 统一工作簿API,负责工作表集合、文档属性、保存/另存为,打开时加载到内存，不长期持有文件流，因此无需 IDisposable。
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
    /// 打开时捕获的、写入器不重建的 OOXML 部件（宏/主题/绘图/图表等），保存时按二进制透传，避免未映射部件被静默删除。新建工作簿为 null
    /// </summary>
    internal OoxmlPreservedParts? PreservedParts { get; set; }

    /// <summary>
    /// 打开时捕获的 VBA 宏工程原始字节（xl/vbaProject.bin）,写入 xlsb 时透传保留,新建工作簿或源文件无宏时为 null
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

    /// <summary>源 xlsb 含尚无转码器覆盖的高级部件（如时间线 timelineCaches/timelines），
    /// 跨格式转换时无法保留，须显式上报/阻止。</summary>
    internal bool SourceHasUncoveredAdvancedXlsbParts { get; set; }

    internal HashSet<int> AdvancedXlsbSheetIndexes { get; } = new();

    /// <summary>是否允许保存时丢失不支持的高级功能（如 VBA 宏、BIFF8 透视表、xlsb 高级部件）。默认 true：不受支持的能力会被丢弃但经 <see cref="SaveDegradations"/> 记录（非静默）。
    /// 用户显式设为 false 后，含此类能力且目标格式不支持时保存被阻止并抛 <see cref="LiteExcelException"/>。 </summary>
    public bool AllowFeatureLossOnSave { get; set; } = true;

    /// <summary>上一次保存期间被丢弃的能力清单（每次 Save/SaveAs 前清空、保存中累积）。
    /// 用于在默认放行下仍能获知"丢失了什么"，避免静默数据丢失。 </summary>
    public IReadOnlyList<DegradationInfo> SaveDegradations => _saveDegradations;

    private readonly List<DegradationInfo> _saveDegradations = new();

    /// <summary>
    /// 当前目标路径
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
        for (int i = 0; i < sheets.Count; i++)
        {
            var sheet = sheets[i];
            // 记录打开时的 0-based 序号：删除/移动表后用于从 preserved 复用该表原始 rels
            sheet.OrigIndex = i;
            var ws = Worksheet.FromSheetData(sheet);
            wb.Worksheets.AddInternal(ws);
            wb.OnWorksheetAdded(ws);
        }
        return wb;
    }

    /// <summary>保存到当前目标路径。若当前无路径（新建），抛出 <see cref="LiteExcelException"/> </summary>
    public void Save()
    {
        ThrowIfReadOnly();
        if (string.IsNullOrEmpty(_currentPath))
            throw new LiteExcelException("当前工作簿没有目标路径，请使用 SaveAs 指定保存位置");
        _saveDegradations.Clear();
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
        _saveDegradations.Clear();
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

        _saveDegradations.Clear();
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
        // 在创建目标文件前完成全部格式能力校验。
        ThrowIfMacroNotSupported(format);
        ThrowIfPasswordNotSupported(format);
        ThrowIfPivotTablesNotPreservable(format);
        ThrowIfAdvancedXlsbPartsNotPreservable(format);

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        WriteTo(fs, format);
    }

    private void SaveCore(Stream stream, ExcelFormat format)
    {
        ThrowIfMacroNotSupported(format);
        ThrowIfPasswordNotSupported(format);
        ThrowIfPivotTablesNotPreservable(format);
        ThrowIfAdvancedXlsbPartsNotPreservable(format);
        WriteTo(stream, format);
    }

    /// <summary>执行实际写出的 switch 分发，避免 path→stream 双重执行导致降级重复上报</summary>
    private void WriteTo(Stream stream, ExcelFormat format)
    {
        void OnDeg(DegradationInfo info)
        {
            _saveDegradations.Add(info);
            DegradationCallback?.Invoke(info);
        }

        switch (format)
        {
            case ExcelFormat.Xlsx:
            case ExcelFormat.Xlsm:
            {
                var sheets = BuildSheetDataList();
                // B1：源为 xlsb 时，把模型中的定义名（可解码的简单引用）与数据模型链接表名合并为 definedNames 写出。
                if (Format == ExcelFormat.Xlsb && PreservedParts is not null)
                    PreservedParts.DefinedNamesXml = BuildXlsbDefinedNamesXml();
                bool structureUnchanged = StructureUnchanged(sheets);
                bool verbatimX = CanVerbatimXlsx(sheets);
                bool surgicalX = CanSurgicalXlsx(sheets);
                bool dropMacros = format == ExcelFormat.Xlsx && VbaProjectBytes is not null;
                var openPwd = Security.GetOpenPassword();
                var (fsHash, fsSalt, fsSpin, fsRo) = BuildFileSharingParams();
                if (!string.IsNullOrEmpty(openPwd))
                {
                    // 先写入 ZIP，再封装为加密 CFB。
                    using var zipMs = new MemoryStream();
                    XlsxWriter.Write(zipMs, sheets, Properties, PreservedParts, mergeSheetRels: structureUnchanged,
                        macroEnabled: format == ExcelFormat.Xlsm, date1904: Date1904,
                        fileSharingHash: fsHash, fileSharingSalt: fsSalt, fileSharingSpin: fsSpin, fileSharingReadOnlyRecommended: fsRo,
                        workbookProtection: Protection, degradationCallback: OnDeg, verbatim: verbatimX, surgical: surgicalX, dropMacros: dropMacros);
                    zipMs.Position = 0;
                    var encrypted = Internal.Encryption.OoxmlEncryptor.Encrypt(zipMs.ToArray(), openPwd);
                    stream.Write(encrypted, 0, encrypted.Length);
                }
                else
                {
                    XlsxWriter.Write(stream, sheets, Properties, PreservedParts, mergeSheetRels: structureUnchanged,
                        macroEnabled: format == ExcelFormat.Xlsm, date1904: Date1904,
                        fileSharingHash: fsHash, fileSharingSalt: fsSalt, fileSharingSpin: fsSpin, fileSharingReadOnlyRecommended: fsRo,
                        workbookProtection: Protection, degradationCallback: OnDeg, verbatim: verbatimX, surgical: surgicalX, dropMacros: dropMacros);
                }
                break;
            }
            case ExcelFormat.Csv:
                if (Worksheets.Count != 1)
                    throw new NotSupportedException("CSV 仅支持单工作表工作簿");
                CsvBackend.Write(stream, Worksheets[0].ToSheetData(), OnDeg, ExcelFormat.Csv, WriteSeparator, WriteEncoding);
                break;
            case ExcelFormat.Xls:
            {
                var xlsSheets = BuildSheetDataList();
                XlsWriter.Write(stream, xlsSheets, Date1904, OnDeg, ExcelFormat.Xls,
                    Names, Properties);
                break;
            }
            case ExcelFormat.Xlsb:
            {
                var xlsbSheets = BuildSheetDataList();
                var openPwdB = Security.GetOpenPassword();
                var (fsHashB, fsSaltB, fsSpinB, fsRoB) = BuildFileSharingParams();
                bool verbatimB = CanVerbatimXlsb(xlsbSheets);
                bool surgicalB = CanSurgicalXlsb(xlsbSheets);
                if (!string.IsNullOrEmpty(openPwdB))
                {
                    using var zipMs = new MemoryStream();
                    XlsbWriter.Write(zipMs, xlsbSheets, VbaProjectBytes, WorkbookCodeName, Date1904,
                        fsHashB, fsSaltB, fsSpinB, fsRoB, OnDeg, ExcelFormat.Xlsb,
                        PreservedParts, Properties, Names, verbatim: verbatimB, surgical: surgicalB, allowFeatureLoss: AllowFeatureLossOnSave);
                    zipMs.Position = 0;
                    var encrypted = Internal.Encryption.OoxmlEncryptor.Encrypt(zipMs.ToArray(), openPwdB);
                    stream.Write(encrypted, 0, encrypted.Length);
                }
                else
                {
                    XlsbWriter.Write(stream, xlsbSheets, VbaProjectBytes, WorkbookCodeName, Date1904,
                        fsHashB, fsSaltB, fsSpinB, fsRoB, OnDeg, ExcelFormat.Xlsb,
                        PreservedParts, Properties, Names, verbatim: verbatimB, surgical: surgicalB, allowFeatureLoss: AllowFeatureLossOnSave);
                }
                break;
            }
            default:
                throw new NotSupportedException($"未知格式：{format}");
        }
    }

    /// <summary>含 VBA 宏的工作簿保存为不支持宏的格式（xlsx/xls）时：默认放行并上报（宏被剥离），
    /// 严格模式（AllowFeatureLossOnSave=false）抛异常阻止。</summary>
    private void ThrowIfMacroNotSupported(ExcelFormat format)
    {
        if (VbaProjectBytes is null || (format != ExcelFormat.Xls && format != ExcelFormat.Xlsx))
            return;
        if (AllowFeatureLossOnSave)
        {
            ReportDegradation(new DegradationInfo
            {
                Capability = DegradationCapability.Macros,
                TargetFormat = format,
                Message = $"源工作簿包含 VBA 宏，而 {format} 格式不支持宏，宏代码将被剥离。调用了 VBA 自定义函数的公式在 Excel 中会显示 #NAME?。",
            });
            return;
        }
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

    /// <summary>源 XLS 含透视表时默认阻止保存
    /// 用户可设 <see cref="AllowFeatureLossOnSave"/> = true 显式允许降级，此时透视表会被丢弃并经降级回调上报。 </summary>
    private void ThrowIfPivotTablesNotPreservable(ExcelFormat format)
    {
        if (!SourceHasPivotTables) return;
        if (AllowFeatureLossOnSave)
        {
            // 显式允许降级时上报并继续写出。
            ReportDegradation(new DegradationInfo
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

    /// <summary>
    /// 源为 xlsb 时构建 workbook.xml 的 definedNames：
    /// 模型中的定义名（rgce 可解码的简单引用/常量）+ 数据模型链接表的 `_xlcn.LinkedTable_*`（由连接合成）。
    /// 复合表达式（结构化引用、函数等）在读取时已跳过，不会产出错误引用。
    /// </summary>
    private string BuildXlsbDefinedNamesXml()
    {
        var sb = new System.Text.StringBuilder();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var n in Names)
        {
            if (string.IsNullOrEmpty(n.Name) || string.IsNullOrEmpty(n.Reference)) continue;
            if (!seen.Add(n.Name + "|" + n.LocalSheetId)) continue;
            var local = n.LocalSheetId >= 0 ? $" localSheetId=\"{n.LocalSheetId}\"" : "";
            sb.Append($"<definedName name=\"{XlsxWriter.XmlEscape(n.Name)}\"{local}>{XlsxWriter.XmlEscape(n.Reference)}</definedName>");
        }
        if (PreservedParts?.Parts.TryGetValue("xl/connections.bin", out var connBin) == true)
        {
            foreach (var c in Internal.Biff12.XlsbConnectionTranscoder.Parse(connBin))
            {
                if (c.Type != 102 || string.IsNullOrEmpty(c.RangeSourceName)) continue;
                var nm = c.RangeSourceName! + "1";
                if (!seen.Add(nm + "|-1")) continue;
                sb.Append($"<definedName name=\"{XlsxWriter.XmlEscape(nm)}\" hidden=\"1\">{XlsxWriter.XmlEscape((c.X15Id ?? "") + "[]")}</definedName>");
            }
        }
        return sb.Length > 0 ? "<definedNames>" + sb + "</definedNames>" : "";
    }

    /// <summary>
    /// 源 XLSB 含高级部件（透视表/切片器/连接/PQ/数据模型等）时的守卫：
    /// - 目标仍为 xlsb：verbatim / surgical 可逐字节保留则放行，否则按既有逻辑上报或阻止。
    /// - 跨格式（目标 xlsx/xlsm/xls/csv）：这些 BIFF12 部件尚无法转码，写出时会被丢弃 —— 显式上报或阻止，绝不静默。
    /// </summary>
    private void ThrowIfAdvancedXlsbPartsNotPreservable(ExcelFormat format)
    {
        if (Format != ExcelFormat.Xlsb || !SourceHasAdvancedXlsbParts)
            return;

        if (format == ExcelFormat.Xlsb)
        {
            var sheets = BuildSheetDataList();
            if (CanVerbatimXlsb(sheets))
                return;
            // 手术式删除通道：仅删除若干工作表，其余二进制部件原样保留（含数据模型/透视/连接等全部高级部件）。
            if (CanSurgicalXlsb(sheets))
                return;
            bool anyAdvancedModified = AdvancedXlsbSheetIndexes.Count > 0
                && AdvancedXlsbSheetIndexes.Any(i => i >= 0 && i < Worksheets.Count && Worksheets[i].IsModified);
            bool structureChanged = _openedSheetNames is not null
                && (Worksheets.Count != _openedSheetNames.Count
                    || !Worksheets.Select((w, i) => w.Name).SequenceEqual(_openedSheetNames));
            if (!anyAdvancedModified && !structureChanged)
                return;
            ReportOrBlockAdvancedXlsb(format,
                $"源 XLSB 文件包含当前模型无法安全合并的高级部件，保存到 {format} 时这些部件可能丢失。");
            return;
        }

        // 跨格式：目标 xlsx/xlsm 时，BIFF12 高级部件（透视表/切片器/连接/PQ/数据模型）会经转码器保留；
        // 仅当目标为 xls/csv 等无法承载这些部件的格式，或含尚无转码器覆盖的部件（如时间线）时才上报/阻止。
        if (format == ExcelFormat.Xlsx || format == ExcelFormat.Xlsm)
        {
            bool disabled = Environment.GetEnvironmentVariable("LITEXCEL_DISABLE_PIVOT_WIRING") == "1";
            if (!disabled && !SourceHasUncoveredAdvancedXlsbParts)
                return; // 透视/切片器/连接/数据模型已转码保留
        }

        ReportOrBlockAdvancedXlsb(format,
            $"源 XLSB 文件包含高级部件（透视表/切片器/连接/Power Query/数据模型等），转换为 {format} 时这些部件无法保留；" +
            "VBA、主题、customXml、媒体、ActiveX 等格式无关部件会保留。");
    }

    private void ReportOrBlockAdvancedXlsb(ExcelFormat format, string message)
    {
        if (AllowFeatureLossOnSave)
        {
            ReportDegradation(new DegradationInfo
            {
                Capability = DegradationCapability.PivotTables,
                TargetFormat = format,
                Message = message,
            });
            return;
        }
        throw new LiteExcelException(
            message + "\n默认已阻止本次保存。如确认接受功能丢失，请设 workbook.AllowFeatureLossOnSave = true 后重试。");
    }

    /// <summary>记录一次能力降级：累积进 <see cref="SaveDegradations"/> 并透传外部回调。</summary>
    private void ReportDegradation(DegradationInfo info)
    {
        _saveDegradations.Add(info);
        DegradationCallback?.Invoke(info);
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
    /// 是否可以「原样写回」（surgical verbatim）：仅从源文件中删除若干工作表而保持其余全部逐字节不变。
    /// 条件：源为 xlsx/xlsm、当前表是打开时表的有序子集（只删不增/不改名/不移动）、且无任何保留表被修改。
    /// 命中时 XlsxWriter 可原样写出保留部件 + 仅摘除被删表引用，最大限度保留透视表/图表/切片器等高级功能。
    /// </summary>
    private bool CanSurgicalXlsx(List<SheetData> sheets)
    {
        if (Format != ExcelFormat.Xlsx && Format != ExcelFormat.Xlsm) return false;
        if (_openedSheetNames is null) return false;
        if (PreservedParts?.VerbatimXmlParts is null) return false;
        // 必须有表被删除（count 减少）才走；无删除时走 verbatim 路径（calcChain 清理等）
        if (sheets.Count >= _openedSheetNames.Count) return false;

        // 当前表名必须是打开时表名的「有序子序列」（去掉若干项后剩余顺序完全一致）。
        int cur = 0;
        foreach (var name in _openedSheetNames)
        {
            if (cur < sheets.Count && string.Equals(sheets[cur].SheetName, name, StringComparison.Ordinal))
            {
                cur++;
            }
            // 否则视为打开时被删除的表，跳过；
            // 若当前表在打开表中找不到且仍有剩余当前表，则不是纯删除（可能改名/新增）→ 不匹配。
        }
        if (cur != sheets.Count) return false; // 当前表中存在不与打开表顺序对齐的项

        // 无任何工作表被修改
        foreach (var ws in Worksheets)
            if (ws.IsModified) return false;

        // 修改密码变动时需重建 workbook（含 fileSharing）
        if (Security.ModifyPasswordTouched || Security.HasModifyPassword) return false;

        return true;
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
    /// 是否可以原样写回（surgical verbatim）xlsb：仅从源文件中删除若干工作表而保持其余全部逐字节不变。
    /// 条件：源为 xlsb、当前表是打开时表的有序子集（只删不增/不改名/不移动）、且无任何保留表被修改、
    /// 原始二进制部件已捕获、无密码变动。命中时 XlsbWriter 可原样写出保留部件 + 仅摘除被删表引用，
    /// 最大限度保留透视表/切片器/宏等高级功能。
    /// </summary>
    private bool CanSurgicalXlsb(List<SheetData> sheets)
    {
        if (Format != ExcelFormat.Xlsb) return false;
        if (_openedSheetNames is null) return false;
        if (PreservedParts?.VerbatimBinaries is null) return false;
        if (!PreservedParts.VerbatimBinaries.ContainsKey("xl/workbook.bin")) return false;
        // 必须有表被删除（count 减少）才走
        if (sheets.Count >= _openedSheetNames.Count) return false;

        // 当前表名必须是打开时表名的「有序子序列」
        int cur = 0;
        foreach (var name in _openedSheetNames)
        {
            if (cur < sheets.Count && string.Equals(sheets[cur].SheetName, name, StringComparison.Ordinal))
                cur++;
        }
        if (cur != sheets.Count) return false;

        // 无任何工作表被修改
        foreach (var ws in Worksheets)
            if (ws.IsModified) return false;

        // 修改密码变动时需重建 workbook.bin
        if (Security.ModifyPasswordTouched || Security.HasModifyPassword) return false;
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

    internal void OnWorksheetAdded(Worksheet ws)
    {
        Properties.Modified = DateTime.Now;
    }

    internal void OnWorksheetRemoved(Worksheet ws, int removedIndex)
    {
        Properties.Modified = DateTime.Now;
        // 删除命名区域：localSheetId 指向被删表的自动失效；删除表前索引大于 removedIndex 的减一。
        Names.RemoveAll(n => n.IsLocalSheet && n.LocalSheetId == removedIndex);
        foreach (var n in Names)
        {
            if (n.IsLocalSheet && n.LocalSheetId > removedIndex)
                n.LocalSheetId--;
        }

        // 同步清理 preserved 原始 definedNames XML，避免写出后 Excel 报「已删除的功能：命名区域」。
        if (PreservedParts?.DefinedNamesXml is { Length: > 0 })
            PreservedParts.DefinedNamesXml = CleanDefinedNamesXmlOnDelete(PreservedParts.DefinedNamesXml, removedIndex);
    }

    /// <summary>删除工作表后，对原 definedNames XML 文本做同步清理：
    /// localSheetId == removedIndex → 删除该 definedName
    /// localSheetId > removedIndex → 局部索引减一（保持次序正确） </summary>
    private static string? CleanDefinedNamesXmlOnDelete(string xml, int removedIndex)
    {
        if (string.IsNullOrEmpty(xml)) return null;

        var doc = System.Xml.Linq.XDocument.Parse("<root>" + xml + "</root>",
            System.Xml.Linq.LoadOptions.PreserveWhitespace | System.Xml.Linq.LoadOptions.SetLineInfo);

        var root = doc.Root;
        if (root is null) return xml;

        var toRemove = new System.Collections.Generic.List<System.Xml.Linq.XElement>();
        foreach (var dn in root.Elements())
        {
            if (dn.Name.LocalName != "definedName") continue;
            var attr = dn.Attribute("localSheetId");
            if (attr is null) continue; // 全局命名区域不动
            if (!int.TryParse(attr.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int idx))
                continue;

            if (idx == removedIndex)
                toRemove.Add(dn); // 引用被删除表，移除
            else if (idx > removedIndex)
                attr.Value = (idx - 1).ToString(System.Globalization.CultureInfo.InvariantCulture); // 排名前移
        }
        foreach (var el in toRemove) el.Remove();

        // root 只在保存为片段时使用。为避免整个 XML 被包进 <root> 也输出，改写 root 的 InnerXml。
        var content = string.Concat(root.Nodes().Select(n => n.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)));
        return string.IsNullOrEmpty(content) ? null : content;
    }
}
