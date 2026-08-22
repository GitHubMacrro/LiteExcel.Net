using LiteExcel.Internal;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace LiteExcel;

/// <summary>
/// xlsx 读取器 零反射，AOT 安全 
/// </summary>
public static partial class XlsxReader
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly XmlReaderSettings XmlSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        IgnoreWhitespace = true,
    };

    // 内置日期格式 ID
    private static readonly HashSet<int> BuiltInDateFmtIds = new()
    {
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47,
        50, 51, 52, 53, 54, 55, 56, 57, 58,
    };

    // ── 公开 API：文件路径重载 ──

    /// <summary>打开文件流并在进入 zip 前检测加密（CFB 容器 + EncryptionInfo），避免误报 zip 损坏 </summary>
    private static FileStream OpenFileStreamChecked(string path)
    {
        var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        try
        {
            EncryptionDetector.ThrowIfEncryptedOoxml(fs, path);
        }
        catch
        {
            fs.Dispose();
            throw;
        }
        fs.Position = 0;
        return fs;
    }

    /// <summary>列出所有工作表名 </summary>
    public static List<string> GetSheetNames(string path)
    {
        using var fs = OpenFileStreamChecked(path);
        return GetSheetNames(fs);
    }

    /// <summary>按索引读取单表 </summary>
    public static SheetData Read(string path, int sheetIndex, bool firstRowIsHeader = true)
    {
        using var fs = OpenFileStreamChecked(path);
        return Read(fs, sheetIndex, firstRowIsHeader);
    }

    /// <summary>按名称读取单表 </summary>
    public static SheetData Read(string path, string sheetName, bool firstRowIsHeader = true)
    {
        using var fs = OpenFileStreamChecked(path);
        return Read(fs, sheetName, firstRowIsHeader);
    }

    /// <summary>读取所有工作表 </summary>
    public static List<SheetData> ReadAll(string path)
    {
        using var fs = OpenFileStreamChecked(path);
        return ReadAll(fs);
    }

    /// <summary>流式读取大文件，逐行回调，不驻留内存 </summary>
    public static void StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)
    {
        using var fs = OpenFileStreamChecked(path);
        StreamRows(fs, sheetName, onRow);
    }

    /// <summary>
    /// 带进度回调的读取 先快速扫描该 sheet 获取总数据行数，再流式逐行读取 
    /// <paramref name="onProgress"/> 的 current 从 1 递增到 total（数据行数，不含表头） 
    /// </summary>
    public static void ReadWithProgress(string path, int sheetIndex, Action<int, int> onProgress)
    {
        if (onProgress is null) throw new ArgumentNullException(nameof(onProgress));

        //   快速扫描获取总数据行数（仅遍历 <row> 元素计数，不解析单元格）
        int totalDataRows;
        using (var fsScan = OpenFileStreamChecked(path))
        using (var zipScan = new ZipArchive(fsScan, ZipArchiveMode.Read))
        {
            var sheetsScan = ReadWorkbook(zipScan);
            if (sheetIndex < 0 || sheetIndex >= sheetsScan.Count)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"工作表索引超出范围：{sheetIndex}（共 {sheetsScan.Count} 张表）");

            var infoScan = sheetsScan[sheetIndex];
            var entryScan = zipScan.GetEntry(infoScan.Path)
                ?? throw new LiteExcelException($"缺少工作表文件: {infoScan.Path}");

            int rowCount = 0;
            using (var readerScan = XmlReader.Create(entryScan.Open(), XmlSettings))
            {
                while (readerScan.Read())
                {
                    if (readerScan.NodeType == XmlNodeType.Element && readerScan.LocalName == "row")
                        rowCount++;
                }
            }
            // 第一行为表头，数据行 = 总行数 - 1
            totalDataRows = rowCount > 0 ? rowCount - 1 : 0;
        }

        //   流式逐行读取，每读一行回调 onProgress(current, total)
        using (var fs = OpenFileStreamChecked(path))
        using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            var shared = ReadSharedStrings(zip);
            var styles = ReadStyles(zip);
            var sheets = ReadWorkbook(zip);
            var info = sheets[sheetIndex];
            var entry = zip.GetEntry(info.Path)
                ?? throw new LiteExcelException($"缺少工作表文件: {info.Path}");

            using var reader = XmlReader.Create(entry.Open(), XmlSettings);
            bool firstRow = true;
            int current = 0;
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") continue;
                using var sub = reader.ReadSubtree();
                var row = ParseRow(sub, shared, styles);
                if (row is null) continue;
                if (firstRow) { firstRow = false; continue; }
                current++;
                onProgress(current, totalDataRows);
            }
        }
    }

    // ── 公开 API：Stream 重载 ──

    /// <summary>列出所有工作表名 </summary>
    public static List<string> GetSheetNames(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        return ReadWorkbook(zip).Select(s => s.Name).ToList();
    }

    /// <summary>按索引读取单表 </summary>
    public static SheetData Read(Stream stream, int sheetIndex, bool firstRowIsHeader = true)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var shared = ReadSharedStrings(zip);
        var styles = ReadStyles(zip);
        var sheets = ReadWorkbook(zip);

        if (sheetIndex < 0 || sheetIndex >= sheets.Count)
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"工作表索引超出范围：{sheetIndex}（共 {sheets.Count} 张表）");

        var info = sheets[sheetIndex];
        return ReadWorksheet(zip, info.Path, info.Name, shared, styles, firstRowIsHeader);
    }

    /// <summary>按名称读取单表 </summary>
    public static SheetData Read(Stream stream, string sheetName, bool firstRowIsHeader = true)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var shared = ReadSharedStrings(zip);
        var styles = ReadStyles(zip);
        var sheets = ReadWorkbook(zip);

        var info = sheets.FirstOrDefault(s => s.Name == sheetName)
            ?? throw new LiteExcelException($"找不到工作表：{sheetName}（共有 {sheets.Count} 张表）");

        return ReadWorksheet(zip, info.Path, info.Name, shared, styles, firstRowIsHeader);
    }

    /// <summary>读取所有工作表 </summary>
    public static List<SheetData> ReadAll(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var shared = ReadSharedStrings(zip);
        var styles = ReadStyles(zip);
        var sheets = ReadWorkbook(zip);

        var result = new List<SheetData>(sheets.Count);
        foreach (var info in sheets)
        {
            result.Add(ReadWorksheet(zip, info.Path, info.Name, shared, styles, firstRowIsHeader: true));
        }
        return result;
    }

    /// <summary>流式读取，逐行回调，不驻留内存 </summary>
    public static void StreamRows(Stream stream, string sheetName, Action<IReadOnlyList<Cell>> onRow)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var shared = ReadSharedStrings(zip);
        var styles = ReadStyles(zip);
        var sheets = ReadWorkbook(zip);

        var info = sheets.FirstOrDefault(s => s.Name == sheetName)
            ?? throw new LiteExcelException($"找不到工作表：{sheetName}（共有 {sheets.Count} 张表）");

        var entry = zip.GetEntry(info.Path)
            ?? throw new LiteExcelException($"缺少工作表文件: {info.Path}");

        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        bool firstRow = true;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") continue;
            using var sub = reader.ReadSubtree();
            var row = ParseRow(sub, shared, styles);
            if (row is null) continue;
            if (firstRow) { firstRow = false; continue; }
            onRow(row);
        }
    }

    /// <summary>读取工作簿文档属性（作者/最后保存者/时间/标题等） 无属性时返回空对象 </summary>
    public static WorkbookProperties ReadProperties(string path)
    {
        using var fs = OpenFileStreamChecked(path);
        return ReadProperties(fs);
    }

    /// <summary>从流读取工作簿文档属性 无属性时返回空对象 </summary>
    public static WorkbookProperties ReadProperties(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        return ReadProperties(zip);
    }

    /// <summary>从 ZipArchive 读取工作簿文档属性（供单次解压复用） </summary>
    internal static WorkbookProperties ReadProperties(ZipArchive zip)
    {
        var props = new WorkbookProperties();

        var coreEntry = zip.GetEntry("docProps/core.xml");
        if (coreEntry is not null)
        {
            var doc = XElement.Load(coreEntry.Open());
            // cp 前缀命名空间（core-properties 元数据）
            var cpNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            var dcNs = XNamespace.Get("http://purl.org/dc/elements/1.1/");
            var dctermsNs = XNamespace.Get("http://purl.org/dc/terms/");

            props.Creator = GetElementText(doc, dcNs + "creator");
            props.LastModifiedBy = GetElementText(doc, cpNs + "lastModifiedBy");
            props.Title = GetElementText(doc, dcNs + "title");
            props.Subject = GetElementText(doc, dcNs + "subject");

            var created = GetElementText(doc, dctermsNs + "created");
            if (DateTime.TryParse(created, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var createdVal))
                props.Created = createdVal.ToLocalTime();

            var modified = GetElementText(doc, dctermsNs + "modified");
            if (DateTime.TryParse(modified, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var modifiedVal))
                props.Modified = modifiedVal.ToLocalTime();
        }

        var appEntry = zip.GetEntry("docProps/app.xml");
        if (appEntry is not null)
        {
            var doc = XElement.Load(appEntry.Open());
            var ns = doc.GetDefaultNamespace();
            props.Application = GetElementText(doc, ns + "Application");
        }

        return props;
    }

    private static string? GetElementText(XElement parent, XName name)
    {
        var el = parent.Element(name);
        return string.IsNullOrEmpty(el?.Value) ? null : el.Value.Trim();
    }
    // ── 内部实现 ──

    private sealed class SheetInfo
    {
        public string Name = "";
        public string Path = "";
    }

    private static List<string> ReadSharedStrings(ZipArchive zip)
    {
        var result = new List<string>();
        var entry = zip.GetEntry("xl/sharedStrings.xml");
        if (entry is null) return result;

        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
            {
                using var sub = reader.ReadSubtree();
                result.Add(ReadRichText(sub));
            }
        }
        return result;
    }

    private static string ReadRichText(XmlReader reader)
    {
        var sb = new StringBuilder();
        var depth = reader.Depth;
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) break;

            if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA)
            {
                sb.Append(reader.Value);
            }
            else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && !reader.IsEmptyElement)
            {
                sb.Append(reader.ReadElementContentAsString());
            }
        }
        return sb.ToString();
    }

    private static StylesheetInfo ReadStyles(ZipArchive zip)
    {
        var entry = zip.GetEntry("xl/styles.xml");
        if (entry is null) return new StylesheetInfo();
        var doc = XElement.Load(entry.Open());
        return Stylesheet.Parse(doc);
    }

    private static List<SheetInfo> ReadWorkbook(ZipArchive zip)
    {
        var result = new List<SheetInfo>();

        var wbEntry = zip.GetEntry("xl/workbook.xml")
            ?? throw new LiteExcelException("这不是有效的 xlsx 文件");
        var workbook = XElement.Load(wbEntry.Open());
        var ns = workbook.GetDefaultNamespace();

        // 检测 1904 日期系统；同时捕获工作簿宿主 VBA 代码名（保存时写回 workbookPr@codeName）
        var wbPr = workbook.Element(ns + "workbookPr");
        var date1904Attr = wbPr?.Attribute("date1904")?.Value;
        var date1904 = date1904Attr == "1" || string.Equals(date1904Attr, "true", StringComparison.OrdinalIgnoreCase);
        s_workbookCodeName = wbPr?.Attribute("codeName")?.Value;

        // 捕获 fileSharing（修改密码/写保护）
        s_fileSharingHash = null;
        var fsEl = workbook.Element(ns + "fileSharing");
        if (fsEl is not null)
        {
            var hash = fsEl.Attribute("hashValue")?.Value;
            var salt = fsEl.Attribute("saltValue")?.Value;
            var algo = fsEl.Attribute("algorithmName")?.Value;
            var spin = fsEl.Attribute("spinCount")?.Value;
            var readOnlyRecommended = fsEl.Attribute("readOnlyRecommended")?.Value == "1";
            if (!string.IsNullOrEmpty(hash))
            {
                s_fileSharingHash = new Internal.Encryption.FileSharingInfo(
                    Convert.FromBase64String(hash),
                    salt is null ? null : Convert.FromBase64String(salt),
                    algo, spin is null ? 0 : int.Parse(spin), readOnlyRecommended);
            }
        }

        // P0-6: 捕获 bookViews / definedNames 原始 XML（保存时原样回写，避免静默丢失命名区域与窗口视图）
        s_bookViewsXml = workbook.Element(ns + "bookViews")?.ToString();
        s_definedNamesXml = workbook.Element(ns + "definedNames")?.ToString();

        // 读取 sheet 列表
        var sheetsEl = workbook.Element(ns + "sheets");
        if (sheetsEl is null) return result;

        var sheetElements = sheetsEl.Elements(ns + "sheet").ToList();
        if (sheetElements.Count == 0) return result;

        // 读取 relationships
        var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
        var relMap = new Dictionary<string, string>();
        if (relsEntry is not null)
        {
            var rels = XElement.Load(relsEntry.Open());
            var relNs = rels.Name.Namespace;
            foreach (var rel in rels.Elements(relNs + "Relationship"))
            {
                var id = rel.Attribute("Id")?.Value;
                var target = rel.Attribute("Target")?.Value ?? "";
                if (id is not null) relMap[id] = target;
            }
        }

        foreach (var s in sheetElements)
        {
            var name = s.Attribute("name")?.Value ?? "";
            var relId = s.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value ?? "";

            string sheetPath = "";
            if (relMap.TryGetValue(relId, out var target))
            {
                if (target.StartsWith("/")) sheetPath = target.TrimStart('/');
                else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)) sheetPath = "xl/" + target;
                else sheetPath = target;
            }

            result.Add(new SheetInfo { Name = name, Path = sheetPath });
        }

        // 把 date1904 标记附加到每个 SheetInfo... 用静态字段更简单
        // 实际上，date1904 是工作簿级别的，不是 sheet 级别的
        // 我们在 ReadWorksheet 时传入
        s_globalDate1904 = date1904;

        return result;
    }

    [ThreadStatic]
    private static bool s_globalDate1904;

    [ThreadStatic]
    private static string? s_workbookCodeName;

    [ThreadStatic]
    private static Internal.Encryption.FileSharingInfo? s_fileSharingHash;

    // P0-6: workbook.xml 中 bookViews / definedNames 的原始 XML 快照
    [ThreadStatic]
    private static string? s_bookViewsXml;

    [ThreadStatic]
    private static string? s_definedNamesXml;

    /// <summary>最近一次 ReadWorkbook 捕获的工作簿 codeName（同线程、单次打开内有效，供 OpenCore 取用） </summary>
    internal static string? WorkbookCodeNameSnapshot => s_workbookCodeName;

    /// <summary>最近一次 ReadWorkbook 捕获的 1904 日期系统标志（同线程、单次打开内有效，供 OpenCore 取用） </summary>
    internal static bool Date1904Snapshot => s_globalDate1904;

    /// <summary>最近一次 ReadWorkbook 捕获的 fileSharing（修改密码）信息 </summary>
    internal static Internal.Encryption.FileSharingInfo? FileSharingSnapshot => s_fileSharingHash;

    /// <summary>P0-6: 最近一次 ReadWorkbook 捕获的 bookViews 原始 XML </summary>
    internal static string? BookViewsXmlSnapshot => s_bookViewsXml;

    /// <summary>P0-6: 最近一次 ReadWorkbook 捕获的 definedNames 原始 XML </summary>
    internal static string? DefinedNamesXmlSnapshot => s_definedNamesXml;

    private static SheetData ReadWorksheet(ZipArchive zip, string sheetPath, string sheetName,
        List<string> shared, StylesheetInfo styles, bool firstRowIsHeader)
    {
        var entry = zip.GetEntry(sheetPath)
            ?? throw new LiteExcelException($"缺少工作表文件: {sheetPath}");

        var sheet = new SheetData { SheetName = sheetName };
        var hiddenRowNumbers = new HashSet<int>(); // 1-based XML row numbers that are hidden

        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;

            // 工作表宿主 VBA 代码名（带宏工作簿保存后与 vbaProject 保持绑定）
            if (reader.LocalName == "sheetPr")
            {
                var codeNameAttr = reader.GetAttribute("codeName");
                if (!string.IsNullOrEmpty(codeNameAttr))
                    sheet.CodeName = codeNameAttr;
            }
            else if (reader.LocalName == "col")
            {
                // P0-1: 回填列宽。仅记录带 customWidth 的用户自定义列，跳过默认宽度的 catch-all 条目。
                // 0 哨兵表示"无自定义列宽"，与各写入器（跳过 <=0）的既有约定一致。
                var minAttr = reader.GetAttribute("min");
                var maxAttr = reader.GetAttribute("max");
                var widthAttr = reader.GetAttribute("width");
                var customAttr = reader.GetAttribute("customWidth");
                bool isCustom = customAttr == "1" || string.Equals(customAttr, "true", StringComparison.OrdinalIgnoreCase);
                if (isCustom && minAttr is not null && maxAttr is not null && widthAttr is not null
                    && int.TryParse(minAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var minC)
                    && int.TryParse(maxAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var maxC)
                    && double.TryParse(widthAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var w)
                    && maxC >= minC && minC >= 1)
                {
                    sheet.ColumnWidths ??= new List<double>();
                    while (sheet.ColumnWidths.Count < maxC)
                        sheet.ColumnWidths.Add(0);
                    for (int c = minC; c <= maxC; c++)
                        sheet.ColumnWidths[c - 1] = w;
                }
            }
            else if (reader.LocalName == "row")
            {
                // Read row attributes before ReadSubtree
                var rAttr = reader.GetAttribute("r");
                var hiddenAttr = reader.GetAttribute("hidden");
                var htAttr = reader.GetAttribute("ht");
                int xmlRowNum = rAttr is not null && int.TryParse(rAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rn) ? rn : 0;
                bool isHidden = hiddenAttr == "1" || string.Equals(hiddenAttr, "true", StringComparison.OrdinalIgnoreCase);

                using var sub = reader.ReadSubtree();
                var row = ParseRow(sub, shared, styles);
                if (row is null) continue;

                if (isHidden)
                    hiddenRowNumbers.Add(xmlRowNum);

                if (firstRowIsHeader && sheet.Headers.Count == 0 && sheet.Rows.Count == 0)
                {
                    foreach (var cell in row)
                    {
                        sheet.Headers.Add(cell.Type == CellType.Text ? (cell.Text ?? "") : cell.Text ?? "");
                    }
                }
                else
                {
                    sheet.Rows.Add(row);
                    // 行高：key = 0-based 数据行索引
                    if (htAttr is not null && double.TryParse(htAttr, NumberStyles.Float, CultureInfo.InvariantCulture, out var ht))
                    {
                        sheet.RowHeights ??= new Dictionary<int, double>();
                        sheet.RowHeights[sheet.Rows.Count - 1] = ht;
                    }
                }
            }
            else if (reader.LocalName == "mergeCell")
            {
                var refAttr = reader.GetAttribute("ref");
                if (refAttr is not null && refAttr.Contains(':'))
                {
                    var parts = refAttr.Split(':');
                    if (parts.Length == 2)
                    {
                        var (r1, c1) = CellRef.Parse(parts[0]);
                        var (r2, c2) = CellRef.Parse(parts[1]);
                        int headerOffset = sheet.Headers.Count > 0 ? 1 : 0;
                        sheet.MergedRanges.Add(new CellRange(
                            Math.Min(r1, r2) - headerOffset, Math.Max(r1, r2) - headerOffset,
                            Math.Min(c1, c2), Math.Max(c1, c2)));
                    }
                }
            }
            else if (reader.LocalName == "dataValidations")
            {
                ParseDataValidations(reader, sheet);
            }
            else if (reader.LocalName == "autoFilter")
            {
                sheet.Filter = ParseAutoFilter(reader);
            }
            else if (reader.LocalName == "conditionalFormatting")
            {
                ParseConditionalFormatting(reader, sheet, styles);
            }
            else if (reader.LocalName == "pane")
            {
                // 冻结窗格：ySplit/xSplit 任意值 + state="frozen"
                var state = reader.GetAttribute("state");
                var ySplit = reader.GetAttribute("ySplit");
                var xSplit = reader.GetAttribute("xSplit");
                if (state == "frozen")
                {
                    if (int.TryParse(ySplit, out int yRows) && yRows > 0)
                        sheet.FreezeRows = yRows;
                    if (int.TryParse(xSplit, out int xCols) && xCols > 0)
                        sheet.FreezeColumns = xCols;
                    if (sheet.FreezeRows > 0 || sheet.FreezeColumns > 0)
                        sheet.FreezeHeader = sheet.FreezeRows == 1;
                }
            }
        }

        // Convert hidden XML row numbers to 0-based data row indices
        if (hiddenRowNumbers.Count > 0 && sheet.Filter is not null)
        {
            int headerOffset = sheet.Headers.Count > 0 ? 1 : 0;
            foreach (var xmlRowNum in hiddenRowNumbers)
            {
                int dataRowIdx = xmlRowNum - 1 - headerOffset;
                if (dataRowIdx >= 0)
                    sheet.Filter.HiddenRows.Add(dataRowIdx);
            }
        }

        // 读取单元格批注（通过 sheet rels 查找 comments 部件）
        ReadCommentsForSheet(zip, sheetPath, sheet);

        // 读取单元格超链接（通过 sheet rels 的 hyperlink 关系 + sheet 的 <hyperlinks> 元素）
        ReadHyperlinksForSheet(zip, sheetPath, sheet);

        // 读取浮动图片（drawing + media）——只读取 BackingFile，不改变表格本身
        ReadImagesForSheet(zip, sheetPath, sheet);

        return sheet;
    }

    /// <summary>读取单元格超链接：sheet rels 的 hyperlink 关系（外部）+ sheet 的 &lt;hyperlinks&gt; 元素（含内部 location） </summary>
    private static void ReadHyperlinksForSheet(ZipArchive zip, string sheetPath, SheetData sheet)
    {
        // 外部超链接经 rels 映射；内部超链接直接在 sheet 中带 location，无需 rels
        var hyperlinkRels = new Dictionary<string, string>(StringComparer.Ordinal);
        var dir = System.IO.Path.GetDirectoryName(sheetPath);
        if (!string.IsNullOrEmpty(dir))
        {
            dir = dir.Replace('\\', '/');
            var file = System.IO.Path.GetFileName(sheetPath);
            var relsPath = $"{dir}/_rels/{file}.rels";

            var relsEntry = zip.GetEntry(relsPath);
            if (relsEntry is not null)
            {
                using (var relsReader = XmlReader.Create(relsEntry.Open(), XmlSettings))
                {
                    while (relsReader.Read())
                    {
                        if (relsReader.NodeType == XmlNodeType.Element && relsReader.LocalName == "Relationship")
                        {
                            var type = relsReader.GetAttribute("Type") ?? "";
                            if (type.EndsWith("/hyperlink", StringComparison.OrdinalIgnoreCase))
                            {
                                var id = relsReader.GetAttribute("Id");
                                var target = relsReader.GetAttribute("Target") ?? "";
                                var mode = relsReader.GetAttribute("TargetMode") ?? "";
                                if (id is not null)
                                {
                                    // 非 External（TargetMode 为空）需解析为相对路径
                                    hyperlinkRels[id] = mode == "External"
                                        ? target
                                        : ResolveRelativePath(sheetPath, target);
                                }
                            }
                        }
                    }
                }
            }
        }

        var entry = zip.GetEntry(sheetPath);
        if (entry is null) return;

        using var reader = XmlReader.Create(entry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "hyperlink") continue;

            var refAttr = reader.GetAttribute("ref");
            var tooltip = reader.GetAttribute("tooltip");
            if (refAttr is null) continue;

            // 内部超链接：location 属性（无需 rels）
            var location = reader.GetAttribute("location");
            if (!string.IsNullOrEmpty(location))
            {
                SetCellHyperlink(sheet, refAttr, new Hyperlink
                {
                    Target = location.StartsWith("#", StringComparison.Ordinal) ? location : "#" + location,
                    Tooltip = tooltip,
                    IsInternal = true,
                });
                continue;
            }

            // 外部超链接：r:id 指向 rels
            var rid = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rid is null || !hyperlinkRels.TryGetValue(rid, out var target)) continue;

            SetCellHyperlink(sheet, refAttr, new Hyperlink
            {
                Target = target,
                Tooltip = tooltip,
                IsInternal = false,
            });
        }
    }

    /// <summary>把超链接挂到 sheet 对应单元格（ref 可能是 "A1" 或区域，取左上角） </summary>
    private static void SetCellHyperlink(SheetData sheet, string refCell, Hyperlink link)
    {
        // ref 可能是 "A1" 或 "A1:B2"（区域），取左上角
        var refCellOnly = refCell.Split(':')[0];
        var (r0, c0) = CellRef.Parse(refCellOnly);
        int rowIdx = r0, colIdx = c0;

        // SheetData 行索引：Header 占 1 行（Read 时 firstRowIsHeader=true），Rows 不含表头
        int dataRow = rowIdx;
        if (sheet.Headers.Count > 0) dataRow = rowIdx - 1;
        if (dataRow < 0 || dataRow >= sheet.Rows.Count) return;
        var row = sheet.Rows[dataRow];
        if (colIdx < 0 || colIdx >= row.Count) return;

        row[colIdx].Hyperlink = link;
    }

    // ── 批注读取 ──

    private static void ReadCommentsForSheet(ZipArchive zip, string sheetPath, SheetData sheet)
    {
        // 从 sheetPath 推导 rels 路径
        // "xl/worksheets/sheet1.xml" -> "xl/worksheets/_rels/sheet1.xml.rels"
        var dir = System.IO.Path.GetDirectoryName(sheetPath);
        if (string.IsNullOrEmpty(dir)) return;
        dir = dir.Replace('\\', '/');
        var file = System.IO.Path.GetFileName(sheetPath);
        var relsPath = $"{dir}/_rels/{file}.rels";

        var relsEntry = zip.GetEntry(relsPath);
        if (relsEntry is null) return;

        string? commentsTarget = null;
        using (var relsReader = XmlReader.Create(relsEntry.Open(), XmlSettings))
        {
            while (relsReader.Read())
            {
                if (relsReader.NodeType == XmlNodeType.Element && relsReader.LocalName == "Relationship")
                {
                    var type = relsReader.GetAttribute("Type") ?? "";
                    if (type.EndsWith("/comments", StringComparison.OrdinalIgnoreCase))
                    {
                        commentsTarget = relsReader.GetAttribute("Target");
                        break;
                    }
                }
            }
        }
        if (commentsTarget is null) return;

        // 解析相对路径
        var commentsPath = ResolveRelativePath(sheetPath, commentsTarget);
        var commentsEntry = zip.GetEntry(commentsPath);
        if (commentsEntry is null) return;

        using var reader = XmlReader.Create(commentsEntry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "comment") continue;

            var refAttr = reader.GetAttribute("ref");
            if (refAttr is null) continue;

            string text = "";
            if (!reader.IsEmptyElement)
            {
                using var sub = reader.ReadSubtree();
                while (sub.Read())
                {
                    if (sub.NodeType == XmlNodeType.Element && sub.LocalName == "t" && !sub.IsEmptyElement)
                    {
                        text += ReadElementText(sub);
                    }
                }
            }

            sheet.Comments ??= new Dictionary<string, string>();
            sheet.Comments[refAttr] = text;
        }
    }

    /// <summary>读取工作表浮动图片：sheet rels 找 drawing → drawing XML 提取锚点+rId → drawing rels 找 media → 读字节 </summary>
    private static void ReadImagesForSheet(ZipArchive zip, string sheetPath, SheetData sheet)
    {
        var dir = System.IO.Path.GetDirectoryName(sheetPath);
        if (string.IsNullOrEmpty(dir)) return;
        dir = dir.Replace('\\', '/');
        var file = System.IO.Path.GetFileName(sheetPath);
        var relsPath = $"{dir}/_rels/{file}.rels";
        var relsEntry = zip.GetEntry(relsPath);
        if (relsEntry is null) return;

        string? drawingTarget = null;
        using (var rr = XmlReader.Create(relsEntry.Open(), XmlSettings))
        {
            while (rr.Read())
            {
                if (rr.NodeType == XmlNodeType.Element && rr.LocalName == "Relationship")
                {
                    var type = rr.GetAttribute("Type") ?? "";
                    if (type.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase))
                    {
                        drawingTarget = rr.GetAttribute("Target");
                        break;
                    }
                }
            }
        }
        if (drawingTarget is null) return;

        var drawingPath = ResolveRelativePath(sheetPath, drawingTarget);
        var drawingEntry = zip.GetEntry(drawingPath);
        if (drawingEntry is null) return;

        // drawing rels（rId → media 路径）
        var drawingRelMap = new Dictionary<string, string>(StringComparer.Ordinal);
        var drawingDir = System.IO.Path.GetDirectoryName(drawingPath) ?? "";
        drawingDir = drawingDir.Replace('\\', '/');
        var drawingName = System.IO.Path.GetFileName(drawingPath);
        var drawingRelsPath = $"{drawingDir}/_rels/{drawingName}.rels";
        var drawingRelsEntry = zip.GetEntry(drawingRelsPath);
        if (drawingRelsEntry is not null)
        {
            using var dr = XmlReader.Create(drawingRelsEntry.Open(), XmlSettings);
            while (dr.Read())
            {
                if (dr.NodeType == XmlNodeType.Element && dr.LocalName == "Relationship")
                {
                    var id = dr.GetAttribute("Id");
                    var type = dr.GetAttribute("Type") ?? "";
                    var target = dr.GetAttribute("Target") ?? "";
                    if (id is not null && type.EndsWith("/image", StringComparison.OrdinalIgnoreCase))
                        drawingRelMap[id] = ResolveRelativePath(drawingPath, target);
                }
            }
        }

        using var reader = XmlReader.Create(drawingEntry.Open(), XmlSettings);
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            bool isOne = reader.LocalName == "oneCellAnchor";
            bool isTwo = reader.LocalName == "twoCellAnchor";
            if (!isOne && !isTwo) continue;

            var editAs = reader.GetAttribute("editAs");
        var img = ReadOneAnchor(reader, drawingRelMap, zip, isOne, isTwo, editAs);
            if (img is null) continue;

            img.Placement = ImagePlacement.Floating;
            sheet.Images ??= new List<WorksheetImage>();
            sheet.Images.Add(img);
        }
    }

    /// <summary>从 oneCellAnchor/twoCellAnchor 子树读取一个 WorksheetImage </summary>
    private static WorksheetImage? ReadOneAnchor(XmlReader anchorReader, Dictionary<string, string> relMap, ZipArchive zip,
        bool isOne, bool isTwo, string? editAs)
    {
        int? col = null, row = null, colOff = null, rowOff = null;
        double? cx = null, cy = null;
        string? name = null, descr = null, embed = null;
        bool inTo = false;

        using var sub = anchorReader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element) continue;
            switch (sub.LocalName)
            {
                case "from":
                    inTo = false;
                    break;
                case "to":
                    inTo = true;
                    continue;
                case "col" when !sub.IsEmptyElement && !inTo:
                    if (int.TryParse(ReadElementText(sub), out var c)) col = c;
                    break;
                case "colOff" when !sub.IsEmptyElement && !inTo:
                    if (int.TryParse(ReadElementText(sub), out var co)) colOff = co;
                    break;
                case "row" when !sub.IsEmptyElement && !inTo:
                    if (int.TryParse(ReadElementText(sub), out var rw)) row = rw;
                    break;
                case "rowOff" when !sub.IsEmptyElement && !inTo:
                    if (int.TryParse(ReadElementText(sub), out var ro)) rowOff = ro;
                    break;
                case "ext":
                    if (!cx.HasValue)
                    {
                        if (double.TryParse(sub.GetAttribute("cx"), NumberStyles.Float, CultureInfo.InvariantCulture, out var w)) cx = w;
                        if (double.TryParse(sub.GetAttribute("cy"), NumberStyles.Float, CultureInfo.InvariantCulture, out var h)) cy = h;
                    }
                    break;
                case "cNvPr":
                    name ??= sub.GetAttribute("name");
                    descr ??= sub.GetAttribute("descr");
                    break;
                case "blip":
                    embed ??= sub.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    break;
            }
        }

        if (embed is null || col is null || row is null) return null;
        if (!relMap.TryGetValue(embed, out var mediaPath)) return null;
        var mediaEntry = zip.GetEntry(mediaPath);
        if (mediaEntry is null) return null;

        byte[] data;
        using (var ms = new MemoryStream())
        {
            using var s = mediaEntry.Open();
            s.CopyTo(ms);
            data = ms.ToArray();
        }

        return new WorksheetImage
        {
            Data = data,
            Extension = System.IO.Path.GetExtension(mediaPath).TrimStart('.'),
            Row = row.Value + 1,
            Column = col.Value + 1,
            Name = name,
            AltText = descr,
            Anchor = new ImageAnchor
            {
                TopLeftCell = CellRef.ToString(row.Value, col.Value),
                TopLeftOffsetX = colOff ?? 0,
                TopLeftOffsetY = rowOff ?? 0,
                WidthPixels = cx.HasValue ? cx.Value / WorksheetImage.EmuPerPixel : 0,
                HeightPixels = cy.HasValue ? cy.Value / WorksheetImage.EmuPerPixel : 0,
                MoveMode = isTwo
                    ? ImageMoveMode.MoveAndSizeWithCells
                    : (string.Equals(editAs, "absolute", StringComparison.OrdinalIgnoreCase)
                        ? ImageMoveMode.FixedPosition
                        : ImageMoveMode.MoveButDontSizeWithCells),
            },
        };
    }

    private static string ResolveRelativePath(string basePath, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/")) return target.TrimStart('/');

        var baseDir = System.IO.Path.GetDirectoryName(basePath)?.Replace('\\', '/') ?? "";
        var parts = (baseDir + "/" + target).Split('/');
        var stack = new List<string>();
        foreach (var part in parts)
        {
            if (string.IsNullOrEmpty(part) || part == ".") continue;
            if (part == "..")
            {
                if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                continue;
            }
            stack.Add(part);
        }
        return string.Join("/", stack);
    }

    private static AutoFilter ParseAutoFilter(XmlReader reader)
    {
        var filter = new AutoFilter();
        filter.Range = reader.GetAttribute("ref") ?? "";

        if (reader.IsEmptyElement) return filter;

        using var sub = reader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element || sub.LocalName != "filterColumn") continue;

            var colIdAttr = sub.GetAttribute("colId");
            if (colIdAttr is null || !int.TryParse(colIdAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var colId))
                continue;

            var col = new FilterColumn { ColumnIndex = colId };
            ParseFilterColumnContent(sub, col);
            filter.Columns.Add(col);
        }

        return filter;
    }

    private static void ParseFilterColumnContent(XmlReader reader, FilterColumn col)
    {
        if (reader.IsEmptyElement) return;

        using var sub = reader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element) continue;

            if (sub.LocalName == "filters")
            {
                col.Type = FilterType.Equals;
                if (sub.IsEmptyElement) continue;
                using var filtersSub = sub.ReadSubtree();
                while (filtersSub.Read())
                {
                    if (filtersSub.NodeType == XmlNodeType.Element && filtersSub.LocalName == "blank")
                    {
                        col.Type = FilterType.Blank;
                    }
                    else if (filtersSub.NodeType == XmlNodeType.Element && filtersSub.LocalName == "filter")
                    {
                        var val = filtersSub.GetAttribute("val");
                        if (val is not null) col.Values.Add(val);
                    }
                }
            }
            else if (sub.LocalName == "customFilters")
            {
                if (sub.IsEmptyElement) continue;
                using var customSub = sub.ReadSubtree();
                while (customSub.Read())
                {
                    if (customSub.NodeType != XmlNodeType.Element || customSub.LocalName != "customFilter") continue;

                    var op = customSub.GetAttribute("operator") ?? "equal";
                    var val = customSub.GetAttribute("val") ?? "";

                    if (val.Contains('*'))
                    {
                        // Wildcard filter: contains/begins/ends
                        if (val.StartsWith("*") && val.EndsWith("*"))
                        {
                            col.Type = FilterType.Contains;
                            col.Values.Add(val.Trim('*'));
                        }
                        else if (val.StartsWith("*"))
                        {
                            col.Type = FilterType.EndsWith;
                            col.Values.Add(val.TrimStart('*'));
                        }
                        else if (val.EndsWith("*"))
                        {
                            col.Type = FilterType.BeginsWith;
                            col.Values.Add(val.TrimEnd('*'));
                        }
                    }
                    else
                    {
                        col.Type = FilterType.Compare;
                        col.Operator = op switch
                        {
                            "greaterThan" => FilterOperator.GreaterThan,
                            "greaterThanOrEqual" => FilterOperator.GreaterThanOrEqual,
                            "lessThan" => FilterOperator.LessThan,
                            "lessThanOrEqual" => FilterOperator.LessThanOrEqual,
                            _ => FilterOperator.GreaterThan,
                        };
                        col.Values.Add(val);
                    }
                }
            }
        }
    }

    private static IReadOnlyList<Cell>? ParseRow(XmlReader rowReader, List<string> shared, StylesheetInfo styles)
    {
        var cells = new List<(int col, Cell cell)>();
        int maxCol = -1;

        while (rowReader.Read())
        {
            if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") continue;

            var refAttr = rowReader.GetAttribute("r");
            int col = refAttr is not null ? CellRef.Parse(refAttr).col : maxCol + 1;
            var t = rowReader.GetAttribute("t") ?? "";
            var sAttr = rowReader.GetAttribute("s");

            string raw = "";
            bool hasValue = false;
            string formula = "";
            bool hasFormula = false;

            if (!rowReader.IsEmptyElement)
            {
                var cDepth = rowReader.Depth;
                while (rowReader.Read())
                {
                    if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == cDepth) break;
                    if (rowReader.NodeType != XmlNodeType.Element) continue;

                    if (rowReader.LocalName == "f" && !rowReader.IsEmptyElement)
                    {
                        formula = ReadElementText(rowReader);
                        hasFormula = true;
                    }
                    else if (rowReader.LocalName == "v" && !rowReader.IsEmptyElement)
                    {
                        raw = ReadElementText(rowReader);
                        hasValue = true;
                    }
                    else if (rowReader.LocalName == "is")
                    {
                        using var sub = rowReader.ReadSubtree();
                        raw = ReadRichText(sub);
                        hasValue = true;
                    }
                }
            }

            if (!hasValue)
            {
                var cell = Cell.Empty;
                if (hasFormula)
                {
                    cell.Type = CellType.Text;
                    cell.Formula = formula;
                    cell.IsFormula = true;
                }
                ApplyStyle(cell, sAttr, styles);
                cells.Add((col, cell));
            }
            else
            {
                var cell = ConvertCellValue(raw, t, sAttr, shared, styles);
                if (hasFormula)
                {
                    // P0-8: 公式串放入 Formula，不覆盖缓存值（Text/Number/Date/Boolean）
                    cell.IsFormula = true;
                    cell.Formula = formula;
                }
                cells.Add((col, cell));
            }

            if (col > maxCol) maxCol = col;
        }

        if (maxCol < 0) return null;

        // 填充为连续数组（空位补 Empty）
        var arr = new Cell[maxCol + 1];
        for (int i = 0; i <= maxCol; i++) arr[i] = Cell.Empty;
        foreach (var (col, cell) in cells) arr[col] = cell;

        return arr;
    }

    private static void ApplyStyle(Cell cell, string? sAttr, StylesheetInfo styles)
    {
        if (sAttr is null || !int.TryParse(sAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var xfId))
            return;
        if (xfId < 0 || xfId >= styles.CellXfs.Count) return;

        var xf = styles.CellXfs[xfId];
        cell.Style = xf.ToCellStyle(styles);
        cell.NumberFormat = GetNumberFormat(xfId, styles);
    }

    private static Cell ConvertCellValue(string raw, string t, string? sAttr, List<string> shared, StylesheetInfo styles)
    {
        Cell cell;

        // 共享字符串
        if (t == "s")
        {
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx)
                && idx >= 0 && idx < shared.Count)
            {
                cell = Cell.FromText(shared[idx]);
            }
            else
            {
                cell = Cell.Empty;
            }
        }
        else if (t == "str" || t == "inlineStr")
        {
            cell = Cell.FromText(raw);
        }
        else if (t == "b")
        {
            cell = Cell.FromBoolean(raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase));
        }
        else if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
        {
            string? fmtCode = GetNumberFormat(sAttr, styles);
            if (fmtCode is not null && IsDateFormat(0, fmtCode))
            {
                cell = Cell.FromDate(SerialToDate(d, s_globalDate1904), fmtCode);
            }
            else
            {
                cell = Cell.FromNumber(d, fmtCode);
            }
        }
        else
        {
            cell = Cell.FromText(raw);
        }

        ApplyStyle(cell, sAttr, styles);
        return cell;
    }

    private static string? GetNumberFormat(string? sAttr, StylesheetInfo styles)
    {
        if (sAttr is null || !int.TryParse(sAttr, NumberStyles.Integer, CultureInfo.InvariantCulture, out var xfId))
            return null;
        return GetNumberFormat(xfId, styles);
    }

    private static string? GetNumberFormat(int xfId, StylesheetInfo styles)
    {
        if (xfId < 0 || xfId >= styles.CellXfs.Count) return null;

        var xf = styles.CellXfs[xfId];
        var numFmtId = xf.NumFmtId;

        if (numFmtId < 164)
        {
            return GetBuiltInFormatCode(numFmtId);
        }

        return styles.NumFmts.TryGetValue(numFmtId, out var code) ? code : null;
    }

    private static bool IsDateFormat(int numFmtId, string formatCode)
    {
        // 如果有格式串，检查是否包含日期/时间标记
        if (!string.IsNullOrEmpty(formatCode))
        {
            return FormatStringIsDate(formatCode);
        }
        // 否则检查内置 ID
        return BuiltInDateFmtIds.Contains(numFmtId);
    }

    private static bool FormatStringIsDate(string fmt)
    {
        // 去掉方括号内容（如 [Red]、[$-404]）
        var sb = new StringBuilder(fmt.Length);
        bool inBracket = false;
        foreach (var ch in fmt)
        {
            if (ch == '[') { inBracket = true; continue; }
            if (ch == ']') { inBracket = false; continue; }
            if (!inBracket) sb.Append(ch);
        }
        var clean = sb.ToString();

        // 检查日期/时间标记（不区分大小写）
        // y=年, d=日, h=时, s=秒, m需要配合y/d判断
        // 简化：只要包含 y d h s 之一，或包含 m 且上下文为日期，就视为日期
        var lower = clean.ToLowerInvariant();
        if (lower.Contains('y') || lower.Contains('d') || lower.Contains('h') || lower.Contains('s'))
            return true;
        // 单独的 m 可能是分钟，但在日期上下文是月份
        // 保守判断：包含 m 且没有其他非日期内容时也视为日期
        // 实际上这里很难完美判断，保守起见不把单独的 m 当日期
        return false;
    }

    private static string? GetBuiltInFormatCode(int numFmtId)
    {
        return numFmtId switch
        {
            0 => null,  // General
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "yyyy-MM-dd",
            15 => "dd-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "yyyy-MM-dd h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => null,
        };
    }

    private static DateTime SerialToDate(double serial, bool date1904)
    {
        if (date1904)
        {
            // 1904 系统：epoch = 1904-01-01
            return new DateTime(1904, 1, 1).AddDays(serial);
        }
        // 1900 系统：用 .NET 内置转换（自动处理 1900 闰年 bug）
        return DateTime.FromOADate(serial);
    }

    private static void ParseDataValidations(XmlReader reader, SheetData sheet)
    {
        if (reader.IsEmptyElement) return;

        using var sub = reader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element || sub.LocalName != "dataValidation") continue;

            var dv = new DataValidation();
            var typeAttr = sub.GetAttribute("type") ?? "";
            dv.Type = typeAttr switch
            {
                "whole" => DataValidationType.WholeNumber,
                "decimal" => DataValidationType.Decimal,
                "date" => DataValidationType.Date,
                _ => DataValidationType.List,
            };

            dv.Sqref = sub.GetAttribute("sqref") ?? "";

            var allowBlankAttr = sub.GetAttribute("allowBlank");
            dv.AllowBlank = allowBlankAttr == "1" || string.Equals(allowBlankAttr, "true", StringComparison.OrdinalIgnoreCase);

            dv.PromptTitle = sub.GetAttribute("promptTitle");
            dv.Prompt = sub.GetAttribute("prompt");

            // Read formula1 and formula2
            if (!sub.IsEmptyElement)
            {
                using var dvSub = sub.ReadSubtree();
                while (dvSub.Read())
                {
                    if (dvSub.NodeType != XmlNodeType.Element) continue;

                    if (dvSub.LocalName == "formula1")
                    {
                        string raw = ReadElementText(dvSub);
                        // List 类型去掉外层双引号
                        if (dv.Type == DataValidationType.List && raw.Length >= 2 && raw[0] == '"' && raw[raw.Length - 1] == '"')
                        {
                            dv.Formula1 = raw.Substring(1, raw.Length - 2);
                        }
                        else
                        {
                            dv.Formula1 = raw;
                        }
                    }
                    else if (dvSub.LocalName == "formula2")
                    {
                        dv.Formula2 = ReadElementText(dvSub);
                    }
                }
            }

            sheet.Validations ??= new List<DataValidation>();
            sheet.Validations.Add(dv);
        }
    }

    /// <summary>解析条件格式（cellIs / expression / colorScale / dataBar） 及其 dxfId 关联的样式 </summary>
    private static void ParseConditionalFormatting(XmlReader reader, SheetData sheet, StylesheetInfo styles)
    {
        if (reader.IsEmptyElement) return;
        var sqref = reader.GetAttribute("sqref") ?? "";

        using var sub = reader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element || sub.LocalName != "cfRule") continue;

            var typeAttr = sub.GetAttribute("type") ?? "";
            var dxfIdAttr = sub.GetAttribute("dxfId");
            int? dxfId = dxfIdAttr is not null && int.TryParse(dxfIdAttr, out var did) ? did : null;
            var prioAttr = sub.GetAttribute("priority");
            int prio = prioAttr is not null && int.TryParse(prioAttr, out var p) ? p : 0;

            var cf = new ConditionalFormat { Sqref = sqref, Priority = prio };

            switch (typeAttr)
            {
                case "cellIs":
                {
                    cf.Type = ConditionalFormatType.CellIs;
                    cf.Operator = sub.GetAttribute("operator") switch
                    {
                        "lessThan" => ConditionalOperator.LessThan,
                        "lessThanOrEqual" => ConditionalOperator.LessThanOrEqual,
                        "equal" => ConditionalOperator.Equal,
                        "notEqual" => ConditionalOperator.NotEqual,
                        "greaterThanOrEqual" => ConditionalOperator.GreaterThanOrEqual,
                        "between" => ConditionalOperator.Between,
                        "notBetween" => ConditionalOperator.NotBetween,
                        _ => ConditionalOperator.GreaterThan,
                    };
                    ReadCfFormulas(sub, out var f1, out var f2);
                    cf.Formula = f1; cf.Formula2 = f2;
                    if (dxfId.HasValue && dxfId.Value < styles.Dxfs.Count)
                        cf.Style = styles.Dxfs[dxfId.Value].Clone();
                    break;
                }
                case "expression":
                {
                    cf.Type = ConditionalFormatType.Expression;
                    ReadCfFormulas(sub, out var f1, out _);
                    cf.Formula = f1;
                    if (dxfId.HasValue && dxfId.Value < styles.Dxfs.Count)
                        cf.Style = styles.Dxfs[dxfId.Value].Clone();
                    break;
                }
                case "colorScale":
                {
                    cf.Type = ConditionalFormatType.ColorScale;
                    ReadColorScale(sub, styles, cf);
                    break;
                }
                case "dataBar":
                {
                    cf.Type = ConditionalFormatType.DataBar;
                    ReadDataBar(sub, cf);
                    if (dxfId.HasValue && dxfId.Value < styles.Dxfs.Count)
                        cf.Style = styles.Dxfs[dxfId.Value].Clone();
                    break;
                }
            }

            sheet.ConditionalFormats.Add(cf);
        }
    }

    private static void ReadCfFormulas(XmlReader cfRuleReader, out string? f1, out string? f2)
    {
        f1 = null; f2 = null;
        if (cfRuleReader.IsEmptyElement) return;
        using var sub = cfRuleReader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element || sub.LocalName != "formula") continue;
            var text = ReadElementText(sub);
            if (f1 is null) f1 = text; else f2 ??= text;
        }
    }

    private static void ReadColorScale(XmlReader cfRuleReader, StylesheetInfo styles, ConditionalFormat cf)
    {
        if (cfRuleReader.IsEmptyElement) return;
        using var sub = cfRuleReader.ReadSubtree();
        var colors = new List<string>();
        bool inScale = false;
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element) continue;
            if (sub.LocalName == "colorScale")
            {
                inScale = true;
                continue;
            }
            if (inScale && sub.LocalName == "color")
            {
                var rgb = sub.GetAttribute("rgb");
                if (!string.IsNullOrEmpty(rgb)) colors.Add(NormalizeToCssColor(rgb));
            }
        }
        var map = new ColorScaleInfo();
        if (colors.Count >= 1) map.LowColor = colors[0];
        if (colors.Count >= 3) { map.MidColor = colors[1]; map.HighColor = colors[2]; }
        else if (colors.Count >= 2) map.HighColor = colors[1];
        cf.ColorScale = map;
    }

    private static void ReadDataBar(XmlReader cfRuleReader, ConditionalFormat cf)
    {
        if (cfRuleReader.IsEmptyElement) return;
        using var sub = cfRuleReader.ReadSubtree();
        var info = new DataBarInfo();
        while (sub.Read())
        {
            if (sub.NodeType != XmlNodeType.Element) continue;
            if (sub.LocalName == "dataBar")
            {
                var minLen = sub.GetAttribute("minLength");
                var maxLen = sub.GetAttribute("maxLength");
                var showVal = sub.GetAttribute("showValue");
                if (minLen is not null && int.TryParse(minLen, out var m1)) info.MinLengthPercent = m1;
                if (maxLen is not null && int.TryParse(maxLen, out var m2)) info.MaxLengthPercent = m2;
                if (showVal is not null) info.ShowValue = showVal == "1" || string.Equals(showVal, "true", StringComparison.OrdinalIgnoreCase);
            }
            else if (sub.LocalName == "color")
            {
                var color = sub.GetAttribute("rgb");
                if (!string.IsNullOrEmpty(color)) info.Color = NormalizeToCssColor(color);
            }
        }
        cf.DataBar = info;
    }

    /// <summary>#FFRRGGBB / RGB → #RRGGBB（CSS 风格），失败回原值 </summary>
    private static string NormalizeToCssColor(string rgb)
    {
        if (rgb.Length == 8) return "#" + rgb.Substring(2);
        if (rgb.Length == 6) return "#" + rgb;
        return rgb;
    }

    private static string ReadElementText(XmlReader reader)
    {
        var sb = new StringBuilder();
        using var sub = reader.ReadSubtree();
        while (sub.Read())
        {
            if (sub.NodeType == XmlNodeType.Text || sub.NodeType == XmlNodeType.CDATA)
            {
                sb.Append(sub.Value);
            }
        }
        return sb.ToString();
    }
}

