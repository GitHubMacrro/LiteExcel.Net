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

    /// <summary>列出所有工作表名 </summary>
    public static List<string> GetSheetNames(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return GetSheetNames(fs);
    }

    /// <summary>按索引读取单表 </summary>
    public static SheetData Read(string path, int sheetIndex, bool firstRowIsHeader = true)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Read(fs, sheetIndex, firstRowIsHeader);
    }

    /// <summary>按名称读取单表 </summary>
    public static SheetData Read(string path, string sheetName, bool firstRowIsHeader = true)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Read(fs, sheetName, firstRowIsHeader);
    }

    /// <summary>读取所有工作表 </summary>
    public static List<SheetData> ReadAll(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ReadAll(fs);
    }

    /// <summary>流式读取大文件，逐行回调，不驻留内存 </summary>
    public static void StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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
        using (var fsScan = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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

    /// <summary>最近一次 ReadWorkbook 捕获的工作簿 codeName（同线程、单次打开内有效，供 OpenCore 取用） </summary>
    internal static string? WorkbookCodeNameSnapshot => s_workbookCodeName;

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
            else if (reader.LocalName == "pane")
            {
                // 冻结窗格：ySplit 或 xSplit 存在且 state="frozen" 视为冻结首行/首列
                var state = reader.GetAttribute("state");
                var ySplit = reader.GetAttribute("ySplit");
                var xSplit = reader.GetAttribute("xSplit");
                if (state == "frozen" && (ySplit == "1" || xSplit == "1"))
                    sheet.FreezeHeader = true;
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

        return sheet;
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
                    cell.Text = formula;
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
                    cell.IsFormula = true;
                    cell.Text = formula;
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

