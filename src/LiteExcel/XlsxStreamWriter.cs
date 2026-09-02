using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LiteExcel;

/// <summary>
/// 流式写入器：逐行写入大文件，不驻留内存。
/// 采用内联字符串（inlineStr），避免共享字符串表预扫描。
/// 支持样式/公式/超链接随行写入（styles.xml 与 sheet rels 在 Close 时统一写出）。
/// 合并/筛选/图片等高级能力不支持。
/// 使用后必须调用 <see cref="Dispose"/> 或 <see cref="Close"/> 完成文件。
/// 注意：超链接数量极大时内存不再恒定（内部缓冲全部超链接引用）。
/// 单表行数上限为 1,048,576；超出时按 <see cref="RowLimitExceededMode"/> 处理（默认抛异常）。
/// 在 <see cref="RowLimitExceededMode.SpillToNewSheet"/> 模式下，可提供 <c>spillHeader</c> 让每张表（含 Sheet1）首行写入表头，调用方只需写数据行。
/// </summary>
public sealed class XlsxStreamWriter : IDisposable
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal const int ExcelMaxRows = 1_048_576;

    private readonly ZipArchive _zip;
    private readonly Stream _underlying;
    private readonly bool _ownsStream;
    private readonly bool _macroEnabled;
    private readonly RowLimitExceededMode _mode;
    private readonly object?[]? _spillHeader;
    private readonly Internal.Stylesheet _stylesheet = new();
    private readonly List<SheetContext> _sheets = new();
    private SheetContext _current = null!;
    private bool _closed;
    private bool _truncated;

    /// <summary>测试钩子：单表行数上限（测试时设小值以避免写百万行）。</summary>
    internal int MaxRowsPerSheet = ExcelMaxRows;

    /// <summary>
    /// 在 <see cref="RowLimitExceededMode.Truncate"/> 模式下，达到上限并丢弃后续行后为 <c>true</c>。
    /// 其他模式恒为 <c>false</c>。
    /// </summary>
    public bool Truncated => _truncated;

    private sealed class SheetContext
    {
        public string Name = "";
        public string Entry = "";
        public Stream? Stream;
        public XmlWriter? Writer;
        public int RowCount;
        public List<(string Ref, string Target, string? Tooltip, bool IsInternal)> Hyperlinks = new();
    }

    private XlsxStreamWriter(Stream stream, bool ownsStream, RowLimitExceededMode mode, bool macroEnabled = false, object?[]? spillHeader = null)
    {
        _underlying = stream;
        _ownsStream = ownsStream;
        _mode = mode;
        _macroEnabled = macroEnabled;
        _spillHeader = spillHeader;
        _zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WriteStableHead();
        StartNewSheet();
    }

    /// <summary>创建流式写入器（写入指定路径，覆盖已存在文件）。<paramref name="spillHeader"/> 仅在 <see cref="RowLimitExceededMode.SpillToNewSheet"/> 下生效，作为每张表的首行表头 </summary>
    public static XlsxStreamWriter Create(string path, RowLimitExceededMode onRowLimitExceeded = RowLimitExceededMode.Throw, object?[]? spillHeader = null)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
        return new XlsxStreamWriter(fs, ownsStream: true,
            macroEnabled: string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase),
            mode: onRowLimitExceeded, spillHeader: spillHeader);
    }

    /// <summary>创建流式写入器（写入流，LeaveOpen 由调用方管理）。<paramref name="spillHeader"/> 仅在 <see cref="RowLimitExceededMode.SpillToNewSheet"/> 下生效，作为每张表的首行表头 </summary>
    public static XlsxStreamWriter Create(Stream stream, RowLimitExceededMode onRowLimitExceeded = RowLimitExceededMode.Throw, object?[]? spillHeader = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("流不可写", nameof(stream));
        return new XlsxStreamWriter(stream, ownsStream: false, mode: onRowLimitExceeded, spillHeader: spillHeader);
    }

    /// <summary>写入一行单元格值 </summary>
    public void WriteRow(IEnumerable<object?> values)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (_closed) throw new InvalidOperationException("写入器已关闭");
        if (_truncated) return;

        WriteRowStart();
        if (_truncated) return;
        int col = 1;
        foreach (var value in values)
        {
            WriteCell(col, CellFactory.FromObject(value));
            col++;
        }
        _current.Writer!.WriteEndElement();
    }

    /// <summary>写入一行 Cell（低层模型） </summary>
    public void WriteRow(IEnumerable<Cell> cells)
    {
        if (cells is null) throw new ArgumentNullException(nameof(cells));
        if (_closed) throw new InvalidOperationException("写入器已关闭");
        if (_truncated) return;

        WriteRowStart();
        if (_truncated) return;
        int col = 1;
        foreach (var cell in cells)
        {
            WriteCell(col, cell);
            col++;
        }
        _current.Writer!.WriteEndElement();
    }

    private void WriteRowStart()
    {
        if (_current.RowCount >= MaxRowsPerSheet)
        {
            switch (_mode)
            {
                case RowLimitExceededMode.Throw:
                    throw new RowLimitExceededException(_current.RowCount + 1, MaxRowsPerSheet);
                case RowLimitExceededMode.SpillToNewSheet:
                    StartNewSheet();
                    break;
                case RowLimitExceededMode.Truncate:
                    _truncated = true;
                    return;
            }
        }
        _current.RowCount++;
        _current.Writer!.WriteStartElement("row");
        _current.Writer.WriteAttributeString("r", _current.RowCount.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>把 <c>_spillHeader</c> 作为当前表的首行写入（仅在 Spill 模式且有表头时由 StartNewSheet 调用）。</summary>
    private void WriteSpillHeaderRow()
    {
        if (_spillHeader is null || _spillHeader.Length == 0) return;
        WriteRowStart();
        int col = 1;
        foreach (var value in _spillHeader)
        {
            WriteCell(col, CellFactory.FromObject(value));
            col++;
        }
        _current.Writer!.WriteEndElement();
    }

    private void WriteCell(int col, Cell cell)
    {
        if (cell is null || cell.IsEmpty) return;

        var reference = CellRef.ToString(_current.RowCount - 1, col - 1);
        var styleId = _stylesheet.GetOrCreateXfId(cell.Style, cell.NumberFormat);
        if (cell.Hyperlink is not null)
            _current.Hyperlinks.Add((reference, cell.Hyperlink.Target, cell.Hyperlink.Tooltip, cell.Hyperlink.IsInternal));
        var styleAttr = styleId > 0 ? styleId.ToString(CultureInfo.InvariantCulture) : null;
        var formula = cell.Formula ?? (cell.IsFormula ? cell.Text : null);
        if (!string.IsNullOrEmpty(formula))
        {
            _current.Writer!.WriteStartElement("c");
            _current.Writer.WriteAttributeString("r", reference);
            if (styleAttr is not null) _current.Writer.WriteAttributeString("s", styleAttr);
            if (cell.Type == CellType.Boolean) _current.Writer.WriteAttributeString("t", "b");
            _current.Writer.WriteStartElement("f");
            _current.Writer.WriteString(formula);
            _current.Writer.WriteEndElement();
            if (cell.Type is CellType.Number or CellType.Date or CellType.Boolean)
            {
                _current.Writer.WriteStartElement("v");
                _current.Writer.WriteString(cell.Type == CellType.Number ? cell.Number.ToString(CultureInfo.InvariantCulture) :
                    cell.Type == CellType.Date ? cell.Date.ToOADate().ToString(CultureInfo.InvariantCulture) : cell.Boolean ? "1" : "0");
                _current.Writer.WriteEndElement();
            }
            _current.Writer.WriteEndElement();
            return;
        }

        switch (cell.Type)
        {
            case CellType.Text:
                _current.Writer!.WriteStartElement("c");
                _current.Writer.WriteAttributeString("r", reference);
                if (styleAttr is not null) _current.Writer.WriteAttributeString("s", styleAttr);
                _current.Writer.WriteAttributeString("t", "inlineStr");
                _current.Writer.WriteStartElement("is");
                _current.Writer.WriteStartElement("t");
                if (cell.Text is { Length: > 0 } && (char.IsWhiteSpace(cell.Text[0]) || char.IsWhiteSpace(cell.Text[cell.Text.Length - 1])))
                    _current.Writer.WriteAttributeString("xml", "space", null, "preserve");
                _current.Writer.WriteString(cell.Text ?? "");
                _current.Writer.WriteEndElement(); // t
                _current.Writer.WriteEndElement(); // is
                _current.Writer.WriteEndElement(); // c
                break;

            case CellType.Number:
                _current.Writer!.WriteStartElement("c");
                _current.Writer.WriteAttributeString("r", reference);
                if (styleAttr is not null) _current.Writer.WriteAttributeString("s", styleAttr);
                _current.Writer.WriteStartElement("v");
                _current.Writer.WriteString(cell.Number.ToString(CultureInfo.InvariantCulture));
                _current.Writer.WriteEndElement();
                _current.Writer.WriteEndElement();
                break;

            case CellType.Date:
                _current.Writer!.WriteStartElement("c");
                _current.Writer.WriteAttributeString("r", reference);
                if (styleAttr is not null) _current.Writer.WriteAttributeString("s", styleAttr);
                _current.Writer.WriteStartElement("v");
                _current.Writer.WriteString(cell.Date.ToOADate().ToString(CultureInfo.InvariantCulture));
                _current.Writer.WriteEndElement();
                _current.Writer.WriteEndElement();
                break;

            case CellType.Boolean:
                _current.Writer!.WriteStartElement("c");
                _current.Writer.WriteAttributeString("r", reference);
                if (styleAttr is not null) _current.Writer.WriteAttributeString("s", styleAttr);
                _current.Writer.WriteAttributeString("t", "b");
                _current.Writer.WriteStartElement("v");
                _current.Writer.WriteString(cell.Boolean ? "1" : "0");
                _current.Writer.WriteEndElement();
                _current.Writer.WriteEndElement();
                break;
        }
    }

    private void WriteStableHead()
    {
        WriteEntry("_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<Relationships xmlns=\"{RelNs}\">" +
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");
    }

    private void CloseCurrentSheet()
    {
        if (_current is null || _current.Writer is null) return;
        _current.Writer.WriteEndElement(); // sheetData
        WriteSheetHyperlinks(_current);
        _current.Writer.WriteEndElement(); // worksheet
        _current.Writer.WriteEndDocument();
        _current.Writer.Flush();
        _current.Writer.Dispose();
        _current.Stream?.Dispose();
        _current.Writer = null;
        _current.Stream = null;
    }

    private void StartNewSheet()
    {
        CloseCurrentSheet();
        int n = _sheets.Count + 1;
        var ctx = new SheetContext { Name = "Sheet" + n.ToString(CultureInfo.InvariantCulture), Entry = "xl/worksheets/sheet" + n.ToString(CultureInfo.InvariantCulture) + ".xml" };
        ctx.Stream = _zip.CreateEntry(ctx.Entry, CompressionLevel.Optimal).Open();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = false,
            Indent = false,
        };
        ctx.Writer = XmlWriter.Create(ctx.Stream!, settings);
        ctx.Writer.WriteStartDocument();
        ctx.Writer.WriteStartElement("worksheet", MainNs);
        ctx.Writer.WriteStartElement("sheetData");
        _sheets.Add(ctx);
        _current = ctx;
        // Spill 模式且有表头：每张表（含 Sheet1 与分表）首行写表头
        if (_mode == RowLimitExceededMode.SpillToNewSheet)
            WriteSpillHeaderRow();
    }

    private void WriteSheetHyperlinks(SheetContext ctx)
    {
        if (ctx.Hyperlinks.Count == 0) return;
        ctx.Writer!.WriteStartElement("hyperlinks");
        int external = 0;
        foreach (var link in ctx.Hyperlinks)
        {
            ctx.Writer.WriteStartElement("hyperlink");
            ctx.Writer.WriteAttributeString("ref", link.Ref);
            if (link.IsInternal)
                ctx.Writer.WriteAttributeString("location", link.Target.TrimStart('#'));
            else
            {
                external++;
                ctx.Writer.WriteAttributeString("r", "id", OfficeRelNs, $"rIdH{external}");
            }
            if (!string.IsNullOrEmpty(link.Tooltip)) ctx.Writer.WriteAttributeString("tooltip", link.Tooltip);
            ctx.Writer.WriteEndElement();
        }
        ctx.Writer.WriteEndElement();
    }

    private void WriteEntry(string name, string xml)
    {
        var entry = _zip.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        var bytes = new UTF8Encoding(false).GetBytes(xml);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string XmlEscape(string value) =>
        System.Security.SecurityElement.Escape(value) ?? string.Empty;

    /// <summary>关闭写入器并完成文件。写入后文件才能被正常读取 </summary>
    public void Close()
    {
        if (_closed) return;
        _closed = true;

        CloseCurrentSheet();

        for (int i = 0; i < _sheets.Count; i++)
        {
            var ctx = _sheets[i];
            if (!ctx.Hyperlinks.Any(h => !h.IsInternal)) continue;
            var rels = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{RelNs}\">");
            int external = 0;
            foreach (var link in ctx.Hyperlinks)
                if (!link.IsInternal)
                    rels.Append($"<Relationship Id=\"rIdH{++external}\" Type=\"{OfficeRelNs}/hyperlink\" Target=\"{XmlEscape(link.Target)}\" TargetMode=\"External\"/>");
            rels.Append("</Relationships>");
            WriteEntry($"xl/worksheets/_rels/sheet{i + 1}.xml.rels", rels.ToString());
        }

        WriteContentTypes();
        WriteWorkbookXml();
        WriteWorkbookRels();
        WriteEntry("xl/styles.xml", _stylesheet.BuildStylesXml());

        _zip.Dispose();
        if (_ownsStream)
            _underlying.Dispose();
    }

    private void WriteContentTypes()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
        sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"");
        sb.Append(_macroEnabled
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        sb.Append("\"/>");
        for (int i = 0; i < _sheets.Count; i++)
            sb.Append($"<Override PartName=\"/xl/worksheets/sheet{i + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
        sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
        sb.Append("</Types>");
        WriteEntry("[Content_Types].xml", sb.ToString());
    }

    private void WriteWorkbookXml()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\">");
        sb.Append("<sheets>");
        for (int i = 0; i < _sheets.Count; i++)
            sb.Append($"<sheet name=\"{XmlEscape(_sheets[i].Name)}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        sb.Append("</sheets>");
        sb.Append("</workbook>");
        WriteEntry("xl/workbook.xml", sb.ToString());
    }

    private void WriteWorkbookRels()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        for (int i = 0; i < _sheets.Count; i++)
            sb.Append($"<Relationship Id=\"rId{i + 1}\" Type=\"{OfficeRelNs}/worksheet\" Target=\"worksheets/sheet{i + 1}.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{_sheets.Count + 1}\" Type=\"{OfficeRelNs}/styles\" Target=\"styles.xml\"/>");
        sb.Append("</Relationships>");
        WriteEntry("xl/_rels/workbook.xml.rels", sb.ToString());
    }

    public void Dispose()
    {
        if (!_closed)
            Close();
    }
}
