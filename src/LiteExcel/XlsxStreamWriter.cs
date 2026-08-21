using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LiteExcel;

/// <summary>
/// 流式写入器：逐行写入大文件，不驻留内存。
/// 采用内联字符串（inlineStr），避免共享字符串表预扫描。
/// 支持单工作表；样式/公式/超链接随行写入（styles.xml 与 sheet rels 在 Close 时统一写出）。
/// 合并/筛选/图片等高级能力不支持。
/// 使用后必须调用 <see cref="Dispose"/> 或 <see cref="Close"/> 完成文件。
/// 注意：超链接数量极大时内存不再恒定（内部缓冲全部超链接引用）。
/// </summary>
public sealed class XlsxStreamWriter : IDisposable
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private readonly ZipArchive _zip;
    private readonly Stream _underlying;
    private readonly bool _ownsStream;
    private readonly Stream _sheetStream;
    private readonly XmlWriter _sheetWriter;
    private readonly bool _macroEnabled;
    private readonly Internal.Stylesheet _stylesheet = new();
    private readonly List<(string Ref, string Target, string? Tooltip, bool IsInternal)> _hyperlinks = new();
    private int _currentRow;
    private bool _closed;

    private XlsxStreamWriter(Stream stream, bool ownsStream, bool macroEnabled = false)
    {
        _underlying = stream;
        _ownsStream = ownsStream;
        _macroEnabled = macroEnabled;
        _zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        WritePackageHead();
        _sheetStream = _zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal).Open();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = false,
            Indent = false,
        };
        _sheetWriter = XmlWriter.Create(_sheetStream, settings);
        _sheetWriter.WriteStartDocument();
        _sheetWriter.WriteStartElement("worksheet", MainNs);
        _sheetWriter.WriteStartElement("sheetData");
    }

    /// <summary>创建流式写入器（写入指定路径，覆盖已存在文件） </summary>
    public static XlsxStreamWriter Create(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("路径不能为空", nameof(path));
        var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
        return new XlsxStreamWriter(fs, ownsStream: true,
            macroEnabled: string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>创建流式写入器（写入流，LeaveOpen 由调用方管理） </summary>
    public static XlsxStreamWriter Create(Stream stream)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("流不可写", nameof(stream));
        return new XlsxStreamWriter(stream, ownsStream: false);
    }

    /// <summary>写入一行单元格值 </summary>
    public void WriteRow(IEnumerable<object?> values)
    {
        if (values is null) throw new ArgumentNullException(nameof(values));
        if (_closed) throw new InvalidOperationException("写入器已关闭");

        WriteRowStart();
        int col = 1;
        foreach (var value in values)
        {
            WriteCell(col, CellFactory.FromObject(value));
            col++;
        }
        _sheetWriter.WriteEndElement();
    }

    /// <summary>写入一行 Cell（低层模型） </summary>
    public void WriteRow(IEnumerable<Cell> cells)
    {
        if (cells is null) throw new ArgumentNullException(nameof(cells));
        if (_closed) throw new InvalidOperationException("写入器已关闭");

        WriteRowStart();
        int col = 1;
        foreach (var cell in cells)
        {
            WriteCell(col, cell);
            col++;
        }
        _sheetWriter.WriteEndElement();
    }

    private void WriteRowStart()
    {
        _currentRow++;
        _sheetWriter.WriteStartElement("row");
        _sheetWriter.WriteAttributeString("r", _currentRow.ToString(CultureInfo.InvariantCulture));
    }

    private void WriteCell(int col, Cell cell)
    {
        if (cell is null || cell.IsEmpty) return;

        var reference = CellRef.ToString(_currentRow - 1, col - 1);
        var styleId = _stylesheet.GetOrCreateXfId(cell.Style, cell.NumberFormat);
        if (cell.Hyperlink is not null)
            _hyperlinks.Add((reference, cell.Hyperlink.Target, cell.Hyperlink.Tooltip, cell.Hyperlink.IsInternal));
        var styleAttr = styleId > 0 ? styleId.ToString(CultureInfo.InvariantCulture) : null;
        var formula = cell.Formula ?? (cell.IsFormula ? cell.Text : null);
        if (!string.IsNullOrEmpty(formula))
        {
            _sheetWriter.WriteStartElement("c");
            _sheetWriter.WriteAttributeString("r", reference);
            if (styleAttr is not null) _sheetWriter.WriteAttributeString("s", styleAttr);
            if (cell.Type == CellType.Boolean) _sheetWriter.WriteAttributeString("t", "b");
            _sheetWriter.WriteStartElement("f");
            _sheetWriter.WriteString(formula);
            _sheetWriter.WriteEndElement();
            if (cell.Type is CellType.Number or CellType.Date or CellType.Boolean)
            {
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Type == CellType.Number ? cell.Number.ToString(CultureInfo.InvariantCulture) :
                    cell.Type == CellType.Date ? cell.Date.ToOADate().ToString(CultureInfo.InvariantCulture) : cell.Boolean ? "1" : "0");
                _sheetWriter.WriteEndElement();
            }
            _sheetWriter.WriteEndElement();
            return;
        }

        switch (cell.Type)
        {
            case CellType.Text:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
                if (styleAttr is not null) _sheetWriter.WriteAttributeString("s", styleAttr);
                _sheetWriter.WriteAttributeString("t", "inlineStr");
                _sheetWriter.WriteStartElement("is");
                _sheetWriter.WriteStartElement("t");
                if (cell.Text is { Length: > 0 } && (char.IsWhiteSpace(cell.Text[0]) || char.IsWhiteSpace(cell.Text[cell.Text.Length - 1])))
                    _sheetWriter.WriteAttributeString("xml", "space", null, "preserve");
                _sheetWriter.WriteString(cell.Text ?? "");
                _sheetWriter.WriteEndElement(); // t
                _sheetWriter.WriteEndElement(); // is
                _sheetWriter.WriteEndElement(); // c
                break;

            case CellType.Number:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
                if (styleAttr is not null) _sheetWriter.WriteAttributeString("s", styleAttr);
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Number.ToString(CultureInfo.InvariantCulture));
                _sheetWriter.WriteEndElement();
                _sheetWriter.WriteEndElement();
                break;

            case CellType.Date:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
                if (styleAttr is not null) _sheetWriter.WriteAttributeString("s", styleAttr);
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Date.ToOADate().ToString(CultureInfo.InvariantCulture));
                _sheetWriter.WriteEndElement();
                _sheetWriter.WriteEndElement();
                break;

            case CellType.Boolean:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
                if (styleAttr is not null) _sheetWriter.WriteAttributeString("s", styleAttr);
                _sheetWriter.WriteAttributeString("t", "b");
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Boolean ? "1" : "0");
                _sheetWriter.WriteEndElement();
                _sheetWriter.WriteEndElement();
                break;
        }
    }

    private void WritePackageHead()
    {
        WriteEntry("[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"" +
            (_macroEnabled
                ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
                : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") +
            "\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "</Types>");

        WriteEntry("_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<Relationships xmlns=\"{RelNs}\">" +
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");

        WriteEntry("xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\">" +
            "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
            "</workbook>");

        WriteEntry("xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<Relationships xmlns=\"{RelNs}\">" +
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            $"<Relationship Id=\"rId2\" Type=\"{OfficeRelNs}/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>");

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

        _sheetWriter.WriteEndElement(); // sheetData
        if (_hyperlinks.Count > 0)
        {
            _sheetWriter.WriteStartElement("hyperlinks");
            int external = 0;
            foreach (var link in _hyperlinks)
            {
                _sheetWriter.WriteStartElement("hyperlink");
                _sheetWriter.WriteAttributeString("ref", link.Ref);
                if (link.IsInternal)
                    _sheetWriter.WriteAttributeString("location", link.Target.TrimStart('#'));
                else
                {
                    external++;
                    _sheetWriter.WriteAttributeString("r", "id", OfficeRelNs, $"rIdH{external}");
                }
                if (!string.IsNullOrEmpty(link.Tooltip)) _sheetWriter.WriteAttributeString("tooltip", link.Tooltip);
                _sheetWriter.WriteEndElement();
            }
            _sheetWriter.WriteEndElement();
        }
        _sheetWriter.WriteEndElement(); // worksheet
        _sheetWriter.WriteEndDocument();
        _sheetWriter.Flush();
        _sheetWriter.Dispose();
        _sheetStream.Dispose();

        if (_hyperlinks.Any(h => !h.IsInternal))
        {
            var rels = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"{RelNs}\">");
            int external = 0;
            foreach (var link in _hyperlinks)
                if (!link.IsInternal)
                    rels.Append($"<Relationship Id=\"rIdH{++external}\" Type=\"{OfficeRelNs}/hyperlink\" Target=\"{XmlEscape(link.Target)}\" TargetMode=\"External\"/>");
            rels.Append("</Relationships>");
            WriteEntry("xl/worksheets/_rels/sheet1.xml.rels", rels.ToString());
        }
        WriteEntry("xl/styles.xml", _stylesheet.BuildStylesXml());

        _zip.Dispose();
        if (_ownsStream)
            _underlying.Dispose();
    }

    public void Dispose()
    {
        if (!_closed)
            Close();
    }
}
