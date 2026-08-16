using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace LiteExcel;

/// <summary>
/// 流式写入器：逐行写入大文件，不驻留内存。
/// 采用内联字符串（inlineStr），避免共享字符串表预扫描。
/// 仅支持单工作表；样式/合并/筛选等高级能力不支持（与大文件场景定位一致）。
/// 使用后必须调用 <see cref="Dispose"/> 或 <see cref="Close"/> 完成文件。
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

        switch (cell.Type)
        {
            case CellType.Text:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
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
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Number.ToString(CultureInfo.InvariantCulture));
                _sheetWriter.WriteEndElement();
                _sheetWriter.WriteEndElement();
                break;

            case CellType.Date:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
                _sheetWriter.WriteStartElement("v");
                _sheetWriter.WriteString(cell.Date.ToOADate().ToString(CultureInfo.InvariantCulture));
                _sheetWriter.WriteEndElement();
                _sheetWriter.WriteEndElement();
                break;

            case CellType.Boolean:
                _sheetWriter.WriteStartElement("c");
                _sheetWriter.WriteAttributeString("r", reference);
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

        WriteEntry("xl/styles.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<styleSheet xmlns=\"{MainNs}\">" +
            "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders>" +
            "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>" +
            "<cellXfs count=\"1\"><xf/></cellXfs>" +
            "</styleSheet>");
    }

    private void WriteEntry(string name, string xml)
    {
        var entry = _zip.CreateEntry(name, CompressionLevel.Optimal);
        using var stream = entry.Open();
        var bytes = new UTF8Encoding(false).GetBytes(xml);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>关闭写入器并完成文件。写入后文件才能被正常读取 </summary>
    public void Close()
    {
        if (_closed) return;
        _closed = true;

        _sheetWriter.WriteEndElement(); // sheetData
        _sheetWriter.WriteEndElement(); // worksheet
        _sheetWriter.WriteEndDocument();
        _sheetWriter.Flush();
        _sheetWriter.Dispose();
        _sheetStream.Dispose();

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
