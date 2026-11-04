using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using LiteExcel;

namespace LiteExcel.Tests;

public class PrefixedNamespaceTests
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static TheoryData<string, string, string, string, string> NamespaceCases => new()
    {
        { "", "", "", "", "r" },
        { "book", "", "", "", "r" },
        { "", "style", "", "", "r" },
        { "", "", "grid", "", "r" },
        { "", "", "", "strings", "r" },
        { "x", "x", "x", "x", "r" },
        { "book", "style", "grid", "strings", "link" },
        { "a1", "b2", "c3", "d4", "relationship" },
    };

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void GetSheetNames_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        Assert.Equal(new[] { "Sheet1" }, XlsxReader.GetSheetNames(stream));
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void ReadByIndex_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        AssertSheet(XlsxReader.Read(stream, 0));
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void ReadByName_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        AssertSheet(XlsxReader.Read(stream, "Sheet1"));
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void Open_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        var opened = Excel.Open(stream, ExcelFormat.Xlsx);
        var sheet = Assert.Single(opened.Worksheets);
        Assert.Equal("Sheet1", sheet.Name);
        Assert.Equal("Text", sheet.Cell("A1").Text);
        AssertRow(new[] { sheet.Cell("A2"), sheet.Cell("B2"), sheet.Cell("C2"), sheet.Cell("D2") });
        Assert.Equal("SampleApp", opened.Properties.Application);
        Assert.Equal(2, opened.Names.Count);
        var globalName = Assert.Single(opened.Names, name => name.Name == "SampleRange");
        Assert.Equal("Sheet1!$A$2:$C$2", globalName.Reference);
        Assert.Equal(-1, globalName.LocalSheetId);
        var localName = Assert.Single(opened.Names, name => name.Name == "LocalRange");
        Assert.Equal("Sheet1!$C$2", localName.Reference);
        Assert.Equal(0, localName.LocalSheetId);
        var conditionalFormat = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal("C2", conditionalFormat.Sqref);
        Assert.Equal(ConditionalFormatType.CellIs, conditionalFormat.Type);
        Assert.Equal(ConditionalOperator.GreaterThan, conditionalFormat.Operator);
        Assert.Equal("10", conditionalFormat.Formula);
        Assert.NotNull(conditionalFormat.Style);
        Assert.Equal("#FF0000", conditionalFormat.Style.FontColor);
        Assert.Equal("#FFF2CC", conditionalFormat.Style.FillColor);
        stream.Position = 0;
        Assert.Equal("SampleApp", XlsxReader.ReadProperties(stream).Application);
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void StreamRows_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        var rows = new List<IReadOnlyList<Cell>>();
        XlsxReader.StreamRows(stream, "Sheet1", rows.Add);
        AssertRow(Assert.Single(rows));
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void ExcelStreamRows_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        var rows = new List<IReadOnlyList<Cell>>();
        Excel.StreamRows(stream, ExcelFormat.Xlsx, "Sheet1", rows.Add);
        AssertRow(Assert.Single(rows));
    }

    [Theory]
    [MemberData(nameof(NamespaceCases))]
    public void EnumerateRows_MatchesDefaultNamespace(string workbook, string styles, string worksheet, string sharedStrings, string relationship)
    {
        using var stream = CreatePackage(workbook, styles, worksheet, sharedStrings, relationship);
        var rows = Excel.EnumerateRows(stream, "Sheet1").ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "Text", "Date", "Number", "Inline" }, rows[0].Select(cell => cell.Text));
        AssertRow(rows[1]);
    }

    [Theory]
    [InlineData("", "r")]
    [InlineData("", "link")]
    [InlineData("grid", "link")]
    [InlineData("arbitrary", "relationship")]
    public void Open_TableRelationshipPrefix_MatchesDefaultNamespace(string worksheet, string relationship)
    {
        using var stream = CreatePackage("", "", worksheet, "", relationship, includeTable: true);
        var opened = Excel.Open(stream, ExcelFormat.Xlsx);
        var table = Assert.Single(opened.Worksheets[0].Tables);
        Assert.Equal("TableOne", table.Name);
        Assert.Equal("A1:D2", table.Ref);
        Assert.Equal(new[] { "Text", "Date", "Number", "Inline" }, table.Columns.Select(column => column.Name));
        Assert.Equal(TableStyleStyle.Medium6, table.Style);
        Assert.True(table.ShowRowStripes);
        Assert.True(table.AutoFilter);
    }

    private static void AssertSheet(SheetData sheet)
    {
        Assert.Equal("Sheet1", sheet.SheetName);
        Assert.Equal(new[] { "Text", "Date", "Number", "Inline" }, sheet.Headers);
        AssertRow(Assert.Single(sheet.Rows));
    }

    private static void AssertRow(IReadOnlyList<Cell> row)
    {
        Assert.Equal(4, row.Count);
        Assert.Equal(CellType.Text, row[0].Type);
        Assert.Equal("sample text", row[0].Text);
        Assert.Equal(CellType.Date, row[1].Type);
        Assert.Equal(new DateTime(2024, 1, 1), row[1].Date);
        Assert.Equal("yyyy-mm-dd", row[1].NumberFormat);
        Assert.Equal(CellType.Number, row[2].Type);
        Assert.Equal(12.5, row[2].Number);
        Assert.Equal("0.000", row[2].NumberFormat);
        Assert.Equal(CellType.Text, row[3].Type);
        Assert.Equal("inline text", row[3].Text);
    }

    private static MemoryStream CreatePackage(string workbook, string styles, string worksheet, string sharedStrings, string relationship, bool includeTable = false)
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddEntry(zip, "[Content_Types].xml", $"""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
                  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
                  {(includeTable ? "<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>" : "")}
                </Types>
                """);
            AddEntry(zip, "_rels/.rels", $"""
                <Relationships xmlns="{PackageRelNs}"><Relationship Id="root" Type="{OfficeRelNs}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="app" Type="{OfficeRelNs}/extended-properties" Target="docProps/app.xml"/></Relationships>
                """);
            AddEntry(zip, "docProps/app.xml", WithNamespace("<Properties><Application>SampleApp</Application></Properties>",
                workbook, "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"));
            AddEntry(zip, "xl/_rels/workbook.xml.rels", $"""
                <Relationships xmlns="{PackageRelNs}">
                  <Relationship Id="sheet" Type="{OfficeRelNs}/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="styles" Type="{OfficeRelNs}/styles" Target="styles.xml"/>
                  <Relationship Id="strings" Type="{OfficeRelNs}/sharedStrings" Target="sharedStrings.xml"/>
                </Relationships>
                """);
            AddEntry(zip, "xl/workbook.xml", WithNamespace($"""
                <workbook xmlns:{relationship}="{OfficeRelNs}"><sheets><sheet name="Sheet1" sheetId="1" {relationship}:id="sheet"/></sheets><definedNames><definedName name="SampleRange">Sheet1!$A$2:$C$2</definedName><definedName name="LocalRange" localSheetId="0">Sheet1!$C$2</definedName></definedNames></workbook>
                """, workbook));
            AddEntry(zip, "xl/styles.xml", WithNamespace("""
                <styleSheet>
                  <numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="0.000"/></numFmts>
                  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
                  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
                  <borders count="1"><border/></borders>
                  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
                  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>
                  <dxfs count="1"><dxf><font><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFF2CC"/></patternFill></fill></dxf></dxfs>
                </styleSheet>
                """, styles));
            AddEntry(zip, "xl/sharedStrings.xml", WithNamespace("""
                <sst count="5" uniqueCount="5"><si><t>Text</t></si><si><t>Date</t></si><si><t>Number</t></si><si><t>Inline</t></si><si><r><t xml:space="preserve">sample </t></r><r><t>text</t></r></si></sst>
                """, sharedStrings));
            AddEntry(zip, "xl/worksheets/sheet1.xml", WithNamespace($"""
                <worksheet xmlns:{relationship}="{OfficeRelNs}">
                  <dimension ref="A1:D2"/>
                  <sheetData>
                    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c><c r="D1" t="s"><v>3</v></c></row>
                    <row r="2"><c r="A2" t="s"><v>4</v></c><c r="B2" s="1"><v>45292</v></c><c r="C2" s="2"><v>12.5</v></c><c r="D2" t="inlineStr"><is><t>inline text</t></is></c></row>
                  </sheetData>
                  <conditionalFormatting sqref="C2"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>
                  {(includeTable ? $"<tableParts count=\"1\"><tablePart {relationship}:id=\"table\"/></tableParts>" : "")}
                </worksheet>
                """, worksheet));
            if (includeTable)
            {
                AddEntry(zip, "xl/worksheets/_rels/sheet1.xml.rels", $"""
                    <Relationships xmlns="{PackageRelNs}"><Relationship Id="table" Type="{OfficeRelNs}/table" Target="../tables/table1.xml"/></Relationships>
                    """);
                AddEntry(zip, "xl/tables/table1.xml", WithNamespace("""
                    <table id="1" name="TableOne" displayName="TableOne" ref="A1:D2" totalsRowShown="0"><autoFilter ref="A1:D2"/><tableColumns count="4"><tableColumn id="1" name="Text"/><tableColumn id="2" name="Date"/><tableColumn id="3" name="Number"/><tableColumn id="4" name="Inline"/></tableColumns><tableStyleInfo name="TableStyleMedium6" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>
                    """, worksheet));
            }
        }
        stream.Position = 0;
        return stream;
    }

    private static string WithNamespace(string xml, string prefix, string namespaceUri = MainNs)
    {
        if (prefix.Length != 0)
            xml = Regex.Replace(xml, @"<(/?)([A-Za-z][A-Za-z0-9]*)", $"<$1{prefix}:$2");
        var rootEnd = xml.IndexOf('>');
        var declaration = prefix.Length == 0 ? $" xmlns=\"{namespaceUri}\"" : $" xmlns:{prefix}=\"{namespaceUri}\"";
        return xml.Insert(rootEnd, declaration);
    }

    private static void AddEntry(ZipArchive zip, string path, string xml)
    {
        using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }
}
