using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 跨格式转换 Stage A：xlsb → xlsx/xlsm 时直通「格式无关」部件
/// （主题 / customXml / 媒体 / ActiveX / VML / printerSettings / docProps.custom / VBA），
/// 并重写关系与内容类型、剔除悬空引用。BIFF12 高级部件（透视/连接/切片/绘图 XML）暂不保留但须显式上报。
/// </summary>
public class XlsbToXlsxConversionTests
{
    private static string TempPath(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexcel_xlsb2ooxml_{Guid.NewGuid():N}{ext}");

    private const string ThemeCt = "application/vnd.openxmlformats-officedocument.theme+xml";
    private const string DrawingCt = "application/vnd.openxmlformats-officedocument.drawing+xml";
    private const string PivotCt = "application/vnd.ms-excel.pivotTable";
    private const string VbaCt = "application/vnd.ms-office.vbaProject";

    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static XlsbTestFile.WorkbookSpec SpecWithCell()
    {
        var spec = new XlsbTestFile.WorkbookSpec();
        spec.Sheets.Add(new XlsbTestFile.SheetSpec { Name = "S0" });
        spec.Sheets[0].Rows.Add(new XlsbTestFile.RowSpec
        {
            Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "hi" } },
        });
        return spec;
    }

    private static XDocument ReadXml(ZipArchive zip, string entry)
    {
        using var s = zip.GetEntry(entry)!.Open();
        return XDocument.Load(s);
    }

    [Fact]
    public void XlsbToXlsm_PreservesFormatAgnosticParts()
    {
        var spec = SpecWithCell();
        spec.ExtraParts["xl/theme/theme1.xml"] =
            Encoding.UTF8.GetBytes("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"/>");
        spec.ExtraParts["customXml/item1.xml"] = Encoding.UTF8.GetBytes("<root/>");
        spec.ExtraOverrides["/xl/theme/theme1.xml"] = ThemeCt;
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] =
            $"<Relationship Id=\"rIdT\" Type=\"{OfficeRelNs}/theme\" Target=\"theme/theme1.xml\"/>" +
            $"<Relationship Id=\"rIdC\" Type=\"{OfficeRelNs}/customXml\" Target=\"../customXml/item1.xml\"/>";

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsm");
        try
        {
            var wb = Excel.Open(src);
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsm);

            using var zip = ZipFile.OpenRead(outPath);
            Assert.NotNull(zip.GetEntry("xl/theme/theme1.xml"));
            Assert.NotNull(zip.GetEntry("customXml/item1.xml"));

            var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels").ToString();
            Assert.Contains("theme/theme1.xml", rels);
            Assert.Contains("../customXml/item1.xml", rels);
            Assert.DoesNotContain(".bin", rels);

            // 内容类型保留 theme Override
            var ct = ReadXml(zip, "[Content_Types].xml").ToString();
            Assert.Contains("/xl/theme/theme1.xml", ct);

            AssertNoDangling(zip);
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToXlsm_DropsDrawingXml_KeepsVml()
    {
        var spec = SpecWithCell();
        spec.ExtraParts["xl/drawings/drawing1.xml"] =
            Encoding.UTF8.GetBytes("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>");
        spec.ExtraParts["xl/drawings/vmlDrawing1.vml"] =
            Encoding.UTF8.GetBytes("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"/>");
        spec.ExtraDefaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        spec.ExtraOverrides["/xl/drawings/drawing1.xml"] = DrawingCt;
        spec.ExtraRels["xl/worksheets/_rels/sheet1.bin.rels"] =
            $"<Relationship Id=\"rIdD\" Type=\"{OfficeRelNs}/drawing\" Target=\"../drawings/drawing1.xml\"/>" +
            $"<Relationship Id=\"rIdV\" Type=\"{OfficeRelNs}/vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/>";

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsm");
        try
        {
            var wb = Excel.Open(src);
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsm);

            using var zip = ZipFile.OpenRead(outPath);
            // xlsb 专有 drawing XML 会被 Excel 拒绝，故 Stage A 排除；VML 保留。
            Assert.Null(zip.GetEntry("xl/drawings/drawing1.xml"));
            Assert.NotNull(zip.GetEntry("xl/drawings/vmlDrawing1.vml"));

            var sheetRels = ReadXml(zip, "xl/worksheets/_rels/sheet1.xml.rels").ToString();
            Assert.DoesNotContain("drawing1.xml", sheetRels);

            AssertNoDangling(zip);
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToXlsx_ReportsAdvancedPartsDegradation()
    {
        var spec = SpecWithCell();
        spec.ExtraParts["xl/pivotTables/pivotTable1.bin"] = new byte[] { 1, 2, 3, 4 };
        spec.ExtraOverrides["/xl/pivotTables/pivotTable1.bin"] = PivotCt;
        spec.ExtraRels["xl/worksheets/_rels/sheet1.bin.rels"] =
            $"<Relationship Id=\"rIdP\" Type=\"{OfficeRelNs}/pivotTable\" Target=\"../pivotTables/pivotTable1.bin\"/>";

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsx);

            Assert.Contains(wb.SaveDegradations, d => d.Capability == DegradationCapability.PivotTables);

            using var zip = ZipFile.OpenRead(outPath);
            Assert.Null(zip.GetEntry("xl/pivotTables/pivotTable1.bin"));
            AssertNoDangling(zip);
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToXlsx_StrictMode_BlocksWhenAdvancedParts()
    {
        var spec = SpecWithCell();
        spec.ExtraParts["xl/pivotTables/pivotTable1.bin"] = new byte[] { 1, 2, 3, 4 };
        spec.ExtraOverrides["/xl/pivotTables/pivotTable1.bin"] = PivotCt;
        spec.ExtraRels["xl/worksheets/_rels/sheet1.bin.rels"] =
            $"<Relationship Id=\"rIdP\" Type=\"{OfficeRelNs}/pivotTable\" Target=\"../pivotTables/pivotTable1.bin\"/>";

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            wb.AllowFeatureLossOnSave = false;
            Assert.Throws<LiteExcelException>(() => wb.SaveAs(outPath, ExcelFormat.Xlsx));
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToOoxml_VbaKeptForXlsm_DroppedForXlsx()
    {
        var spec = SpecWithCell();
        spec.ExtraParts["xl/vbaProject.bin"] = new byte[] { 0xCF, 0x11, 0xE0, 0xA1 };
        spec.ExtraOverrides["/xl/vbaProject.bin"] = VbaCt;
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] =
            "<Relationship Id=\"rIdVba\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>";

        var src = XlsbTestFile.Build(spec);
        var xlsm = TempPath(".xlsm");
        var xlsx = TempPath(".xlsx");
        try
        {
            var wb1 = Excel.Open(src);
            wb1.AllowFeatureLossOnSave = true;
            wb1.SaveAs(xlsm, ExcelFormat.Xlsm);
            using (var zip = ZipFile.OpenRead(xlsm))
            {
                Assert.NotNull(zip.GetEntry("xl/vbaProject.bin"));
                AssertNoDangling(zip);
            }

            var wb2 = Excel.Open(src);
            wb2.AllowFeatureLossOnSave = true;
            wb2.SaveAs(xlsx, ExcelFormat.Xlsx);
            using (var zip = ZipFile.OpenRead(xlsx))
                Assert.Null(zip.GetEntry("xl/vbaProject.bin"));
            Assert.Contains(wb2.SaveDegradations, d => d.Capability == DegradationCapability.Macros);
        }
        finally { Cleanup(src, xlsm); Cleanup(src, xlsx); }
    }

    [Fact]
    public void XlsbToXlsx_TranscodesBasicStyles()
    {
        // B2：xlsb styles.bin 的字体/填充/边框应转码为 xlsx styles.xml 并作用于单元格。
        var spec = SpecWithCell();
        spec.Fonts = new List<XlsbTestFile.FontSpec>
        {
            new() { Name = "Calibri", Size = 11 },
            new() { Name = "Arial", Size = 14, Bold = true, ColorRgb = "FF0000" },
        };
        spec.Fills = new List<string?> { null, null, "FFFF00" }; // 索引 0/1 保留，2 = 黄底
        spec.Borders = new List<bool> { false, true };
        spec.CellXfs.Add(0);
        spec.CellXfs.Add(0);
        spec.CellXfRefs.Add((0, 0, 0)); // 默认
        spec.CellXfRefs.Add((1, 2, 1)); // Arial 14 粗体红字 + 黄底 + thin 边框
        spec.Sheets[0].Rows[0].Cells[0].Style = 1;

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            var cell = wb.Worksheets[0].Cell("A1");
            Assert.NotNull(cell.Style);
            Assert.Equal("Arial", cell.Style!.FontName);
            Assert.True(cell.Style.Bold);
            Assert.Equal("#FF0000", cell.Style.FontColor);
            Assert.Equal("#FFFF00", cell.Style.FillColor);
            Assert.Equal("thin", cell.Style.Border!.Top!.Style);

            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsx);

            using var zip = ZipFile.OpenRead(outPath);
            var styles = ReadXml(zip, "xl/styles.xml").ToString();
            Assert.Contains("Arial", styles);
            Assert.Contains("FF0000", styles);
            Assert.Contains("FFFF00", styles);
            Assert.Contains("thin", styles);
            AssertNoDangling(zip);
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToXlsx_TranscodesConnectionsAndQueryTables()
    {
        var spec = SpecWithCell();
        var conn1 = XlsbTestFile.BuildConnectionBin(5, 1, "Query - X",
            dbConn: "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=X;Extended Properties=\"\"",
            dbCmd: "SELECT * FROM [X]");
        var conn2 = XlsbTestFile.BuildConnectionBin(102, 2, "LinkedTable_T",
            x15Id: "T", sourceName: "_xlcn.LinkedTable_T");
        spec.ExtraParts["xl/connections.bin"] = conn1.Concat(conn2).ToArray();
        spec.ExtraOverrides["/xl/connections.bin"] = "application/vnd.ms-excel.connections";
        spec.ExtraRels["xl/_rels/workbook.bin.rels"] =
            $"<Relationship Id=\"rIdConn\" Type=\"{OfficeRelNs}/connections\" Target=\"connections.bin\"/>";

        spec.ExtraParts["xl/queryTables/queryTable1.bin"] = XlsbTestFile.BuildQueryTableBin(1, "ExternalData_1",
            new[] { (1u, 5u, "FACTORY") });
        spec.ExtraOverrides["/xl/queryTables/queryTable1.bin"] = "application/vnd.ms-excel.queryTable";
        spec.ExtraRels["xl/worksheets/_rels/sheet1.bin.rels"] =
            $"<Relationship Id=\"rIdQ\" Type=\"{OfficeRelNs}/queryTable\" Target=\"../queryTables/queryTable1.bin\"/>";

        var src = XlsbTestFile.Build(spec);
        var outPath = TempPath(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsx);

            using var zip = ZipFile.OpenRead(outPath);
            Assert.NotNull(zip.GetEntry("xl/connections.xml"));
            Assert.NotNull(zip.GetEntry("xl/queryTables/queryTable1.xml"));

            var connXml = ReadXml(zip, "xl/connections.xml").ToString();
            Assert.Contains("Query - X", connXml);
            Assert.Contains("LinkedTable_T", connXml);
            Assert.Contains("_xlcn.LinkedTable_T", connXml);

            var wbXml = ReadXml(zip, "xl/workbook.xml").ToString();
            Assert.Contains("definedNames", wbXml);
            Assert.Contains("_xlcn.LinkedTable_T", wbXml);

            var qtXml = ReadXml(zip, "xl/queryTables/queryTable1.xml").ToString();
            Assert.Contains("ExternalData_1", qtXml);
            Assert.Contains("FACTORY", qtXml);

            AssertNoDangling(zip);
        }
        finally { Cleanup(src, outPath); }
    }

    [Fact]
    public void XlsbToXlsx_EmitsDefinedNames()
    {
        // B1：xlsb 打开后读回的定义名（rgce 可解码的简单引用）应写出到 workbook.xml 的 definedNames。
        var fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-namedranges.xlsb");
        Assert.True(File.Exists(fixture), $"缺少 fixture: {fixture}");

        var outPath = TempPath(".xlsx");
        try
        {
            var wb = Excel.Open(fixture);
            Assert.NotEmpty(wb.Names);
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsx);

            using var zip = ZipFile.OpenRead(outPath);
            var wbXml = ReadXml(zip, "xl/workbook.xml").ToString();
            Assert.Contains("<definedNames>", wbXml);
            Assert.Contains("G_Area", wbXml);
            Assert.Contains("G_BetaRef", wbXml);
            AssertNoDangling(zip);
        }
        finally { if (File.Exists(outPath)) File.Delete(outPath); }
    }

    /// <summary>校验输出包：所有关系目标可解析、每个部件都有内容类型。</summary>
    private static void AssertNoDangling(ZipArchive zip)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var e in zip.Entries) names.Add(e.FullName);

        var ct = ReadXml(zip, "[Content_Types].xml");
        var ns = ct.Root!.GetDefaultNamespace();
        var defaults = ct.Root.Elements(ns + "Default")
            .ToDictionary(x => ((string)x.Attribute("Extension")!).ToLowerInvariant(), x => x);
        var overrides = ct.Root.Elements(ns + "Override")
            .Select(x => ((string)x.Attribute("PartName")!).TrimStart('/'))
            .ToHashSet(StringComparer.Ordinal);

        foreach (var name in names)
        {
            if (name == "[Content_Types].xml") continue;
            var ext = Path.GetExtension(name).TrimStart('.').ToLowerInvariant();
            Assert.True(overrides.Contains(name) || defaults.ContainsKey(ext),
                $"部件缺少内容类型声明: {name}");
        }
        foreach (var part in overrides)
            Assert.True(names.Contains(part), $"Content_Types Override 指向不存在的部件: {part}");

        foreach (var relsName in names.Where(n => n.EndsWith(".rels", StringComparison.Ordinal)))
        {
            var baseDir = BaseDir(relsName);
            var relsNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");
            foreach (var rel in ReadXml(zip, relsName).Root!.Elements(relsNs + "Relationship"))
            {
                if ((string?)rel.Attribute("TargetMode") == "External") continue;
                var abs = Resolve(baseDir, (string)rel.Attribute("Target")!);
                Assert.True(names.Contains(abs), $"关系悬空: {relsName} -> {rel.Attribute("Target")} (解析为 {abs})");
            }
        }
    }

    private static string BaseDir(string relsPath)
    {
        var i = relsPath.IndexOf("/_rels/", StringComparison.Ordinal);
        return i < 0 ? "" : relsPath.Substring(0, i);
    }

    private static string Resolve(string baseDir, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/", StringComparison.Ordinal)) return target.TrimStart('/');
        var combined = string.IsNullOrEmpty(baseDir) ? target : baseDir + "/" + target;
        var stack = new List<string>();
        foreach (var p in combined.Split('/'))
        {
            if (p.Length == 0 || p == ".") continue;
            if (p == "..") { if (stack.Count > 0) stack.RemoveAt(stack.Count - 1); continue; }
            stack.Add(p);
        }
        return string.Join("/", stack);
    }

    private static void Cleanup(string src, string outPath)
    {
        if (File.Exists(src)) File.Delete(src);
        if (File.Exists(outPath)) File.Delete(outPath);
    }
}
