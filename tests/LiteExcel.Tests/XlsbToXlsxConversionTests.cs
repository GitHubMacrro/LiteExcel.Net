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
