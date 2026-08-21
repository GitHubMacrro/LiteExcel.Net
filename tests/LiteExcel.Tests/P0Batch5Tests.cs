using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 5：xlsb 保真透传（P0-14 xlsb）、xls/xlsb 文档属性（P0-18 xlsb）。
/// </summary>
public class P0Batch5Tests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"p0b5_{Guid.NewGuid():N}{ext}");

    private static string ReadEntry(string file, string entry)
    {
        using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var r = new StreamReader(e.Open(), Encoding.UTF8);
        return r.ReadToEnd();
    }

    /// <summary>向 xlsb 注入 chart 部件 + workbook.bin.rels 关系 + content types 声明 </summary>
    private static void InjectChartPart(string file)
    {
        using var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update);

        var chart = zip.CreateEntry("xl/charts/chart1.xml");
        using (var s = chart.Open())
        {
            var bytes = Encoding.UTF8.GetBytes("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");
            s.Write(bytes, 0, bytes.Length);
        }
        var wbRels = zip.GetEntry("xl/_rels/workbook.bin.rels")!;
        using (var s = wbRels.Open())
        {
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            doc.Root.Add(new XElement(ns + "Relationship",
                new XAttribute("Id", "rIdC1"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                new XAttribute("Target", "charts/chart1.xml")));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }
        var ct = zip.GetEntry("[Content_Types].xml")!;
        using (var s = ct.Open())
        {
            var doc = XDocument.Load(s);
            var ns = doc.Root!.GetDefaultNamespace();
            doc.Root.Add(new XElement(ns + "Override",
                new XAttribute("PartName", "/xl/charts/chart1.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
            s.Position = 0;
            s.SetLength(0);
            doc.Save(s);
        }
    }

    [Fact]
    public void Xlsb_OpenModifySave_PreservesChartPart()
    {
        // P0-14(xlsb): 打开含图表的 xlsb → 保存，图表部件/关系/content types 保留
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].SetValue("A1", "数据");
            wb.SaveAs(file);
            InjectChartPart(file);

            var opened = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xlsb, opened.Format);
            opened.Worksheets[0].SetValue("B1", "新增");
            opened.Save();

            // 图表部件保留
            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            Assert.NotNull(zip.GetEntry("xl/charts/chart1.xml"));
            // workbook.bin.rels 合并：chart 关系保留
            var wbRels = ReadEntry(file, "xl/_rels/workbook.bin.rels");
            Assert.Contains("chart", wbRels);
            Assert.Contains("charts/chart1.xml", wbRels);
            // content types 保留 chart Override
            var ct = ReadEntry(file, "[Content_Types].xml");
            Assert.Contains("chart1.xml", ct);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_DocumentProperties_RoundTrip()
    {
        // P0-18(xlsb): xlsb 文档属性读回 + 写出
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Properties.Creator = "批测作者";
            wb.Properties.Title = "批次5标题";
            wb.Properties.Subject = "主题";
            wb.Worksheets[0].SetValue("A1", "x");
            wb.SaveAs(file);

            // 写出 docProps 部件
            using (var zip1 = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read))
                Assert.NotNull(zip1.GetEntry("docProps/core.xml"));
            using (var zip2 = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read))
                Assert.NotNull(zip2.GetEntry("docProps/app.xml"));

            // 读回属性
            var reopened = Excel.Open(file);
            Assert.Equal("批测作者", reopened.Properties.Creator);
            Assert.Equal("批次5标题", reopened.Properties.Title);
            Assert.Equal("主题", reopened.Properties.Subject);

            // 修改属性再保存，仍写回
            reopened.Properties.Creator = "新作者";
            reopened.Save();
            var reopened2 = Excel.Open(file);
            Assert.Equal("新作者", reopened2.Properties.Creator);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void XlsAndXlsb_StyleDegradation_Reported()
    {
        // P0-17: xls/xlsb 完整样式降级为仅数字格式，须显式上报而非静默
        foreach (var (ext, fmt) in new[] { (".xls", ExcelFormat.Xls), (".xlsb", ExcelFormat.Xlsb) })
        {
            var file = GetTempFile(ext);
            var reported = new List<DegradationInfo>();
            try
            {
                var wb = Excel.Create(ExcelFormat.Xlsx);
                var ws = wb.Worksheets[0];
                ws.SetValue("A1", "x");
                ws.Cell("B1").SetValue("y");
                ws.Cell("B1").Style = new CellStyle { Bold = true, FontColor = "#FF0000" };

                Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => reported.Add(d) });

                Assert.Contains(reported, d => d.Capability == DegradationCapability.Styles && d.TargetFormat == fmt);
            }
            finally { if (File.Exists(file)) File.Delete(file); }
        }
    }

    [Fact]
    public void Xlsb_NumberFormatOnly_NoStyleDegradation()
    {
        // 仅数字格式不应触发 Styles 降级（数字格式 xlsb 已支持）
        var file = GetTempFile(".xlsb");
        var reported = new List<DegradationInfo>();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.Cell("B1").SetValue(Cell.FromNumber(3.14, "0.00"));

            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => reported.Add(d) });

            Assert.DoesNotContain(reported, d => d.Capability == DegradationCapability.Styles);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_SheetLevelRels_Preserved()
    {
        // xlsb 工作表级保留 rels（图表挂 sheet）在打开-保存后保留
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].SetValue("A1", "x");
            wb.SaveAs(file);

            using (var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update))
            {
                var chart = zip.CreateEntry("xl/charts/chart1.xml");
                using var s = chart.Open();
                var bytes = Encoding.UTF8.GetBytes("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>");
                s.Write(bytes, 0, bytes.Length);

                var rels = zip.CreateEntry("xl/worksheets/_rels/sheet1.bin.rels");
                using var rs = rels.Open();
                var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" Target=\"../../charts/chart1.xml\"/>" +
                    "</Relationships>";
                var bytes2 = Encoding.UTF8.GetBytes(xml);
                rs.Write(bytes2, 0, bytes2.Length);
            }

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "y");
            opened.Save();

            var sheetRels = ReadEntry(file, "xl/worksheets/_rels/sheet1.bin.rels");
            Assert.Contains("chart", sheetRels);
            Assert.Contains("charts/chart1.xml", sheetRels);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
