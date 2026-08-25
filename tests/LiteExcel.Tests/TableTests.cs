using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 2.5.0 批 2：超级表（Table/ListObject）写出/读回/校验。
/// </summary>
public class TableTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"tbl_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void AddTable_WritesParts()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "姓名"); ws.SetValue("B1", "薪资");
            ws.SetValue("A2", "张三"); ws.SetValue("B2", 5000);
            ws.AddTable("A1:B2", "员工表");

            wb.SaveAs(file);

            using var zip = ZipFile.OpenRead(file);
            Assert.NotNull(zip.GetEntry("xl/tables/table1.xml"));

            string tableXml;
            using (var s = zip.GetEntry("xl/tables/table1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                tableXml = r.ReadToEnd();
            Assert.Contains("name=\"员工表\"", tableXml);
            Assert.Contains("ref=\"A1:B2\"", tableXml);
            Assert.Contains("TableStyleMedium9", tableXml);
            Assert.Contains("showRowStripes=\"1\"", tableXml);

            string sheetXml;
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();
            Assert.Contains("<tableParts count=\"1\">", sheetXml);
            Assert.Contains("<tablePart r:id=\"rIdT1\"/>", sheetXml);

            // rels
            string rels;
            using (var s = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                rels = r.ReadToEnd();
            Assert.Contains("/table", rels);
            Assert.Contains("../tables/table1.xml", rels);

            // ContentTypes
            string ct;
            using (var s = zip.GetEntry("[Content_Types].xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                ct = r.ReadToEnd();
            Assert.Contains("table+xml", ct);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AddTable_ColumnFormat_WritesDxfAndDataDxf()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "日期"); ws.SetValue("B1", "金额");
            ws.SetValue("A2", "2020-01-01"); ws.SetValue("B2", 12.5);
            var tbl = ws.AddTable("A1:B2", "T");
            tbl.Column("日期").NumberFormat = "yyyy/m/d";
            tbl.Column("金额").NumberFormat = "#,##0.00";

            wb.SaveAs(file);

            string tableXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/tables/table1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                tableXml = r.ReadToEnd();
            Assert.Contains("dataDxfId=\"0\"", tableXml);
            Assert.Contains("dataDxfId=\"1\"", tableXml);

            string stylesXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/styles.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                stylesXml = r.ReadToEnd();
            Assert.Contains("yyyy/m/d", stylesXml);
            Assert.Contains("#,##0.00", stylesXml);
            Assert.Contains("<tableStyles", stylesXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AddTable_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "姓名"); ws.SetValue("B1", "薪资");
            ws.SetValue("A2", "张三"); ws.SetValue("B2", 5000);
            ws.SetValue("A3", "李四"); ws.SetValue("B3", 7000);
            ws.AddTable("A1:B3", "员工表", TableStyleStyle.Medium6);

            wb.SaveAs(file);

            var opened = Excel.Open(file);
            var tbl = opened.Worksheets[0].Tables.SingleOrDefault();
            Assert.NotNull(tbl);
            Assert.Equal("员工表", tbl!.Name);
            Assert.Equal("A1:B3", tbl.Ref);
            Assert.Equal(TableStyleStyle.Medium6, tbl.Style);
            Assert.Equal("TableStyleMedium6", tbl.CustomStyleName);
            Assert.True(tbl.ShowRowStripes);
            Assert.True(tbl.AutoFilter);
            Assert.Equal(2, tbl.Columns.Count);
            Assert.Equal("姓名", tbl.Columns[0].Name);
            Assert.Equal("薪资", tbl.Columns[1].Name);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AddTable_CustomStyleName_UnknownGoesThrough()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
            ws.AddTable("A1:A2", "表T", "TableStyleDark3");

            wb.SaveAs(file);
            string tableXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/tables/table1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                tableXml = r.ReadToEnd();
            Assert.Contains("TableStyleDark3", tableXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void UnknownStyleName_ReportsDegradation()
    {
        var file = GetTempFile();
        var degs = new List<DegradationCapability>();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
            ws.AddTable("A1:A2", "表T", "MyCustomStyle99");
            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => degs.Add(d.Capability) });
            Assert.Contains(DegradationCapability.Tables, degs);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void TableOnCsv_ReportsDegradation()
    {
        var file = GetTempFile().Replace(".xlsx", ".csv");
        var degs = new List<DegradationCapability>();
        try
        {
            var wb = Excel.Create("S", ExcelFormat.Csv);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
            ws.AddTable("A1:A2", "表T");
            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => degs.Add(d.Capability) });
            Assert.Contains(DegradationCapability.Tables, degs);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── 校验 ──

    [Fact]
    public void AddTable_InvalidName_Throws()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
        // 以数字开头
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:A2", "1abc"));
        // 含空格
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:A2", "a b"));
        // 单元格地址（T1 = 第 T 列第 1 行，Excel 拒绝）
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:A2", "T1"));
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:A2", "C1"));
    }

    [Fact]
    public void AddTable_TooFewRows_Throws()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "A");
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:A1", "表T"));
    }

    [Fact]
    public void AddTable_Overlap_Throws()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
        ws.AddTable("A1:A2", "表1");
        ws.AddTable("C1:C2", "表2"); // 不重叠可以
        Assert.Throws<LiteExcelException>(() => ws.AddTable("A1:B2", "表3")); // 重叠
    }

    [Fact]
    public void AddTable_DuplicateName_Throws()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
        ws.SetValue("C1", "C"); ws.SetValue("C2", "2");
        ws.AddTable("A1:A2", "表T");
        Assert.Throws<LiteExcelException>(() => ws.AddTable("C1:C2", "表T"));
    }

    [Fact]
    public void RemoveTable_Removes()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "A"); ws.SetValue("A2", "1");
        ws.AddTable("A1:A2", "表T");
        Assert.True(ws.RemoveTable("表T"));
        Assert.False(ws.RemoveTable("表T"));
        Assert.Empty(ws.Tables);
    }

    // ── 保真：打开含表的真实文件再保存，表仍存在 ──

    [Fact]
    public void OpenResave_PreservesTable()
    {
        var src = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-compatibility.xlsx");
        if (!File.Exists(src)) return; // fixture 缺失时跳过

        var file = GetTempFile();
        try
        {
            var opened = Excel.Open(src);
            var hadTables = opened.Worksheets.Any(ws => ws.Tables.Count > 0);
            if (!hadTables) return;
            opened.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.True(reopened.Worksheets.Any(ws => ws.Tables.Count > 0), "另存后表丢失");
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
