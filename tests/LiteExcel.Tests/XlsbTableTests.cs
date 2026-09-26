using LiteExcel;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 T2-A：XLSB 超级表（Table/ListObject）读 + 写。
/// </summary>
public class XlsbTableTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"xlsbtbl_{Guid.NewGuid():N}{ext}");

    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Fact]
    public void Xlsb_Table_WriteThenRead_RoundTrip()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.Name = "Data";
            ws.SetValue("A1", "Name"); ws.SetValue("B1", "Qty");
            ws.SetValue("A2", "a"); ws.SetValue("B2", 1);
            ws.SetValue("A3", "b"); ws.SetValue("B3", 2);
            ws.SetValue("A4", "c"); ws.SetValue("B4", 3);
            ws.AddTable("A1:B4", "MyTable", TableStyleStyle.Medium9);
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            var tables = reopened.Worksheets["Data"].ToSheetData().Tables;
            var t = Assert.Single(tables);
            Assert.Equal("MyTable", t.Name);
            Assert.Equal("A1:B4", t.Ref);
            Assert.Equal(2, t.Columns.Count);
            Assert.Equal("Name", t.Columns[0].Name);
            Assert.Equal("Qty", t.Columns[1].Name);
            Assert.Equal("TableStyleMedium9", t.CustomStyleName);
            Assert.True(t.ShowRowStripes);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_Table_ReadFromRealFixture()
    {
        var path = @"D:\AiStory\Test\Test.xlsb";
        if (!File.Exists(path)) return; // 本机样本可选
        var wb = Excel.Open(path);
        var all = wb.Worksheets.SelectMany(w => w.ToSheetData().Tables).ToList();
        Assert.Equal(12, all.Count);
        Assert.Contains(all, t => t.Name == "Years" && t.Ref == "A1:A1048576");
        Assert.Contains(all, t => t.Name == "BUSP_VW_BELOTLOSSDETAILS" && t.Columns.Count == 17);
    }

    [Fact]
    public void Xlsb_Table_MultipleSheets_GlobalIds()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Name = "S1";
            wb.Worksheets[0].SetValue("A1", "H"); wb.Worksheets[0].SetValue("A2", "v");
            wb.Worksheets[0].AddTable("A1:A2", "TableOne", TableStyleStyle.Light1);
            var s2 = wb.Worksheets.Add("S2");
            s2.SetValue("A1", "H"); s2.SetValue("A2", "v");
            s2.AddTable("A1:A2", "TableTwo", TableStyleStyle.Dark1);
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets["S1"].ToSheetData().Tables);
            Assert.Single(reopened.Worksheets["S2"].ToSheetData().Tables);
            Assert.Equal("TableOne", reopened.Worksheets["S1"].ToSheetData().Tables[0].Name);
            Assert.Equal("TableTwo", reopened.Worksheets["S2"].ToSheetData().Tables[0].Name);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_RealAuthoredTable_SaveAs_NoTablesDegradation()
    {
        // 真实 Excel 产出的含超级表样本：打开-保存不得误报 Tables 降级，
        // 且 xl/tables/table1.bin 必须保留（历史误报已移除）。
        var src = Fixture("excel-authored-table.xlsb");
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Open(src);
            var srcTables = wb.Worksheets.SelectMany(w => w.ToSheetData().Tables).ToList();
            Assert.NotEmpty(srcTables);

            wb.SaveAs(file, ExcelFormat.Xlsb);

            Assert.DoesNotContain(wb.SaveDegradations, d => d.Capability == DegradationCapability.Tables);

            using (var zip = ZipFile.OpenRead(file))
                Assert.NotNull(zip.GetEntry("xl/tables/table1.bin"));

            var reopened = Excel.Open(file);
            var reopenedTables = reopened.Worksheets.SelectMany(w => w.ToSheetData().Tables).ToList();
            Assert.Equal(srcTables.Count, reopenedTables.Count);
            Assert.Equal(srcTables[0].Name, reopenedTables[0].Name);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_Table_Rebuild_NoTablesDegradation()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "Name"); ws.SetValue("A2", "a");
            ws.AddTable("A1:A2", "MyTable", TableStyleStyle.Medium9);

            wb.SaveAs(file);

            Assert.DoesNotContain(wb.SaveDegradations, d => d.Capability == DegradationCapability.Tables);

            using (var zip = ZipFile.OpenRead(file))
                Assert.NotNull(zip.GetEntry("xl/tables/table1.bin"));

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets[0].ToSheetData().Tables);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_Table_AddOnOpenedWorkbook_Persists()
    {
        // 回归：在打开的文件上 AddTable 曾因未标记工作表修改而被手术式路径丢弃，新表必须落盘。
        var baseFile = GetTempFile(".xlsb");
        var outFile = GetTempFile(".xlsb");
        try
        {
            {
                var wb = Excel.Create(ExcelFormat.Xlsb);
                var ws = wb.Worksheets[0];
                ws.Name = "S0";
                ws.SetValue("A1", "h"); ws.SetValue("A2", "a");
                ws.AddTable("A1:A2", "TblA");
                wb.SaveAs(baseFile, ExcelFormat.Xlsb);
            }

            var wb2 = Excel.Open(baseFile);
            var ws2 = wb2.Worksheets[0];
            ws2.SetValue("C1", "H"); ws2.SetValue("C2", "v");
            ws2.AddTable("C1:C2", "TblB");
            wb2.AllowFeatureLossOnSave = true;
            wb2.SaveAs(outFile, ExcelFormat.Xlsb);

            Assert.Equal(2, Excel.Open(outFile).Worksheets[0].ToSheetData().Tables.Count);
        }
        finally
        {
            if (File.Exists(baseFile)) File.Delete(baseFile);
            if (File.Exists(outFile)) File.Delete(outFile);
        }
    }

    [Fact]
    public void Xlsb_Table_AddOnNewSheetOfOpenedWorkbook_NoDuplicateParts()
    {
        // 回归：新表上加超级表曾产生重复的表部件与残留关系；现要求包内无重复条目、两表均可读回。
        var baseFile = GetTempFile(".xlsb");
        var outFile = GetTempFile(".xlsb");
        try
        {
            {
                var wb = Excel.Create(ExcelFormat.Xlsb);
                var s0 = wb.Worksheets[0];
                s0.Name = "S0";
                s0.SetValue("A1", "h"); s0.SetValue("A2", "a");
                s0.AddTable("A1:A2", "TblA");
                var s1 = wb.Worksheets.Add("S1");
                s1.SetValue("A1", "x");
                wb.SaveAs(baseFile, ExcelFormat.Xlsb);
            }

            var wb2 = Excel.Open(baseFile);
            var s2 = wb2.Worksheets.Add("S2");
            s2.SetValue("A1", "H"); s2.SetValue("A2", "v");
            s2.AddTable("A1:A2", "TblC");
            wb2.AllowFeatureLossOnSave = true;
            wb2.SaveAs(outFile, ExcelFormat.Xlsb);

            using (var zip = ZipFile.OpenRead(outFile))
            {
                var names = zip.Entries.Select(e => e.FullName).ToList();
                Assert.Equal(names.Count, names.Distinct().Count());
                Assert.Equal(2, names.Count(n => n.StartsWith("xl/tables/table") && n.EndsWith(".bin")));
            }
            Assert.Single(Excel.Open(outFile).Worksheets["S0"].ToSheetData().Tables);
            Assert.Single(Excel.Open(outFile).Worksheets["S2"].ToSheetData().Tables);
        }
        finally
        {
            if (File.Exists(baseFile)) File.Delete(baseFile);
            if (File.Exists(outFile)) File.Delete(outFile);
        }
    }
}