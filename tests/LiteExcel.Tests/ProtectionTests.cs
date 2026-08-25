using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 2.5.0 批 1：工作表/工作簿保护（sheetProtection / workbookProtection）往返与降级。
/// </summary>
public class ProtectionTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"prot_{Guid.NewGuid():N}.xlsx");

    // ── SheetProtection 写出 ──

    [Fact]
    public void SheetProtection_WritesXml()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.Protection = new SheetProtection
            {
                Enabled = true,
                Sort = false,
                AutoFilter = false,
                SelectLockedCells = true,
            };
            ws.Protection.SetPassword("secret");
            wb.SaveAs(file);

            string sheetXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();

            Assert.Contains("<sheetProtection", sheetXml);
            Assert.Contains("algorithmName=\"SHA-512\"", sheetXml);
            Assert.Contains("hashValue=\"", sheetXml);
            Assert.Contains("saltValue=\"", sheetXml);
            Assert.Contains("spinCount=\"100000\"", sheetXml);
            Assert.Contains("sort=\"0\"", sheetXml);
            Assert.Contains("selectLockedCells=\"1\"", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SheetProtection_NoProtection_WritesNothing()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.SaveAs(file);

            string sheetXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                sheetXml = r.ReadToEnd();

            Assert.DoesNotContain("sheetProtection", sheetXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── SheetProtection 读回 ──

    [Fact]
    public void SheetProtection_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets[0];
            ws.Protection = new SheetProtection
            {
                Enabled = true,
                SelectLockedCells = true,
                SelectUnlockedCells = false,
                Sort = false,
                AutoFilter = false,
            };
            ws.Protection.SetPassword("pw123");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            var prot = opened.Worksheets[0].Protection;
            Assert.NotNull(prot);
            Assert.True(prot!.Enabled);
            Assert.True(prot.SelectLockedCells);
            Assert.False(prot.SelectUnlockedCells);
            Assert.False(prot.Sort);
            Assert.True(prot.VerifyPassword("pw123"));
            Assert.False(prot.VerifyPassword("wrong"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── WorkbookProtection 写出/读回 ──

    [Fact]
    public void WorkbookProtection_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.Protection = new WorkbookProtection
            {
                Enabled = true,
                LockStructure = true,
                LockWindows = false,
            };
            wb.Protection.SetPassword("wb123");
            wb.SaveAs(file);

            string wbXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/workbook.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                wbXml = r.ReadToEnd();
            Assert.Contains("<workbookProtection", wbXml);
            Assert.Contains("lockStructure=\"1\"", wbXml);
            Assert.Contains("lockWindows=\"0\"", wbXml);

            var opened = Excel.Open(file);
            Assert.NotNull(opened.Protection);
            Assert.True(opened.Protection!.LockStructure);
            Assert.False(opened.Protection.LockWindows);
            Assert.True(opened.Protection.VerifyPassword("wb123"));
            Assert.False(opened.Protection.VerifyPassword("x"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void WorkbookProtection_NoProtection_WritesNothing()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.SaveAs(file);

            string wbXml;
            using (var zip = ZipFile.OpenRead(file))
            using (var s = zip.GetEntry("xl/workbook.xml")!.Open())
            using (var r = new StreamReader(s, Encoding.UTF8))
                wbXml = r.ReadToEnd();
            Assert.DoesNotContain("workbookProtection", wbXml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── 与文件级密码（打开/修改）共存 ──

    [Fact]
    public void SheetProtection_CoexistsWithModifyPassword()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.Worksheets[0].Protection = new SheetProtection { Enabled = true };
            wb.Security.SetModifyPassword("modify");
            wb.SaveAs(file);

            var opened = Excel.Open(file, new ExcelReadOptions { ModifyPassword = "modify" });
            Assert.NotNull(opened.Worksheets[0].Protection);
            Assert.True(opened.Security.HasModifyPassword);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── 保存后原保护保留（透传）──

    [Fact]
    public void Protection_PreservedAfterReopenResave()
    {
        var file1 = GetTempFile();
        var file2 = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.Worksheets[0].Protection = new SheetProtection { Enabled = true, Sort = false };
            wb.Protection = new WorkbookProtection { Enabled = true, LockStructure = true };
            wb.SaveAs(file1);

            var opened = Excel.Open(file1);
            opened.SaveAs(file2);

            var reopened = Excel.Open(file2);
            Assert.NotNull(reopened.Worksheets[0].Protection);
            Assert.False(reopened.Worksheets[0].Protection!.Sort);
            Assert.NotNull(reopened.Protection);
            Assert.True(reopened.Protection!.LockStructure);
        }
        finally
        {
            if (File.Exists(file1)) File.Delete(file1);
            if (File.Exists(file2)) File.Delete(file2);
        }
    }
}
