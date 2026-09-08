using LiteExcel;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 保真写入回归：verbatim 保留（XLSX/XLSB）与 XLS 透视表有损保存保护。
/// </summary>
public class FidelityWriteVerbatimTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_fwv_{Guid.NewGuid():N}{ext}");

    private static string ReadText(ZipArchive zip, string entry)
    {
        var e = zip.GetEntry(entry);
        if (e is null) return "";
        using var s = e.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }

    private static byte[] ReadBytes(ZipArchive zip, string entry)
    {
        var e = zip.GetEntry(entry);
        if (e is null) return Array.Empty<byte>();
        using var s = e.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private static string FixturePath(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    private static string TakeFixture(string name, string dstExt)
    {
        var loaded = FixturePath(name);
        Assert.True(File.Exists(loaded), $"缺失 fixture：{loaded}");
        var dst = GetTempFile(dstExt);
        File.Copy(loaded, dst, true);
        return dst;
    }

    // ── XLSX/XLSM verbatim：扩展样式保留 ──

    [Fact]
    public void Xlsx_OpenSaveExendedStyles_KeepsSlicerTimelinePivotStyles()
    {
        var src = TakeFixture("excel-authored-pivot.xlsx", ".xlsx");
        var dst = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            wb.SaveAs(dst);

            using var zip = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var styles = ReadText(zip, "xl/styles.xml");
            // verbatim 生效：扩展样式原样保留（重建会丢）
            Assert.Contains("<extLst>", styles);
            Assert.Contains("slicerStyles", styles);
            Assert.Contains("timelineStyles", styles);
            Assert.Contains("pivotButton", styles);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }

    [Fact]
    public void Xlsx_OpenSave_MarkCellModified_VerbatimDisabled_RebuildsStyles()
    {
        var src = TakeFixture("excel-authored-pivot.xlsx", ".xlsx");
        var dst = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Open(src);
            // 修改单元格 → 禁用 verbatim → 重建 styles.xml（扩展样式会丢）
            wb.Worksheets[0].SetValue("B2", "changed");
            wb.SaveAs(dst);

            using var zip = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var styles = ReadText(zip, "xl/styles.xml");
            Assert.DoesNotContain("slicerStyles", styles);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }

    [Fact]
    public void Xlsx_OpenSave_NoExtendedStyles_UsesRebuildPath_WhenStylesSimple()
    {
        // 简单文件（新建生成的 styles.xml 无扩展）不放行 verbatim，走重建路径
        var src = GetTempFile(".xlsx");
        var dst = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "v1");
            wb.SaveAs(src);

            var opened = Excel.Open(src);
            opened.SaveAs(dst);

            using var zip = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var styles = ReadText(zip, "xl/styles.xml");
            Assert.DoesNotContain("<extLst>", styles);
            Assert.DoesNotContain("slicerStyles", styles);
            // 重建路径读回数据应一致
            var check = Excel.Open(dst);
            Assert.Equal("v1", check.Worksheets[0].Cells[1, 1].Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }

    // ── XLSB verbatim：二进制原样保留 ──

    [Fact]
    public void Xlsb_OpenSaveUnmodified_KeyBinariesByteIdentical()
    {
        var src = TakeFixture("excel-authored.xlsb", ".xlsb");
        var dst = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Open(src);
            wb.SaveAs(dst);

            using var z1 = new ZipArchive(File.OpenRead(src), ZipArchiveMode.Read);
            using var z2 = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            foreach (var part in new[] { "xl/workbook.bin", "xl/styles.bin", "xl/worksheets/sheet1.bin" })
            {
                var b1 = ReadBytes(z1, part);
                var b2 = ReadBytes(z2, part);
                Assert.True(b1.Length > 0, $"{part} 原始缺失");
                Assert.Equal(b1.Length, b2.Length);
                Assert.Equal(b1, b2);
            }
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }

    [Fact]
    public void Xlsb_OpenSave_MarkCellModified_VerbatimDisabled_RebuildsSheet()
    {
        var src = TakeFixture("excel-authored.xlsb", ".xlsb");
        var dst = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Open(src);
            wb.Worksheets[0].SetValue("A1", "modified");
            wb.SaveAs(dst);

            using var z1 = new ZipArchive(File.OpenRead(src), ZipArchiveMode.Read);
            using var z2 = new ZipArchive(File.OpenRead(dst), ZipArchiveMode.Read);
            var b1 = ReadBytes(z1, "xl/worksheets/sheet1.bin");
            var b2 = ReadBytes(z2, "xl/worksheets/sheet1.bin");
            // 修改后重建，sheet1.bin 应与原始不同（值已变）
            Assert.NotEqual(b1, b2);
            // 数据读回应含修改
            var check = Excel.Open(dst);
            Assert.Equal("modified", check.Worksheets[0].Cells[1, 1].Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }

    [Fact]
    public void Xlsb_AdvancedParts_BlockOnlyModifiedHostSheet()
    {
        var safePath = GetTempFile(".xlsb");
        var blockedPath = GetTempFile(".xlsb");
        try
        {
            var safe = Excel.Create(ExcelFormat.Xlsb);
            safe.Worksheets.Add("Other");
            safe.SourceHasAdvancedXlsbParts = true;
            safe.AdvancedXlsbSheetIndexes.Add(0);
            safe.Worksheets[1].SetValue("A1", "safe");
            safe.SaveAs(safePath);
            Assert.True(File.Exists(safePath));

            var blocked = Excel.Create(ExcelFormat.Xlsb);
            blocked.Worksheets.Add("Other");
            blocked.SourceHasAdvancedXlsbParts = true;
            blocked.AdvancedXlsbSheetIndexes.Add(0);
            blocked.Worksheets[0].SetValue("A1", "unsafe");
            Assert.Throws<LiteExcelException>(() => blocked.SaveAs(blockedPath));
            Assert.False(File.Exists(blockedPath));
        }
        finally
        {
            if (File.Exists(safePath)) File.Delete(safePath);
            if (File.Exists(blockedPath)) File.Delete(blockedPath);
        }
    }

    // ── XLS 透视表有损保存保护 ──

    [Fact]
    public void Xls_SourceWithPivot_DefaultSaveThrows_NoFileCreated()
    {
        var src = GetTempFile(".xls");
        var bad = GetTempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xls);
            wb.Worksheets[0].SetValue("A1", "v1");
            wb.SaveAs(src);

            var opened = Excel.Open(src);
            opened.SourceHasPivotTables = true; // 模拟 BIFF8 SXVIEW 检测命中

            Assert.Throws<LiteExcelException>(() => opened.SaveAs(bad));
            // 目标文件未被创建/截断
            Assert.False(File.Exists(bad));
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(bad)) File.Delete(bad);
        }
    }

    [Fact]
    public void Xls_SourceWithPivot_AllowLossSaveSucceeds()
    {
        var src = GetTempFile(".xls");
        var dst = GetTempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xls);
            wb.Worksheets[0].SetValue("A1", "v1");
            wb.SaveAs(src);

            var opened = Excel.Open(src);
            opened.SourceHasPivotTables = true;
            opened.AllowFeatureLossOnSave = true;
            opened.SaveAs(dst);

            Assert.True(File.Exists(dst));
            var check = Excel.Open(dst);
            Assert.Equal("v1", check.Worksheets[0].Cells[1, 1].Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(dst)) File.Delete(dst);
        }
    }
}
