using LiteExcel;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 基础功能测试缺陷修复回归（批次 C）：
/// Bug 2：无数据行的行高此前被静默丢弃（xlsx/xlsb/xls 三格式）。
/// 修复后须补写仅含行高的行记录，且高层读回可见。
/// </summary>
public class BugfixBatchC_EmptyRowHeightTests
{
    private static string TempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_bugfixC2_{Guid.NewGuid():N}{ext}");

    private static string ReadSheetXml(string xlsxPath, string entry = "xl/worksheets/sheet1.xml")
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var e = zip.GetEntry(entry);
        Assert.NotNull(e);
        using var sr = new StreamReader(e!.Open());
        return sr.ReadToEnd();
    }

    [Fact]
    public void Bug2_Xlsx_EmptyRowHeight_IsWrittenAndReadBack()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "r1");                                      // 仅第 1 行有数据
            ws.RowHeights = new Dictionary<int, double> { { 4, 40.0 } };  // 第 5 行（0-based=4）
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("<row r=\"5\" ht=\"40\" customHeight=\"1\"/>", xml);

            // 高层读回（用户视角）：第 5 行行高 = 40
            var opened = Excel.Open(file);
            var rh = opened.Worksheets[0].RowHeights;
            Assert.NotNull(rh);
            Assert.True(rh!.ContainsKey(4), "第 5 行行高应读回");
            Assert.Equal(40.0, rh[4], 0.001);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug2_Xlsx_EmptyRowInsideGrid_Height_IsWritten()
    {
        // 网格内空行（第 2 行无数据）也要保留行高。
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "r1");
            ws.SetValue("A3", "r3");                        // 第 2 行无数据
            ws.RowHeights = new Dictionary<int, double> { { 1, 33.0 } };
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("<row r=\"2\" ht=\"33\" customHeight=\"1\"/>", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug2_Xlsb_EmptyRowHeight_IsWrittenAndReadBack()
    {
        var file = TempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "r1");
            ws.RowHeights = new Dictionary<int, double> { { 4, 40.0 } };
            wb.SaveAs(file);

            var opened = Excel.Open(file, ExcelFormat.Xlsb);
            var rh = opened.Worksheets[0].RowHeights;
            Assert.NotNull(rh);
            Assert.True(rh!.ContainsKey(4), "xlsb 空行行高应写出并可读回");
            Assert.Equal(40.0, rh[4], 0.5);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug2_Xls_EmptyRowHeight_IsWrittenAndReadBack()
    {
        var file = TempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xls);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "r1");
            ws.RowHeights = new Dictionary<int, double> { { 4, 40.0 } };
            wb.SaveAs(file);

            var opened = Excel.Open(file, ExcelFormat.Xls);
            var rh = opened.Worksheets[0].RowHeights;
            Assert.NotNull(rh);
            Assert.True(rh!.ContainsKey(4), "xls 空行行高应写出并可读回");
            Assert.Equal(40.0, rh[4], 0.5);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
