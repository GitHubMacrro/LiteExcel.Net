using LiteExcel;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 基础功能测试缺陷修复回归（批次 C）：
/// Bug 3：对 ws.Cell(...) 取得的（游离）单元格直接赋值 Formula / IsFormula（或 Text 等值属性）
///        此前不落盘（setter 不触发 OnCellChanged）。修复后须正确写出。
/// 断言以「解包 XML / 高层读回」为准，不依赖库自读的低层索引。
/// </summary>
public class BugfixBatchC_FormulaPropertyWriteTests
{
    private static string TempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_bugfixC_{Guid.NewGuid():N}{ext}");

    private static string ReadSheetXml(string xlsxPath, string entry = "xl/worksheets/sheet1.xml")
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var e = zip.GetEntry(entry);
        Assert.NotNull(e);
        using var sr = new StreamReader(e!.Open());
        return sr.ReadToEnd();
    }

    [Fact]
    public void Bug3_BareFormulaProperty_IsWritten()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 3);
            ws.SetValue("A2", 7);

            // 文档 §5.7 原文写法：只设 Formula + IsFormula，不写缓存值。
            var cell = ws.Cell("A3");
            cell.Formula = "SUM(A1:A2)";
            cell.IsFormula = true;

            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("r=\"A3\"", xml);
            Assert.Contains("<f>SUM(A1:A2)</f>", xml);

            // 高层读回（用户视角）
            var opened = Excel.Open(file);
            var a3 = opened.Worksheets[0].Cell("A3");
            Assert.True(a3.IsFormula);
            Assert.Equal("SUM(A1:A2)", a3.Formula);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug3_BareTextProperty_IsWritten()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "h");
            var c = ws.Cell("B1");
            c.Text = "hello";
            c.Type = CellType.Text;
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal("hello", opened.Worksheets[0].Cell("B1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug3_BareNumberProperty_IsWritten()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "h");
            var c = ws.Cell("B1");
            c.Type = CellType.Number;
            c.Number = 42.5;
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal(42.5, opened.Worksheets[0].Cell("B1").GetDouble(), 9);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug3_ModifyingExistingCellFormula_MarksModified()
    {
        // 已入网格的单元格改 Formula 也应落盘（回归 setter 回调）。
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", 1);
            ws.SetValue("A2", 2);
            ws.SetValue("A3", 0);
            ws.Cell("A3").Formula = "A1+A2";
            ws.Cell("A3").IsFormula = true;
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal("A1+A2", opened.Worksheets[0].Cell("A3").Formula);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
