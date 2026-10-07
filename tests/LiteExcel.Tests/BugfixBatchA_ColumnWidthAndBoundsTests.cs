using LiteExcel;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 基础功能测试缺陷修复回归（批次 A）：
/// - Bug 1：xlsx 稀疏列宽不再把未设置的前置列写成 width="0"（隐藏列）。
/// - Bug 5：Cell(row,col) 越界（行 &gt; 1048576 / 列 &gt; 16384）抛 ArgumentOutOfRangeException。
/// 断言以「解包 XML / 直接调用」为准，不依赖库自读。
/// </summary>
public class BugfixBatchA_ColumnWidthAndBoundsTests
{
    private static string TempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_bugfixA_{Guid.NewGuid():N}{ext}");

    private static string ReadSheetXml(string xlsxPath, string entry = "xl/worksheets/sheet1.xml")
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var e = zip.GetEntry(entry);
        Assert.NotNull(e);
        using var sr = new StreamReader(e!.Open());
        return sr.ReadToEnd();
    }

    // ── Bug 1：稀疏列宽 ──

    [Fact]
    public void Bug1_SparseColumnWidth_DoesNotWriteZeroWidthColumns()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "a");
            ws.SetValue("B1", "bb");
            ws.SetValue("C1", "ccc");
            ws.ColumnWidths = new Dictionary<int, double> { { 2, 25.0 } }; // 只设 C 列（0-based）
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            // 修复前：会写出 A/B 两条 width="0"（Excel 语义=隐藏列）。
            Assert.DoesNotContain("width=\"0\"", xml);
            // 只应有一条列宽定义，且指向第 3 列。
            var colMatches = System.Text.RegularExpressions.Regex.Matches(xml, "<col [^>]*/>");
            Assert.Single(colMatches);
            Assert.Contains("min=\"3\"", colMatches[0].Value);
            Assert.Contains("customWidth=\"1\"", colMatches[0].Value);

            // 读回：C 列宽度保留，A/B 无自定义宽度。
            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.ColumnWidths);
            Assert.Equal(25.0, read.ColumnWidths![2], 2);
            Assert.True(read.ColumnWidths[0] <= 0);
            Assert.True(read.ColumnWidths[1] <= 0);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Bug1_NoHoleWidths_StillWritesAllCustomColumns()
    {
        var file = TempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            wb.Worksheets[0].ColumnWidths = new Dictionary<int, double> { { 0, 20.0 }, { 1, 10.5 } };
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            var colMatches = System.Text.RegularExpressions.Regex.Matches(xml, "<col [^>]*/>");
            Assert.Equal(2, colMatches.Count);
            Assert.DoesNotContain("width=\"0\"", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── Bug 5：越界地址 ──

    [Theory]
    [InlineData(2_000_000, 3)]   // 行越界
    [InlineData(5, 20_000)]      // 列越界
    [InlineData(1_048_577, 1)]   // 行 = 上限+1
    [InlineData(1, 16_385)]      // 列 = 上限+1
    [InlineData(0, 1)]           // 行下限
    [InlineData(1, 0)]           // 列下限
    public void Bug5_Cell_OutOfRange_Throws(int row, int col)
    {
        var wb = Excel.Create();
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.Worksheets[0].Cell(row, col));
    }

    [Fact]
    public void Bug5_CellByAddress_OutOfRange_Throws()
    {
        var wb = Excel.Create();
        // "ZZZ5" → 列 18278 > 16384
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.Worksheets[0].Cell("ZZZ5"));
        // "A1048577" → 行越界
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.Worksheets[0].Cell("A1048577"));
    }

    [Fact]
    public void Bug5_SetValue_OutOfRange_Throws()
    {
        var wb = Excel.Create();
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.Worksheets[0].SetValue(2_000_000, 1, "x"));
        Assert.Throws<ArgumentOutOfRangeException>(() => wb.Worksheets[0].SetValue(1, 20_000, "y"));
    }

    [Fact]
    public void Bug5_BoundaryCell_IsAcceptedAndWritable()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        // 恰好在上限上：不应抛异常（Cell 返回游离单元格，不物化整张网格）。
        var lastRowCell = ws.Cell(1_048_576, 1);
        Assert.NotNull(lastRowCell);
        var lastColCell = ws.Cell(1, 16_384);
        Assert.NotNull(lastColCell);

        // 写出最后一个合法列（轻量：仅展开一行），验证产物 XML 合法。
        var file = TempFile(".xlsx");
        try
        {
            ws.Cell(1, 16_384).SetValue("lastCol");
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("r=\"XFD1\"", xml);

            var reopened = Excel.Open(file);
            Assert.NotNull(reopened.Worksheets[0].Cell("XFD1"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
