using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 公式文本解析测试：BIFF8（xls）与 BIFF12（xlsb）的 RPN → A1 文本。
/// 使用真实 Excel 生成的 fixture（含常见公式）与既有真实文件 fixture（含 =B2*2）。
/// </summary>
public class FormulaTests
{
    private static string GetFixturePath(string name)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", name);
        Assert.True(File.Exists(path), $"Required fixture is missing: {path}");
        return path;
    }

    [Fact]
    public void Xls_ReadsFormulas_FromExcelAuthored()
    {
        var s = Excel.Open(GetFixturePath("excel-formulas.xls")).Worksheets[0];

        AssertFormula(s, 1, 3, "B2*2", 0.0);
        AssertFormula(s, 1, 4, "SUM(A1:B1)", 30.0);
        AssertFormula(s, 1, 5, "IF(A1>5,1,0)", 1.0);
        AssertFormula(s, 1, 6, "A1+B1", 30.0);
        AssertFormula(s, 1, 7, "ROUND(A1,1)", 10.0);
        AssertFormula(s, 1, 9, "A1&\"x\"");
        AssertFormula(s, 1, 10, "1/3");
        AssertFormula(s, 1, 11, "MAX(A1:B1)", 20.0);
        AssertFormula(s, 1, 12, "ABS(-5)", 5.0);
    }

    [Fact]
    public void Xls_ReadsFormula_FromRealFixture()
    {
        var s = Excel.Open(GetFixturePath("excel-authored.xls")).Worksheets[0];
        var cell = s.Cell("E2");
        Assert.True(cell.IsFormula);
        // P0-8: 公式串在 Formula，缓存值在 Number/Text
        Assert.Equal("B2*2", cell.Formula);
        Assert.Equal(50.0, cell.GetDouble());
    }

    [Fact]
    public void Xlsb_ReadsFormula_FromRealFixture()
    {
        var s = Excel.Open(GetFixturePath("excel-authored.xlsb")).Worksheets[0];
        var cell = s.Cell("E2");
        Assert.True(cell.IsFormula);
        Assert.Equal("B2*2", cell.Formula);
        Assert.Equal(50.0, cell.GetDouble());
    }

    private static void AssertFormula(Worksheet s, int row, int col, string expected, double? cached = null)
    {
        var cell = s.Cell(row, col);
        Assert.True(cell.IsFormula, $"R{row}C{col} 应为公式");
        Assert.Equal(expected, cell.Formula);
        if (cached is { } v)
            Assert.Equal(v, cell.GetDouble(), 9);
    }
}
