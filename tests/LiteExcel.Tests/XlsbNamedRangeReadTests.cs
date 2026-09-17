using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 T1-B：XLSB 定义名称读回（BrtDefinedName + rgce 解码）。
/// 仅支持单引用/区域/常量；复合表达式跳过，不产出错误引用。
/// 用真实 Excel 生成样本（Fixtures/excel-authored-namedranges.xlsb）验证。
/// </summary>
public class XlsbNamedRangeReadTests
{
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Fact]
    public void Xlsb_ReadsGlobalAndLocalNames()
    {
        var path = Fixture("excel-authored-namedranges.xlsb");
        Assert.True(File.Exists(path), $"缺少 fixture: {path}");

        var wb = Excel.Open(path);
        var byName = wb.Names.ToDictionary(n => n.Name, n => n);

        // 全局：区域引用 / 常量 / 跨表单格引用（Excel 保存时把重复的单格名合并，样本实际含 4 项）
        Assert.Equal("Alpha!$A$1:$B$5", byName["G_Area"].Reference);
        Assert.Equal("42", byName["G_Const"].Reference);
        Assert.Equal("Beta!$A$1", byName["G_BetaRef"].Reference);

        // sheet 级（localSheetId 0-based）
        Assert.Equal("Beta!$A$1", byName["L_Beta"].Reference);
        Assert.Equal(1, byName["L_Beta"].LocalSheetId);
        Assert.Equal(-1, byName["G_Area"].LocalSheetId);
    }

    [Fact]
    public void Xlsb_ReadsFilterDatabase_FromCommentsFilterFixture()
    {
        // 既有 fixture：含 _FilterDatabase（区域引用）
        var path = Fixture("excel-authored-comments-filter.xlsb");
        Assert.True(File.Exists(path));
        var wb = Excel.Open(path);
        Assert.Contains(wb.Names, n => n.Name == "_FilterDatabase" && n.Reference.Contains("$A$1:$B$5"));
    }

    [Fact]
    public void Xlsb_ComplexExpressions_AreSkippedNotCorrupted()
    {
        // Test.xlsb 含 _xlcn.LinkedTable_* / _xlfn.CUBESET 等复合表达式：
        // 必须跳过（不产出错误引用），已解码项引用合法。
        var path = @"D:\AiStory\Test\Test.xlsb";
        if (!File.Exists(path)) return; // 本机样本可选

        var wb = Excel.Open(path);
        Assert.DoesNotContain(wb.Names, n => n.Name.StartsWith("_xlcn.") || n.Name.StartsWith("_xlfn."));
        foreach (var n in wb.Names)
        {
            Assert.False(string.IsNullOrEmpty(n.Name));
            Assert.False(string.IsNullOrEmpty(n.Reference));
            Assert.DoesNotContain("null", n.Reference);
        }
    }
}
