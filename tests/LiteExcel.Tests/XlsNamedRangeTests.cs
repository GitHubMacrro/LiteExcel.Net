using LiteExcel;
using System.IO;

namespace LiteExcel.Tests;

/// <summary>
/// xls（BIFF8）命名区域读回：真实 Excel 样本 + 复用 XlsTestFile 构造样本。
/// 范围：仅支持 PtgRef3d / PtgArea3d 的简单单元格/区域引用；复杂公式跳过。
/// </summary>
public class XlsNamedRangeTests
{
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Fact]
    public void Open_ExcelAuthored_GlobalAndLocalNames_ReadBack()
    {
        var wb = Excel.Open(Fixture("excel-authored-namedranges.xls"));

        // 4 个全局名（引用 4 个 sheet）+ 4 个局部名
        Assert.Equal(8, wb.Names.Count);

        var global = wb.Names.Where(n => n.LocalSheetId < 0).ToList();
        var local = wb.Names.Where(n => n.LocalSheetId >= 0).ToList();
        Assert.Equal(4, global.Count);
        Assert.Equal(4, local.Count);

        // 全局名引用文本与 Excel COM RefersTo 一致
        Assert.Contains(wb.Names, n => n.Name == "G_One" && n.Reference == "One!A1");
        Assert.Contains(wb.Names, n => n.Name == "G_Four" && n.Reference == "Four!A1");

        // 局部名：LocalSheetId = sheet 索引，且属于正确 sheet
        Assert.Contains(wb.Names, n => n.Name == "L_One" && n.Reference == "One!B1" && n.LocalSheetId == 0);
        Assert.Contains(wb.Names, n => n.Name == "L_Two" && n.Reference == "Two!B1" && n.LocalSheetId == 1);
        Assert.Contains(wb.Names, n => n.Name == "L_Three" && n.Reference == "Three!B1" && n.LocalSheetId == 2);
        Assert.Contains(wb.Names, n => n.Name == "L_Four" && n.Reference == "Four!B1" && n.LocalSheetId == 3);
    }

    [Fact]
    public void Open_ExcelAuthored_SimpleNames_ReadBack()
    {
        var wb = Excel.Open(Fixture("excel-authored-namedranges-simple.xls"));

        Assert.Contains(wb.Names, n => n.Name == "CrossSheet" && n.Reference == "Sheet2!B2");
        Assert.Contains(wb.Names, n => n.Name == "MyRange" && n.Reference == "Sheet1!A1:C9");
        Assert.Contains(wb.Names, n => n.Name == "LocalOne" && n.Reference == "Sheet1!D1");
    }

    [Fact]
    public void Open_NoNames_Empty()
    {
        // 库生成的无命名区域 xls：Names 应为空
        var tmp = Path.Combine(Path.GetTempPath(), $"xlsnr_{Guid.NewGuid():N}.xls");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            wb.SaveAs(tmp, ExcelFormat.Xls);
            var reopened = Excel.Open(tmp);
            Assert.Empty(reopened.Names);
        }
        finally { if (File.Exists(tmp)) File.Delete(tmp); }
    }
}
