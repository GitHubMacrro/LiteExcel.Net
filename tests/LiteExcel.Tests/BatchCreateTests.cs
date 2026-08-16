using LiteExcel;

namespace LiteExcel.Tests;

public class BatchCreateTests
{
    [Fact]
    public void BatchCreate_CreatesSheetsInOrder()
    {
        var wb = Excel.Create(new[] { "一月", "二月", "三月" });

        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal(new[] { "一月", "二月", "三月" }, wb.Worksheets.Names);
    }

    [Fact]
    public void BatchCreate_NullKeepsDefaultSheet1()
    {
        var wb = Excel.Create((string[]?)null);

        Assert.Single(wb.Worksheets);
        Assert.Equal("Sheet1", wb.Worksheets.Names[0]);
    }

    [Fact]
    public void BatchCreate_EmptyArrayKeepsDefaultSheet1()
    {
        var wb = Excel.Create(Array.Empty<string>());

        Assert.Single(wb.Worksheets);
        Assert.Equal("Sheet1", wb.Worksheets.Names[0]);
    }

    [Fact]
    public void BatchCreate_DuplicateNamesThrow()
    {
        Assert.Throws<LiteExcelException>(() => Excel.Create(new[] { "表A", "表A" }));
    }

    [Fact]
    public void BatchCreate_BlankNameThrows()
    {
        Assert.Throws<ArgumentException>(() => Excel.Create(new[] { "表A", " " }));
    }

    [Fact]
    public void BatchCreate_WritesAndReadsBackAllSheets()
    {
        var file = Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Create(new[] { "第一张", "第二张" });
            wb.Worksheets[0].SetValue("A1", "a");
            wb.Worksheets[1].SetValue("A1", "b");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal(new[] { "第一张", "第二张" }, opened.Worksheets.Names);
            Assert.Equal("a", opened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("b", opened.Worksheets[1].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
