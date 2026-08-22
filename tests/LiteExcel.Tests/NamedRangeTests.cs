using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Workbook.Names：definedNames 读回。
/// </summary>
public class NamedRangeTests
{
    [Fact]
    public void Open_ExcelAuthored_NamesRead_Back()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-comments-filter.xlsx");
        Assert.True(File.Exists(path));

        var wb = Excel.Open(path);
        Assert.NotNull(wb.Names);
        Assert.True(wb.Names.Count >= 1, "expected at least one defined name");
        Assert.Contains(wb.Names, n => n.Name == "Score");
    }

    [Fact]
    public void Names_RoundTrip_PreserveXmlWhenUntouched()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-comments-filter.xlsx");
        Assert.True(File.Exists(path));

        var tmp = Path.Combine(Path.GetTempPath(), $"nr_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Open(path);
            wb.SaveAs(tmp);

            using var zip = System.IO.Compression.ZipFile.OpenRead(tmp);
            using var s = zip.GetEntry("xl/workbook.xml")!.Open();
            using var r = new StreamReader(s);
            var xml = r.ReadToEnd();
            Assert.Contains("definedNames", xml);
            Assert.Contains("Score", xml);
        }
        finally { if (File.Exists(tmp)) File.Delete(tmp); }
    }
}
