using LiteExcel;
using System.IO;
using System.IO.Compression;

namespace LiteExcel.Tests;

/// <summary>
/// 真实 Excel 样本对拍（保真验证）。
/// 用 Excel COM 生成的 real-world 样本，验证库读取已有的能力不丢。
/// </summary>
public class ExcelAuthoredFidelityTests
{
    private static string Fixture(string name) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", name);

    [Fact]
    public void Validation_Sample_Reads_Back()
    {
        var path = Fixture("excel-authored-validation.xlsx");
        Assert.True(File.Exists(path), $"missing fixture: {path}");

        var wb = Excel.Open(path);
        var ws = wb.Worksheets[0];
        var vs = ws.Validations;
        Assert.NotNull(vs);
        Assert.NotEmpty(vs);

        var dv = vs.FirstOrDefault(v => v.Sqref == "A1:A8");
        Assert.NotNull(dv);
        Assert.Equal(DataValidationType.WholeNumber, dv.Type);

        var dvList = vs.FirstOrDefault(v => v.Sqref == "B1:B8");
        Assert.NotNull(dvList);
        Assert.Equal(DataValidationType.List, dvList.Type);
        Assert.Contains("IT", dvList.Formula1);
    }

    [Fact]
    public void Comments_And_AutoFilter_Sample_Reads_Back()
    {
        var path = Fixture("excel-authored-comments-filter.xlsx");
        Assert.True(File.Exists(path), $"missing fixture: {path}");

        var wb = Excel.Open(path);
        var ws = wb.Worksheets[0];

        // 批注：B2 / B4
        Assert.True(ws.Comments is { Count: > 0 });
        Assert.True(ws.Comments.ContainsKey("B2") || ws.Comments.ContainsKey("B4"));

        // AutoFilter：A1:B6
        Assert.NotNull(ws.Filter);
        Assert.False(string.IsNullOrEmpty(ws.Filter.Range));
        Assert.Contains("A1", ws.Filter.Range);
    }

    [Fact]
    public void NamedRange_Sample_RecordedInWorkbookXml()
    {
        var path = Fixture("excel-authored-comments-filter.xlsx");
        var file = Path.Combine(Path.GetTempPath(), $"named_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Open(path);
            wb.SaveAs(file);

            using var zip = ZipFile.OpenRead(file);
            using var s = zip.GetEntry("xl/workbook.xml")!.Open();
            string xml;
            using (var r = new StreamReader(s)) xml = r.ReadToEnd();
            Assert.Contains("definedNames", xml);
            Assert.Contains("Score", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
