using LiteExcel;

namespace LiteExcel.Tests;

public class DeleteSheetTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"deletesheet_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void Delete_ByName_RemovesSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("Keep");
        wb.Worksheets.Add("Remove");

        Assert.Equal(3, wb.Worksheets.Count);
        wb.Worksheets.Remove("Remove");
        Assert.Equal(2, wb.Worksheets.Count);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        Assert.Equal("Keep", wb.Worksheets[1].Name);
    }

    [Fact]
    public void Delete_ByIndex_RemovesSheet()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("B");
        wb.Worksheets.Add("C");

        Assert.Equal(4, wb.Worksheets.Count);
        wb.Worksheets.RemoveAt(2);
        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        Assert.Equal("A", wb.Worksheets[1].Name);
        Assert.Equal("C", wb.Worksheets[2].Name);
    }

    [Fact]
    public void Delete_ViaWorksheetDelete_RemovesSheet()
    {
        var wb = Excel.Create();
        var target = wb.Worksheets.Add("ToDelete");

        Assert.Equal(2, wb.Worksheets.Count);
        Assert.True(target.Delete());
        Assert.Single(wb.Worksheets);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
    }

    [Fact]
    public void Delete_LastSheet_Throws()
    {
        var wb = Excel.Create();
        Assert.Single(wb.Worksheets);
        Assert.Throws<LiteExcelException>(() => wb.Worksheets[0].Delete());
    }

    [Fact]
    public void Delete_NonExistent_ReturnsFalse()
    {
        var wb = Excel.Create();
        Assert.False(wb.Worksheets.Remove("NoExist"));
        Assert.Single(wb.Worksheets);
    }

    [Fact]
    public void Delete_SaveAs_PreservesRemains()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "Keep");
            wb.Worksheets.Add("Drop").SetValue("A1", "x");
            wb.Worksheets[1].Delete();
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets);
            Assert.Equal("Keep", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_SheetWithComments_Cleanup()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws1 = wb.Worksheets[0];
            ws1.SetValue("A1", "keep");
            var ws2 = wb.Worksheets.Add("WithComments");
            ws2.SetValue("A1", "to remove");
            ws2.Comments = new Dictionary<string, string> { { "A1", "note" } };

            ws2.Delete();
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Single(reopened.Worksheets);
            Assert.Equal("keep", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Delete_Idempotent_ReturnsFalseOnSecondCall()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets.Add("X");

        Assert.True(ws.Delete());
        Assert.False(ws.Delete());       // 已删除的意识
        Assert.Single(wb.Worksheets);
    }

    [Fact]
    public void Delete_ForeachWithListSnapshot_Works()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("A");
        wb.Worksheets.Add("XSheet_1");
        wb.Worksheets.Add("XSheet_2");
        wb.Worksheets.Add("Keep");

        foreach (var ws in wb.Worksheets.Where(w => w.Name.Contains("XSheet")).ToList())
            ws.Delete();

        Assert.Equal(3, wb.Worksheets.Count);
        Assert.Equal("A", wb.Worksheets[1].Name);
        Assert.Equal("Keep", wb.Worksheets[2].Name);
    }
}
