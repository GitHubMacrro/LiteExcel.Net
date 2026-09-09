using LiteExcel;

namespace LiteExcel.Tests;

public class InsertDeleteTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"insdel_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void InsertRows_ShiftsCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "h1");
        ws.SetValue("A2", "v1");
        ws.SetValue("A3", "v2");

        ws.InsertRows(2, 1);

        Assert.Equal("h1", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("A2").IsEmpty);
        Assert.Equal("v1", ws.Cell("A3").GetString());
        Assert.Equal("v2", ws.Cell("A4").GetString());
    }

    [Fact]
    public void DeleteRows_ShiftsCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "r1");
        ws.SetValue("A2", "r2");
        ws.SetValue("A3", "r3");
        ws.SetValue("A4", "r4");

        ws.DeleteRows(2, 1);

        Assert.Equal("r1", ws.Cell("A1").GetString());
        Assert.Equal("r3", ws.Cell("A2").GetString());
        Assert.Equal("r4", ws.Cell("A3").GetString());
    }

    [Fact]
    public void InsertColumns_ShiftsCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "a");
        ws.SetValue("B1", "b");
        ws.SetValue("C1", "c");

        ws.InsertColumns(2, 1);

        Assert.Equal("a", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("B1").IsEmpty);
        Assert.Equal("b", ws.Cell("C1").GetString());
        Assert.Equal("c", ws.Cell("D1").GetString());
    }

    [Fact]
    public void DeleteColumns_ShiftsCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "c1");
        ws.SetValue("B1", "c2");
        ws.SetValue("C1", "c3");

        ws.DeleteColumns(2, 1);

        Assert.Equal("c1", ws.Cell("A1").GetString());
        Assert.Equal("c3", ws.Cell("B1").GetString());
    }

    [Fact]
    public void InsertRows_ShiftsMergedRanges()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.Merge("A3:B3");
        ws.InsertRows(1, 1);

        var m = Assert.Single(ws.MergedRanges);
        Assert.Equal(3, m.FirstRow);
        Assert.Equal(3, m.LastRow);
    }

    [Fact]
    public void DeleteRows_RemovesFullyOverlappingMerged()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.Merge("A2:A4");
        ws.Merge("A6:A8");
        ws.DeleteRows(2, 3);

        Assert.Single(ws.MergedRanges);
    }

    [Fact]
    public void InsertRows_ShiftsComments()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "x");
        ws.SetValue("A3", "y");
        ws.Comments = new Dictionary<string, string> { ["A3"] = "note" };

        ws.InsertRows(1, 1);

        Assert.True(ws.Comments!.ContainsKey("A4"));
        Assert.False(ws.Comments.ContainsKey("A3"));
        Assert.Equal("note", ws.Comments["A4"]);
    }

    [Fact]
    public void InsertColumns_ShiftsComments()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "x");
        ws.SetValue("C1", "y");
        ws.Comments = new Dictionary<string, string> { ["C1"] = "note" };

        ws.InsertColumns(1, 1);

        Assert.True(ws.Comments!.ContainsKey("D1"));
        Assert.Equal("note", ws.Comments["D1"]);
    }

    [Fact]
    public void InsertRows_ShiftsFilter()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.Filter = new AutoFilter { Range = "A1:B5" };

        ws.InsertRows(2, 1);

        Assert.Equal("A1:B6", ws.Filter!.Range);
    }

    [Fact]
    public void InsertRows_ShiftsRowHeights()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "x");
        ws.SetValue("A3", "y");
        ws.RowHeights = new Dictionary<int, double> { [0] = 20, [2] = 30 };

        ws.InsertRows(1, 1);

        Assert.Equal(20, ws.RowHeights![1]);
        Assert.Equal(30, ws.RowHeights![3]);
    }

    [Fact]
    public void InsertRows_RoundTripXlsx()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "h");
            ws.SetValue("A2", "v1");
            ws.SetValue("A3", "v2");
            ws.InsertRows(2, 1);
            ws.SetValue("B2", "inserted");
            wb.SaveAs(file);

            var read = Excel.Open(file);
            Assert.Equal("h", read.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("v1", read.Worksheets[0].Cell("A3").GetString());
            Assert.Equal("v2", read.Worksheets[0].Cell("A4").GetString());
            Assert.Equal("inserted", read.Worksheets[0].Cell("B2").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void DeleteColumns_RoundTripXlsx()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "a");
            ws.SetValue("B1", "b");
            ws.SetValue("C1", "c");
            ws.DeleteColumns(1, 1);
            wb.SaveAs(file);

            var read = Excel.Open(file);
            Assert.Equal("b", read.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("c", read.Worksheets[0].Cell("B1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void InsertRows_MultipleRows()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.SetValue("A1", "r1");
        ws.SetValue("A2", "r2");

        ws.InsertRows(1, 3);

        Assert.Equal("r1", ws.Cell("A4").GetString());
        Assert.Equal("r2", ws.Cell("A5").GetString());
    }
}
