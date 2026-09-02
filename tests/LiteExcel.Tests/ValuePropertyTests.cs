using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Cell.Value 与 ExcelRange.Value 属性：单格标量、多格 2D 数组读写、与 GetString/GetValue 一致。
/// </summary>
public class ValuePropertyTests
{
    [Fact]
    public void Cell_Value_SetAndGet_Scalar()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];

        ws.Cell("A1").Value = "123";
        Assert.Equal("123", ws.Cell("A1").Value);
        Assert.Equal("123", ws.Cell("A1").GetString());

        ws.Cell("A1").Value = 25;
        Assert.Equal(25.0, ws.Cell("A1").Value);
        Assert.Equal(25.0, ws.Cell("A1").GetDouble());

        ws.Cell("A1").Value = 1.5;
        Assert.Equal(1.5, ws.Cell("A1").Value);

        ws.Cell("A1").Value = true;
        Assert.Equal(true, ws.Cell("A1").Value);

        ws.Cell("A1").Value = null;
        Assert.True(ws.Cell("A1").IsEmpty);
        Assert.Null(ws.Cell("A1").Value);
    }

    [Fact]
    public void Cell_Value_MatchesGetValue()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Cell("B2").Value = "hello";
        Assert.Equal(ws.Cell("B2").GetValue(), ws.Cell("B2").Value);
        Assert.Equal(CellType.Text, ws.Cell("B2").Type);
    }

    [Fact]
    public void Cells_Indexer_Value()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Cells["C3"].Value = "via-cells";
        Assert.Equal("via-cells", ws.Cells["C3"].Value);
        ws.Cells[4, 4].Value = 99;
        Assert.Equal(99.0, ws.Cells[4, 4].Value);
    }

    [Fact]
    public void Range_Value_SingleCell_Scalar()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Range("A1").Value = "123";
        Assert.Equal("123", ws.Range("A1").Value);
        // 单格返回标量，不是数组
        Assert.IsNotType<object?[,]>(ws.Range("A1").Value);
    }

    [Fact]
    public void Range_Value_MultiCell_2D()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        // 写 2D
        ws.Range("A1:B2").Value = new object?[,] { { 1, 2 }, { 3, 4 } };
        // 多格读返回 2D
        var values = Assert.IsType<object?[,]>(ws.Range("A1:B2").Value);
        Assert.Equal(2, values.GetLength(0));
        Assert.Equal(2, values.GetLength(1));
        Assert.Equal(1.0, values[0, 0]);
        Assert.Equal(2.0, values[0, 1]);
        Assert.Equal(3.0, values[1, 0]);
        Assert.Equal(4.0, values[1, 1]);
    }

    [Fact]
    public void Range_Value_MultiCell_ScalarFill()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Range("A1:C3").Value = "x";
        var values = (object?[,])ws.Range("A1:C3").Value;
        Assert.Equal("x", values[0, 0]);
        Assert.Equal("x", values[2, 2]);
        // 单格标量
        Assert.Equal("x", ws.Range("B2").Value);
    }

    [Fact]
    public void Range_Value_RoundTripAfterSave()
    {
        var path = Path.Combine(Path.GetTempPath(), $"valprop_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.Range("A1:B2").Value = new object?[,] { { 1, 2 }, { 3, 4 } };
            wb.SaveAs(path);

            var opened = Excel.Open(path);
            var values = (object?[,])opened.Worksheets[0].Range("A1:B2").Value;
            Assert.Equal(1.0, values[0, 0]);
            Assert.Equal(4.0, values[1, 1]);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
