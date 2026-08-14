using LiteExcel;

namespace LiteExcel.Tests;

public class CellRefTests
{
    [Theory]
    [InlineData("A1", 0, 0)]
    [InlineData("B3", 2, 1)]
    [InlineData("Z1", 0, 25)]
    [InlineData("AA1", 0, 26)]
    [InlineData("AB10", 9, 27)]
    [InlineData("AZ1", 0, 51)]
    [InlineData("BA1", 0, 52)]
    [InlineData("ZZ1", 0, 701)]
    [InlineData("AAA1", 0, 702)]
    public void Parse_ValidRefs(string refStr, int expectedRow, int expectedCol)
    {
        var (row, col) = CellRef.Parse(refStr);
        Assert.Equal(expectedRow, row);
        Assert.Equal(expectedCol, col);
    }

    [Theory]
    [InlineData(0, 0, "A1")]
    [InlineData(2, 1, "B3")]
    [InlineData(0, 25, "Z1")]
    [InlineData(0, 26, "AA1")]
    [InlineData(9, 27, "AB10")]
    [InlineData(0, 51, "AZ1")]
    [InlineData(0, 52, "BA1")]
    [InlineData(0, 701, "ZZ1")]
    [InlineData(0, 702, "AAA1")]
    public void ToString_ValidRefs(int row, int col, string expected)
    {
        Assert.Equal(expected, CellRef.ToString(row, col));
    }

    [Theory]
    [InlineData(0, "A")]
    [InlineData(25, "Z")]
    [InlineData(26, "AA")]
    [InlineData(27, "AB")]
    [InlineData(51, "AZ")]
    [InlineData(52, "BA")]
    [InlineData(701, "ZZ")]
    [InlineData(702, "AAA")]
    public void ColToLetter(int col, string expected)
    {
        Assert.Equal(expected, CellRef.ColToLetter(col));
    }

    [Theory]
    [InlineData("A", 0)]
    [InlineData("Z", 25)]
    [InlineData("AA", 26)]
    [InlineData("ZZ", 701)]
    [InlineData("AAA", 702)]
    public void LetterToCol(string letters, int expected)
    {
        Assert.Equal(expected, CellRef.LetterToCol(letters));
    }

    [Fact]
    public void Parse_Lowercase()
    {
        var (row, col) = CellRef.Parse("b3");
        Assert.Equal(2, row);
        Assert.Equal(1, col);
    }
}
