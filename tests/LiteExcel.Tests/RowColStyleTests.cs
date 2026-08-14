using LiteExcel;

namespace LiteExcel.Tests;

public class RowColStyleTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void DefaultStyle_AppliesToAllCells()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                    new Cell[] { Cell.FromText("z"), Cell.FromText("w") },
                },
                DefaultStyle = new CellStyle { FontName = "Arial", FontSize = 12 },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(2, read.Rows.Count);
            Assert.NotNull(read.Rows[0][0].Style);
            Assert.Equal("Arial", read.Rows[0][0].Style!.FontName);
            Assert.Equal(12, read.Rows[0][0].Style.FontSize);
            Assert.Equal("Arial", read.Rows[1][1].Style?.FontName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowStyle_AppliesToWholeRow()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("r0c0"), Cell.FromText("r0c1"), Cell.FromText("r0c2") },
                    new Cell[] { Cell.FromText("r1c0"), Cell.FromText("r1c1"), Cell.FromText("r1c2") },
                },
                RowStyles = new()
                {
                    { 1, new CellStyle { FillColor = "#FF0000", Bold = true } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Row 0 should have no style
            Assert.Null(read.Rows[0][0].Style);
            // Row 1 should have red fill + bold
            Assert.Equal("#FF0000", read.Rows[1][0].Style?.FillColor);
            Assert.True(read.Rows[1][0].Style?.Bold);
            Assert.Equal("#FF0000", read.Rows[1][1].Style?.FillColor);
            Assert.True(read.Rows[1][2].Style?.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ColumnStyle_AppliesToWholeColumn()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                    new Cell[] { Cell.FromText("z"), Cell.FromText("w") },
                },
                ColumnStyles = new()
                {
                    { 1, new CellStyle { Italic = true, FontColor = "#0000FF" } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Column 0: no style
            Assert.Null(read.Rows[0][0].Style);
            Assert.Null(read.Rows[1][0].Style);
            // Column 1: italic + blue
            Assert.True(read.Rows[0][1].Style?.Italic);
            Assert.Equal("#0000FF", read.Rows[0][1].Style?.FontColor);
            Assert.True(read.Rows[1][1].Style?.Italic);
            Assert.Equal("#0000FF", read.Rows[1][1].Style?.FontColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void CellStyle_OverridesRowStyle()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new() { Type = CellType.Text, Text = "cell-style", Style = new CellStyle { Bold = true } },
                        Cell.FromText("row-style"),
                    },
                },
                RowStyles = new()
                {
                    { 0, new CellStyle { FillColor = "#00FF00" } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Cell 0: has its own style (bold), should NOT get row's green fill
            Assert.True(read.Rows[0][0].Style?.Bold);
            Assert.Null(read.Rows[0][0].Style?.FillColor);
            // Cell 1: no own style, should get row's green fill
            Assert.Equal("#00FF00", read.Rows[0][1].Style?.FillColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowStyle_OverridesColumnStyle()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                },
                RowStyles = new()
                {
                    { 0, new CellStyle { FillColor = "#FF0000" } },
                },
                ColumnStyles = new()
                {
                    { 0, new CellStyle { Italic = true } },
                    { 1, new CellStyle { Italic = true } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Row style wins over column style
            Assert.Equal("#FF0000", read.Rows[0][0].Style?.FillColor);
            Assert.False(read.Rows[0][0].Style?.Italic);
            Assert.Equal("#FF0000", read.Rows[0][1].Style?.FillColor);
            Assert.False(read.Rows[0][1].Style?.Italic);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ColumnStyle_OverridesDefaultStyle()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                },
                DefaultStyle = new CellStyle { FontName = "Calibri" },
                ColumnStyles = new()
                {
                    { 1, new CellStyle { FontName = "Times New Roman" } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Col 0: default style (Calibri = same as built-in default, may read back as null)
            // Col 1: column style wins over default
            Assert.Equal("Times New Roman", read.Rows[0][1].Style?.FontName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FullPriorityChain_CellOverRowOverColumnOverDefault()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new() { Type = CellType.Text, Text = "cell", Style = new CellStyle { Bold = true } },
                        Cell.FromText("row"),
                        Cell.FromText("col"),
                    },
                },
                DefaultStyle = new CellStyle { FontName = "Default" },
                RowStyles = new()
                {
                    { 0, new CellStyle { FontName = "Row" } },
                },
                ColumnStyles = new()
                {
                    { 2, new CellStyle { FontName = "Col" } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Col 0: cell style wins (Bold, font may come from default font table)
            Assert.True(read.Rows[0][0].Style?.Bold);
            // Col 1: row style wins (no column style for col 1)
            Assert.Equal("Row", read.Rows[0][1].Style?.FontName);
            // Col 2: row style wins over column style
            Assert.Equal("Row", read.Rows[0][2].Style?.FontName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void HeaderStyle_OverridesColumnAndDefault()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                Rows = new() { new Cell[] { Cell.FromText("x"), Cell.FromText("y") } },
                HeaderStyle = new CellStyle { Bold = true, FillColor = "#4472C4" },
                ColumnStyles = new()
                {
                    { 0, new CellStyle { Italic = true } },
                },
                DefaultStyle = new CellStyle { FontName = "Arial" },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Data row should still get column/default styles
            Assert.True(read.Rows[0][0].Style?.Italic);
            Assert.Equal("Arial", read.Rows[0][1].Style?.FontName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RowColumnStyle_WithNumberFormat()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Price", "Qty" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(99.5), Cell.FromNumber(10) },
                    new Cell[] { Cell.FromNumber(49.9), Cell.FromNumber(20) },
                },
                ColumnStyles = new()
                {
                    { 0, new CellStyle { FontColor = "#FF0000" } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("#FF0000", read.Rows[0][0].Style?.FontColor);
            Assert.Equal("#FF0000", read.Rows[1][0].Style?.FontColor);
            Assert.Null(read.Rows[0][1].Style);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoStyles_ReadsNull()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Null(read.Rows[0][0].Style);
            Assert.Null(read.DefaultStyle);
            Assert.Null(read.RowStyles);
            Assert.Null(read.ColumnStyles);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
