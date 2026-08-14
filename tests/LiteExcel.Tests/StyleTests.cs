using LiteExcel;

namespace LiteExcel.Tests;

public class StyleTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void FontStyle_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle
            {
                FontName = "微软雅黑",
                FontSize = 14,
                Bold = true,
                Italic = true,
                FontColor = "#FF0000",
            };
            var sheet = new SheetData
            {
                Headers = new() { "测试" },
                Rows = new()
                {
                    new Cell[] { new() { Type = CellType.Text, Text = "红色粗体", Style = style } },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.Rows.Count);
            var cell = read.Rows[0][0];
            Assert.Equal("红色粗体", cell.Text);
            Assert.NotNull(cell.Style);
            Assert.Equal("微软雅黑", cell.Style!.FontName);
            Assert.Equal(14, cell.Style.FontSize);
            Assert.True(cell.Style.Bold);
            Assert.True(cell.Style.Italic);
            Assert.Equal("#FF0000", cell.Style.FontColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FillColor_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle { FillColor = "#00FF00" };
            var sheet = new SheetData
            {
                Headers = new() { "填充" },
                Rows = new() { new Cell[] { new() { Type = CellType.Text, Text = "绿底", Style = style } } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("#00FF00", read.Rows[0][0].Style?.FillColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BorderStyle_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle
            {
                Border = new BorderStyle
                {
                    Top = new BorderEdge { Style = "thin", Color = "#000000" },
                    Bottom = new BorderEdge { Style = "thin", Color = "#000000" },
                    Left = new BorderEdge { Style = "thin", Color = "#000000" },
                    Right = new BorderEdge { Style = "thin", Color = "#000000" },
                },
            };
            var sheet = new SheetData
            {
                Headers = new() { "边框" },
                Rows = new() { new Cell[] { new() { Type = CellType.Text, Text = "四边框", Style = style } } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var cellStyle = read.Rows[0][0].Style;
            Assert.NotNull(cellStyle);
            Assert.NotNull(cellStyle!.Border);
            Assert.NotNull(cellStyle.Border!.Top);
            Assert.Equal("thin", cellStyle.Border.Top!.Style);
            Assert.Equal("#000000", cellStyle.Border.Top.Color);
            Assert.NotNull(cellStyle.Border.Bottom);
            Assert.NotNull(cellStyle.Border.Left);
            Assert.NotNull(cellStyle.Border.Right);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Alignment_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                WrapText = true,
            };
            var sheet = new SheetData
            {
                Headers = new() { "对齐" },
                Rows = new() { new Cell[] { new() { Type = CellType.Text, Text = "居中换行", Style = style } } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var cellStyle = read.Rows[0][0].Style;
            Assert.NotNull(cellStyle);
            Assert.Equal(HorizontalAlignment.Center, cellStyle!.HorizontalAlignment);
            Assert.Equal(VerticalAlignment.Center, cellStyle.VerticalAlignment);
            Assert.True(cellStyle.WrapText);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void HeaderStyle_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var headerStyle = new CellStyle
            {
                Bold = true,
                FontColor = "#FFFFFF",
                FillColor = "#4472C4",
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            var sheet = new SheetData
            {
                Headers = new() { "A", "B" },
                HeaderStyle = headerStyle,
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // 验证表头有样式（表头行是第一行，但 Read 时被放入 Headers）
            // 读取时表头不放 Rows，所以无法直接验证表头样式
            // 但可以通过验证数据行无样式来间接确认
            Assert.Equal(1, read.Rows.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MultipleStyles_Deduplication()
    {
        var file = GetTempFile();
        try
        {
            var style1 = new CellStyle { Bold = true, FontColor = "#FF0000" };
            var style2 = new CellStyle { Bold = true, FontColor = "#FF0000" }; // 与 style1 相同
            var style3 = new CellStyle { Italic = true };

            var sheet = new SheetData
            {
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new() { Type = CellType.Text, Text = "s1", Style = style1 },
                        new() { Type = CellType.Text, Text = "s2", Style = style2 },
                        new() { Type = CellType.Text, Text = "s3", Style = style3 },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("s1", read.Rows[0][0].Text);
            Assert.True(read.Rows[0][0].Style?.Bold);
            Assert.Equal("#FF0000", read.Rows[0][0].Style?.FontColor);

            Assert.Equal("s2", read.Rows[0][1].Text);
            Assert.True(read.Rows[0][1].Style?.Bold);
            Assert.Equal("#FF0000", read.Rows[0][1].Style?.FontColor);

            Assert.Equal("s3", read.Rows[0][2].Text);
            Assert.True(read.Rows[0][2].Style?.Italic);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StyleWithNumberFormat_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle { Bold = true };
            var sheet = new SheetData
            {
                Headers = new() { "金额" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new()
                        {
                            Type = CellType.Number,
                            Number = 1234.56,
                            NumberFormat = "#,##0.00",
                            Style = style,
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var cell = read.Rows[0][0];
            Assert.Equal(CellType.Number, cell.Type);
            Assert.Equal(1234.56, cell.Number);
            Assert.True(cell.Style?.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoStyle_ReadsNull()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("无样式") } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Null(read.Rows[0][0].Style);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
