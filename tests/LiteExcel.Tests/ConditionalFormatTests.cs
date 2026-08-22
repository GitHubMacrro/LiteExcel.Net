using LiteExcel;

namespace LiteExcel.Tests;

public class ConditionalFormatTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"p1c_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Write_And_ReadBack_CellIs()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs,
                Sqref = "B2:B10",
                Operator = ConditionalOperator.GreaterThan,
                Formula = "50",
                Style = new CellStyle { FontColor = "#FF0000" },
            });
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var cf = rb.Worksheets[0].ConditionalFormats;
            Assert.Single(cf);
            Assert.Equal("B2:B10", cf[0].Sqref);
            Assert.Equal(ConditionalFormatType.CellIs, cf[0].Type);
            Assert.Equal(ConditionalOperator.GreaterThan, cf[0].Operator);
            Assert.Equal("50", cf[0].Formula);
            Assert.NotNull(cf[0].Style);
            Assert.Equal("#FF0000", cf[0].Style!.FontColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Write_And_ReadBack_Expression()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.Expression,
                Sqref = "C1:C5",
                Formula = "ISODD(C1)",
                Style = new CellStyle { FillColor = "#FFF2CC" },
            });
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var cf = rb.Worksheets[0].ConditionalFormats[0];
            Assert.Equal(ConditionalFormatType.Expression, cf.Type);
            Assert.Equal("ISODD(C1)", cf.Formula);
            Assert.Equal("#FFF2CC", cf.Style?.FillColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Write_And_ReadBack_ColorScale()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.ColorScale,
                Sqref = "D1:D20",
                ColorScale = new ColorScaleInfo
                {
                    LowColor = "#FF0000",
                    MidColor = "#FFFF00",
                    HighColor = "#00FF00",
                },
            });
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var cf = rb.Worksheets[0].ConditionalFormats[0];
            Assert.Equal(ConditionalFormatType.ColorScale, cf.Type);
            Assert.NotNull(cf.ColorScale);
            Assert.Equal("#FF0000", cf.ColorScale!.LowColor);
            Assert.Equal("#FFFF00", cf.ColorScale.MidColor);
            Assert.Equal("#00FF00", cf.ColorScale.HighColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Write_And_ReadBack_DataBar()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.DataBar,
                Sqref = "E2:E100",
                DataBar = new DataBarInfo
                {
                    Color = "#63C384",
                    ShowValue = false,
                    MinLengthPercent = 5,
                    MaxLengthPercent = 90,
                },
            });
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var cf = rb.Worksheets[0].ConditionalFormats[0];
            Assert.Equal(ConditionalFormatType.DataBar, cf.Type);
            Assert.NotNull(cf.DataBar);
            Assert.Equal("#63C384", cf.DataBar!.Color);
            Assert.False(cf.DataBar.ShowValue);
            Assert.Equal(5, cf.DataBar.MinLengthPercent);
            Assert.Equal(90, cf.DataBar.MaxLengthPercent);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BetweenOperator_UsesTwoFormulas()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].ConditionalFormats.Add(new ConditionalFormat
            {
                Type = ConditionalFormatType.CellIs,
                Sqref = "A1:A10",
                Operator = ConditionalOperator.Between,
                Formula = "10",
                Formula2 = "20",
                Style = new CellStyle { Bold = true },
            });
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var cf = rb.Worksheets[0].ConditionalFormats[0];
            Assert.Equal(ConditionalOperator.Between, cf.Operator);
            Assert.Equal("10", cf.Formula);
            Assert.Equal("20", cf.Formula2);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
