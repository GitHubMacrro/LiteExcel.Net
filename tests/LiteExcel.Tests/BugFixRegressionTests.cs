using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// C1/C2/C3 回归测试：稀疏行 Cell 错位 / Modified 覆盖 / HeaderStyle/DefaultStyle 读回。
/// </summary>
public class BugFixRegressionTests
{
    private static string GetTempFile(string ext = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"bugfix_{Guid.NewGuid():N}{ext}");

    // ── C1：稀疏行 Cell 错位 ──

    [Fact]
    public void C1_SparseRows_CellAccess_AfterReopen()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            ws.SetValue("A2", "Row2Val");
            ws.SetValue("A4", "Row4Val");
            ws.SetValue("A7", "Row7Val");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var rws = rb.Worksheets["S"];
            Assert.Equal("Row2Val", rws.Cell("A2").GetString());
            Assert.Equal("Row4Val", rws.Cell("A4").GetString());
            Assert.Equal("Row7Val", rws.Cell("A7").GetString());
            // 空行不应返回错误数据
            Assert.True(rws.Cell("A1").IsEmpty);
            Assert.True(rws.Cell("A3").IsEmpty);
            Assert.True(rws.Cell("A5").IsEmpty);
            Assert.True(rws.Cell("A6").IsEmpty);
            Assert.True(rws.Cell("A8").IsEmpty);
            // Cell(row,col) / Cells[row,col] 一致
            Assert.Equal("Row2Val", rws.Cell(2, 1).GetString());
            Assert.Equal("Row4Val", rws.Cells[4, 1].GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void C1_FirstRowNotRow1_CellAccess_AfterReopen()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            ws.SetValue("A3", "V3");
            ws.SetValue("A4", "V4");
            ws.SetValue("A5", "V5");
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var rws = rb.Worksheets["S"];
            Assert.Equal("V3", rws.Cell("A3").GetString());
            Assert.Equal("V4", rws.Cell("A4").GetString());
            Assert.Equal("V5", rws.Cell("A5").GetString());
            Assert.True(rws.Cell("A1").IsEmpty);
            Assert.True(rws.Cell("A2").IsEmpty);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── C2：Modified 被打开流程覆盖 ──

    [Fact]
    public void C2_Modified_NotOverwrittenOnOpen()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            wb.Properties.Created = new DateTime(2024, 1, 1, 8, 30, 0, DateTimeKind.Local);
            wb.Properties.Modified = new DateTime(2024, 6, 1, 18, 0, 0, DateTimeKind.Local);
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            // Modified 应保持为文件原始值，不被打开时刻覆盖
            Assert.NotNull(rb.Properties.Modified);
            Assert.Equal(new DateTime(2024, 6, 1, 18, 0, 0), rb.Properties.Modified!.Value);
            Assert.Equal(new DateTime(2024, 1, 1, 8, 30, 0), rb.Properties.Created!.Value);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── C3：DefaultStyle 读回 ──

    [Fact]
    public void C3_DefaultStyle_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            ws.SetValue("A1", "Header");
            ws.SetValue("A2", "x");
            ws.DefaultStyle = new CellStyle { FontName = "Arial", FontSize = 12 };
            wb.SaveAs(file);

            var rb = Excel.Open(file);
            var rws = rb.Worksheets["S"];
            // DefaultStyle 应能读回（所有单元格共享同一样式时可恢复）
            Assert.NotNull(rws.DefaultStyle);
            Assert.Equal("Arial", rws.DefaultStyle!.FontName);
            Assert.Equal(12, rws.DefaultStyle.FontSize);
            // 单元格样式不回归
            Assert.NotNull(rws.Cell("A1").Style);
            Assert.Equal("Arial", rws.Cell("A1").Style!.FontName);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void C3_HeaderStyle_RoundTrip_LowLevel()
    {
        var file = GetTempFile();
        try
        {
            var sd = new SheetData
            {
                SheetName = "S",
                Headers = new() { "H1", "H2" },
                HeaderStyle = new CellStyle { Bold = true, FillColor = "#00FF00" },
                Rows = new() { new Cell[] { Cell.FromText("x"), Cell.FromText("y") } },
            };
            XlsxWriter.Write(file, sd);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.HeaderStyle);
            Assert.True(read.HeaderStyle!.Bold);
            Assert.Equal("#00FF00", read.HeaderStyle.FillColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
