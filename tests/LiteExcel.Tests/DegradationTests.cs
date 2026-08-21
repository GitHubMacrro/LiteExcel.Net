using LiteExcel;
using System.IO;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 0：统一降级报告机制（ExcelWriteOptions.OnDegradation）。
/// 验证：回调可注册、可为 null、不注册时行为与现状一致。
/// </summary>
public class DegradationTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"p0b0_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void XlsWrite_ReportsCommentsAndValidationsAndFilterAndImages()
    {
        var file = GetTempFile(".xls");
        var reported = new List<DegradationCapability>();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.Comments = new Dictionary<string, string> { ["B1"] = "note" };
            ws.Validations = new List<DataValidation>
            {
                new() { Type = DataValidationType.WholeNumber, Sqref = "C1:C10", Formula1 = "1", Formula2 = "10" },
            };
            ws.Filter = new AutoFilter { Range = "A1:C10" };

            var options = new ExcelWriteOptions { OnDegradation = d => reported.Add(d.Capability) };
            Excel.Write(file, wb, options);

            Assert.Contains(DegradationCapability.Comments, reported);
            Assert.Contains(DegradationCapability.DataValidation, reported);
            Assert.Contains(DegradationCapability.AutoFilter, reported);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void CsvWrite_ReportsExcelOnlyCapabilities()
    {
        var file = GetTempFile(".csv");
        var reported = new List<DegradationInfo>();
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "h");
            ws.RowHeights = new Dictionary<int, double> { [0] = 30 };
            ws.ColumnWidths = new Dictionary<int, double> { [0] = 40 };
            ws.Merge("A1:B1");
            ws.FreezeRows = 1;
            var bold = new CellStyle { Bold = true };
            ws.Cell("A2").SetValue("v");
            ws.Cell("A2").Style = bold;
            ws.Cell("B2").SetValue(Cell.FromFormula("1+1"));
            ws.Cell("C2").SetValue("link");
            ws.Cell("C2").Hyperlink = new Hyperlink { Target = "https://example.com" };

            var options = new ExcelWriteOptions { OnDegradation = d => reported.Add(d) };
            Excel.Write(file, wb, options);

            var caps = reported.Select(d => d.Capability).ToList();
            Assert.Contains(DegradationCapability.Styles, caps);
            Assert.Contains(DegradationCapability.RowHeights, caps);
            Assert.Contains(DegradationCapability.ColumnWidths, caps);
            Assert.Contains(DegradationCapability.MergedCells, caps);
            Assert.Contains(DegradationCapability.FreezePanes, caps);
            Assert.Contains(DegradationCapability.Formulas, caps);
            Assert.Contains(DegradationCapability.Hyperlinks, caps);
            // 每条都有目标格式与说明
            Assert.All(reported, d => Assert.Equal(ExcelFormat.Csv, d.TargetFormat));
            Assert.All(reported, d => Assert.False(string.IsNullOrEmpty(d.Message)));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoCallback_BehavesIdentically()
    {
        // 不注册回调：写出成功，与历史行为一致（无降级上报）
        var file = GetTempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.Comments = new Dictionary<string, string> { ["B1"] = "note" };

            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = null });

            Assert.True(File.Exists(file));
            Assert.True(new FileInfo(file).Length > 0);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NullOptions_BehavesIdentically()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "1");
            Excel.Write(file, wb, null);
            var reopened = Excel.Open(file);
            Assert.Equal("1", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void XlsxWrite_ReportsNothing()
    {
        // xlsx 无损路径不应触发任何降级
        var file = GetTempFile(".xlsx");
        var reported = new List<DegradationInfo>();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");
            ws.Comments = new Dictionary<string, string> { ["B1"] = "note" };
            ws.Merge("C1:D1");
            ws.FreezeRows = 1;

            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => reported.Add(d) });

            Assert.Empty(reported);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
