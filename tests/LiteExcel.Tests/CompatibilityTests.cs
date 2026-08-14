using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Validates files authored by Microsoft Excel instead of only files emitted by LiteExcel.
/// The anonymous fixture is safe to publish and is required for this test suite.
/// </summary>
public class CompatibilityTests
{
    private static string FixturePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "excel-authored-compatibility.xlsx");

    private static string GetFixturePath()
    {
        Assert.True(File.Exists(FixturePath), $"Required Excel compatibility fixture is missing: {FixturePath}");
        return FixturePath;
    }

    [Fact]
    public void ReadSheetNames_FromExcelAuthoredWorkbook()
    {
        var names = XlsxReader.GetSheetNames(GetFixturePath());

        Assert.Equal(2, names.Count);
        Assert.Contains("Employees", names);
        Assert.Contains("Summary", names);
    }

    [Fact]
    public void ReadHeaders_AndDateCells_FromExcelAuthoredWorkbook()
    {
        var sheet = XlsxReader.Read(GetFixturePath(), "Employees");

        Assert.Equal(new[] { "Name", "Department", "StartDate", "Score" }, sheet.Headers);
        Assert.Equal(3, sheet.Rows.Count);
        Assert.Equal(CellType.Date, sheet.Rows[0][2].Type);
        Assert.Equal(new DateTime(2024, 1, 15), sheet.Rows[0][2].Date.Date);
    }

    [Fact]
    public void ReadExcelTableAndFrozenHeader_WithoutCrash()
    {
        var sheet = XlsxReader.Read(GetFixturePath(), "Employees");

        // Excel ListObject stores its auto filter in tableN.xml, which this
        // lightweight reader intentionally ignores. Reading must remain safe.
        Assert.Equal(3, sheet.Rows.Count);
        Assert.Equal("Alice", sheet.Rows[0][0].Text);
    }

    [Fact]
    public void ReadAll_IgnoresTableThemeAndExtendedParts()
    {
        var sheets = XlsxReader.ReadAll(GetFixturePath());

        Assert.Equal(2, sheets.Count);
        var employees = Assert.Single(sheets, s => s.SheetName == "Employees");
        var summary = Assert.Single(sheets, s => s.SheetName == "Summary");
        Assert.Equal("Name", employees.Headers[0]);
        Assert.Equal("Summary", summary.Headers[0]);
    }

    [Fact]
    public void RoundTrip_ExcelAuthoredWorkbook_PreservesDataShape()
    {
        var source = XlsxReader.Read(GetFixturePath(), "Employees");
        var output = Path.Combine(Path.GetTempPath(), $"liteexcel_compat_{Guid.NewGuid():N}.xlsx");
        try
        {
            XlsxWriter.Write(output, source);
            var readBack = XlsxReader.Read(output, 0);

            Assert.Equal(source.Headers, readBack.Headers);
            Assert.Equal(source.Rows.Count, readBack.Rows.Count);
            Assert.Equal(CellType.Date, readBack.Rows[0][2].Type);
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }
}
