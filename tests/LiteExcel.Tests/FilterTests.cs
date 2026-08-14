using LiteExcel;

namespace LiteExcel.Tests;

public class FilterTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void EqualsFilter_WritesAndReads()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Filter",
                Headers = new() { "Name", "City" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Zhang"), Cell.FromText("Beijing") },
                    new Cell[] { Cell.FromText("Li"), Cell.FromText("Shanghai") },
                    new Cell[] { Cell.FromText("Wang"), Cell.FromText("Beijing") },
                    new Cell[] { Cell.FromText("Zhao"), Cell.FromText("Guangzhou") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:B5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 1,
                            Type = FilterType.Equals,
                            Values = new() { "Beijing" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Filter);
            Assert.Equal("A1:B5", read.Filter!.Range);
            Assert.Single(read.Filter.Columns);
            Assert.Equal(1, read.Filter.Columns[0].ColumnIndex);
            Assert.Equal(FilterType.Equals, read.Filter.Columns[0].Type);
            Assert.Contains("Beijing", read.Filter.Columns[0].Values);

            // Rows 1 (Shanghai) and 3 (Guangzhou) should be hidden
            Assert.Contains(1, read.Filter.HiddenRows);
            Assert.Contains(3, read.Filter.HiddenRows);
            Assert.DoesNotContain(0, read.Filter.HiddenRows);
            Assert.DoesNotContain(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void CompareFilter_GreaterThan()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "NumFilter",
                Headers = new() { "Value" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(10) },
                    new Cell[] { Cell.FromNumber(50) },
                    new Cell[] { Cell.FromNumber(100) },
                    new Cell[] { Cell.FromNumber(5) },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.Compare,
                            Operator = FilterOperator.GreaterThan,
                            Values = new() { "20" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Filter);
            // Rows with value <= 20 should be hidden: 10 (idx 0) and 5 (idx 3)
            Assert.Contains(0, read.Filter!.HiddenRows);
            Assert.Contains(3, read.Filter.HiddenRows);
            Assert.DoesNotContain(1, read.Filter.HiddenRows);
            Assert.DoesNotContain(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BetweenFilter()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Score" },
                Rows = new()
                {
                    new Cell[] { Cell.FromNumber(30) },
                    new Cell[] { Cell.FromNumber(60) },
                    new Cell[] { Cell.FromNumber(80) },
                    new Cell[] { Cell.FromNumber(95) },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.Compare,
                            Operator = FilterOperator.Between,
                            MinValue = "50",
                            MaxValue = "90",
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // 30 (idx 0) and 95 (idx 3) should be hidden
            Assert.Contains(0, read.Filter!.HiddenRows);
            Assert.Contains(3, read.Filter.HiddenRows);
            Assert.DoesNotContain(1, read.Filter.HiddenRows);
            Assert.DoesNotContain(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ContainsFilter()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Text" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("apple") },
                    new Cell[] { Cell.FromText("banana") },
                    new Cell[] { Cell.FromText("cherry") },
                    new Cell[] { Cell.FromText("pineapple") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.Contains,
                            Values = new() { "app" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // banana (idx 1) and cherry (idx 2) should be hidden
            Assert.Contains(1, read.Filter!.HiddenRows);
            Assert.Contains(2, read.Filter.HiddenRows);
            Assert.DoesNotContain(0, read.Filter.HiddenRows);
            Assert.DoesNotContain(3, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BeginsWithFilter()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Text" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("apple") },
                    new Cell[] { Cell.FromText("apricot") },
                    new Cell[] { Cell.FromText("banana") },
                    new Cell[] { Cell.FromText("cherry") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.BeginsWith,
                            Values = new() { "ap" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Contains(2, read.Filter!.HiddenRows);
            Assert.Contains(3, read.Filter.HiddenRows);
            Assert.DoesNotContain(0, read.Filter.HiddenRows);
            Assert.DoesNotContain(1, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void EndsWithFilter()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Text" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("test.txt") },
                    new Cell[] { Cell.FromText("test.csv") },
                    new Cell[] { Cell.FromText("data.txt") },
                    new Cell[] { Cell.FromText("data.csv") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.EndsWith,
                            Values = new() { ".txt" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Contains(1, read.Filter!.HiddenRows);
            Assert.Contains(3, read.Filter.HiddenRows);
            Assert.DoesNotContain(0, read.Filter.HiddenRows);
            Assert.DoesNotContain(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void BlankFilter()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Val" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x") },
                    new Cell[] { Cell.Empty },
                    new Cell[] { Cell.FromText("y") },
                    new Cell[] { Cell.Empty },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A5",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.Blank,
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Non-blank rows should be hidden: 0 (x) and 2 (y)
            Assert.Contains(0, read.Filter!.HiddenRows);
            Assert.Contains(2, read.Filter.HiddenRows);
            Assert.DoesNotContain(1, read.Filter.HiddenRows);
            Assert.DoesNotContain(3, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ManualHiddenRows()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "A" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("1") },
                    new Cell[] { Cell.FromText("2") },
                    new Cell[] { Cell.FromText("3") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:A4",
                    HiddenRows = new() { 1 }, // manually hide row index 1
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Contains(1, read.Filter!.HiddenRows);
            Assert.DoesNotContain(0, read.Filter.HiddenRows);
            Assert.DoesNotContain(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoFilter_RoundTrip()
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
            Assert.Null(read.Filter);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MultipleConditions_AllMustMatch()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                Headers = new() { "Name", "City" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Zhang"), Cell.FromText("Beijing") },
                    new Cell[] { Cell.FromText("Li"), Cell.FromText("Beijing") },
                    new Cell[] { Cell.FromText("Wang"), Cell.FromText("Shanghai") },
                },
                Filter = new AutoFilter
                {
                    Range = "A1:B4",
                    Columns = new()
                    {
                        new FilterColumn
                        {
                            ColumnIndex = 0,
                            Type = FilterType.Equals,
                            Values = new() { "Zhang" },
                        },
                        new FilterColumn
                        {
                            ColumnIndex = 1,
                            Type = FilterType.Equals,
                            Values = new() { "Beijing" },
                        },
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // Only row 0 (Zhang/Beijing) matches both conditions
            Assert.DoesNotContain(0, read.Filter!.HiddenRows);
            Assert.Contains(1, read.Filter.HiddenRows);
            Assert.Contains(2, read.Filter.HiddenRows);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
