using LiteExcel;

namespace LiteExcel.Tests;

public class MergeTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void SingleMerge_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Merge",
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("1"), Cell.FromText("2"), Cell.FromText("3") },
                    new Cell[] { Cell.FromText("4"), Cell.FromText("5"), Cell.FromText("6") },
                },
                MergedRanges = new()
                {
                    new CellRange(0, 0, 0, 2), // merge A1:C1 in data rows (row 0)
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.MergedRanges.Count);
            var range = read.MergedRanges[0];
            Assert.Equal(0, range.FirstRow);
            Assert.Equal(0, range.LastRow);
            Assert.Equal(0, range.FirstCol);
            Assert.Equal(2, range.LastCol);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MultipleMerges_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "MultiMerge",
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
                    new Cell[] { Cell.FromText("z"), Cell.FromText("w") },
                    new Cell[] { Cell.FromText("a"), Cell.FromText("b") },
                },
                MergedRanges = new()
                {
                    new CellRange(0, 0, 0, 1), // A1:B1
                    new CellRange(1, 2, 0, 0), // A2:A3
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0, firstRowIsHeader: false);
            Assert.Equal(2, read.MergedRanges.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoMerges_ReadsEmpty()
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
            Assert.Empty(read.MergedRanges);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MergeWithStyle_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var style = new CellStyle
            {
                Bold = true,
                FillColor = "#FFFF00",
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            var sheet = new SheetData
            {
                Headers = new() { "A", "B", "C" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new() { Type = CellType.Text, Text = "合并", Style = style },
                        Cell.Empty,
                        Cell.Empty,
                    },
                },
                MergedRanges = new()
                {
                    new CellRange(1, 1, 0, 2), // A2:C2 (row 1, col 0-2)
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(1, read.MergedRanges.Count);
            Assert.Equal(1, read.Rows.Count);
            var cell = read.Rows[0][0];
            Assert.Equal("合并", cell.Text);
            Assert.NotNull(cell.Style);
            Assert.True(cell.Style!.Bold);
            Assert.Equal("#FFFF00", cell.Style.FillColor);
            Assert.Equal(HorizontalAlignment.Center, cell.Style.HorizontalAlignment);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void LargeMergeRange_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var rows = new List<IReadOnlyList<Cell>>();
            for (int i = 0; i < 5; i++)
            {
                rows.Add(new Cell[] { Cell.FromText($"R{i}"), Cell.FromText($"C{i}") });
            }
            var sheet = new SheetData
            {
                SheetName = "LargeMerge",
                Headers = new() { "Col1", "Col2" },
                Rows = rows,
                MergedRanges = new()
                {
                    new CellRange(0, 4, 0, 0), // A1:A5
                    new CellRange(0, 4, 1, 1), // B1:B5
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(2, read.MergedRanges.Count);
            var r1 = read.MergedRanges[0];
            Assert.Equal(0, r1.FirstRow);
            Assert.Equal(4, r1.LastRow);
            Assert.Equal(0, r1.FirstCol);
            Assert.Equal(0, r1.LastCol);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
