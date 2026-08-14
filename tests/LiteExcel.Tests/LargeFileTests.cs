using LiteExcel;

namespace LiteExcel.Tests;

public class LargeFileTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void StreamRows_5000Rows()
    {
        var file = GetTempFile();
        try
        {
            const int rowCount = 5000;
            var rows = new List<IReadOnlyList<Cell>>(rowCount);
            for (int i = 0; i < rowCount; i++)
            {
                rows.Add(new Cell[]
                {
                    Cell.FromNumber(i),
                    Cell.FromText($"Row{i}"),
                    Cell.FromNumber(i * 1.5),
                });
            }
            var sheet = new SheetData
            {
                SheetName = "Big",
                Headers = new() { "ID", "Name", "Value" },
                Rows = rows,
            };
            XlsxWriter.Write(file, sheet);

            // 流式读取，不驻留内存
            int count = 0;
            int lastId = -1;
            XlsxReader.StreamRows(file, "Big", row =>
            {
                Assert.Equal(3, row.Count);
                Assert.Equal(count, (int)row[0].Number);
                Assert.Equal($"Row{count}", row[1].Text);
                count++;
                lastId = (int)row[0].Number;
            });

            Assert.Equal(rowCount, count);
            Assert.Equal(rowCount - 1, lastId);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void ReadAll_1000Rows_MatchesWritten()
    {
        var file = GetTempFile();
        try
        {
            const int rowCount = 1000;
            var rows = new List<IReadOnlyList<Cell>>(rowCount);
            for (int i = 0; i < rowCount; i++)
            {
                rows.Add(new Cell[]
                {
                    Cell.FromText($"Item-{i:D4}"),
                    Cell.FromNumber(i * 100),
                });
            }
            var sheet = new SheetData
            {
                SheetName = "Data",
                Headers = new() { "Name", "Count" },
                Rows = rows,
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal(rowCount, read.Rows.Count);

            // 抽检首尾
            Assert.Equal("Item-0000", read.Rows[0][0].Text);
            Assert.Equal(0, read.Rows[0][1].Number);
            Assert.Equal("Item-0999", read.Rows[999][0].Text);
            Assert.Equal(99900, read.Rows[999][1].Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
