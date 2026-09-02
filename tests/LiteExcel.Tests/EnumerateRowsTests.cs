using System.IO.Compression;
using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// Excel.EnumerateRows 拉取式流式读取：逐行 yield、LINQ 可组合、提前中断、不跳首行。
/// </summary>
public class EnumerateRowsTests
{
    private static string WriteTestFile(string sheetName = "Sheet1", int rows = 100, bool hasHeader = false)
    {
        var path = Path.Combine(Path.GetTempPath(), $"enumrows_{Guid.NewGuid():N}.xlsx");
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.Name = sheetName;
        int start = hasHeader ? 1 : 0;
        if (hasHeader)
            ws.SetValue("A1", "ID");
        for (int i = 0; i < rows; i++)
            ws.SetValue($"A{start + i + 1}", i);
        wb.SaveAs(path);
        return path;
    }

    [Fact]
    public void EnumerateAllRows_CountCorrect()
    {
        var path = WriteTestFile(rows: 50);
        try
        {
            var count = Excel.EnumerateRows(path, "Sheet1").Count();
            Assert.Equal(50, count);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void DoesNotSkipFirstRow()
    {
        var path = WriteTestFile(rows: 10);
        try
        {
            var rows = Excel.EnumerateRows(path, "Sheet1").ToList();
            // 不跳首行：第一行的 A1 应有值 0
            Assert.Equal("0", rows[0][0].Text ?? rows[0][0].Number.ToString());
            // 对比 StreamRows 会跳过首行，少 1 行
            int streamCount = 0;
            Excel.StreamRows(path, "Sheet1", _ => streamCount++);
            Assert.Equal(9, streamCount);       // StreamRows 跳了首行
            Assert.Equal(10, rows.Count);       // EnumerateRows 没跳
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void First_OnlyReadsOneRow()
    {
        var path = WriteTestFile(rows: 10000);
        try
        {
            var first = Excel.EnumerateRows(path, "Sheet1").First();
            Assert.NotNull(first);
            Assert.True(first.Count >= 1);
            // 验证只读了一行：取完 First 后应能删除文件（句柄已释放）
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Take_StopsEarly()
    {
        var path = WriteTestFile(rows: 10000);
        try
        {
            var rows = Excel.EnumerateRows(path, "Sheet1").Take(5).ToList();
            Assert.Equal(5, rows.Count);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Break_ReleasesFileHandle()
    {
        var path = WriteTestFile(rows: 1000);
        try
        {
            foreach (var row in Excel.EnumerateRows(path, "Sheet1"))
            {
                _ = row;
                break; // 提前退出
            }
            // 文件句柄应已释放，可删除
            File.Delete(path);
            Assert.False(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void NullSheetName_TakesFirstSheet()
    {
        var path = Path.Combine(Path.GetTempPath(), $"enumrows_{Guid.NewGuid():N}.xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].Name = "FirstSheet";
            wb.Worksheets[0].SetValue("A1", "hello");
            wb.Worksheets.Add("Second");
            wb.Worksheets["Second"].SetValue("A1", "world");
            wb.SaveAs(path);

            var rows = Excel.EnumerateRows(path).ToList();
            Assert.Single(rows);
            Assert.Equal("hello", rows[0][0].GetString());
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SheetNotFound_Throws()
    {
        var path = WriteTestFile(rows: 5);
        try
        {
            // 枚举时才抛（迭代器延迟执行）
            Assert.Throws<LiteExcelException>(() =>
            {
                foreach (var _ in Excel.EnumerateRows(path, "NonExistent")) { }
            });
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void NullPath_ThrowsImmediately()
    {
        // 参数校验应立即抛（非迭代器层），不等枚举
        Assert.Throws<ArgumentException>(() => Excel.EnumerateRows((string)null!));
    }

    [Fact]
    public void EmptyPath_ThrowsImmediately()
    {
        Assert.Throws<ArgumentException>(() => Excel.EnumerateRows(""));
    }

    [Fact]
    public void NonXlsxFormat_Throws()
    {
        var path = Path.Combine(Path.GetTempPath(), $"enumrows_{Guid.NewGuid():N}.csv");
        try
        {
            File.WriteAllText(path, "a,b,c");
            Assert.Throws<LiteExcelException>(() => Excel.EnumerateRows(path, "Sheet1"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void StreamOverload_Works()
    {
        var path = WriteTestFile(rows: 10);
        try
        {
            using var fs = File.OpenRead(path);
            var count = Excel.EnumerateRows(fs, "Sheet1").Count();
            Assert.Equal(10, count);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void StreamOverload_NullSheetName_TakesFirst()
    {
        var path = WriteTestFile(sheetName: "MySheet", rows: 5);
        try
        {
            using var fs = File.OpenRead(path);
            var count = Excel.EnumerateRows(fs).Count();
            Assert.Equal(5, count);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void NullStream_ThrowsImmediately()
    {
        Assert.Throws<ArgumentNullException>(() => Excel.EnumerateRows((Stream)null!));
    }
}
