using System.IO.Compression;
using System.Text.RegularExpressions;
using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// XlsxStreamWriter 行数上限处理：Throw 模式抛异常、Spill 模式自动分表。
/// 用 internal MaxRowsPerSheet 钩子设小值，避免写百万行。
/// 行数验证直接数 sheet XML 的 &lt;row&gt; 标签（不受读取侧表头跳行影响）。
/// </summary>
public class RowLimitWriterTests
{
    private static string Tmp() => Path.Combine(Path.GetTempPath(), $"rlim_{Guid.NewGuid():N}.xlsx");

    private static XlsxStreamWriter CreateTestWriter(string path, RowLimitExceededMode mode, int maxRows)
    {
        var w = XlsxStreamWriter.Create(path, mode);
        w.MaxRowsPerSheet = maxRows;
        return w;
    }

    private static int CountRows(string path, int sheetIndex)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml");
        if (entry is null) return 0;
        using var s = entry.Open();
        using var sr = new StreamReader(s);
        return Regex.Matches(sr.ReadToEnd(), @"<row r=""").Count;
    }

    private static int SheetCount(string path)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        return zip.Entries.Count(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal));
    }

    [Fact]
    public void Throw_ReachedLimit_ThrowsAndFileValid()
    {
        var path = Tmp();
        var w = CreateTestWriter(path, RowLimitExceededMode.Throw, maxRows: 10);
        for (int i = 0; i < 10; i++)
            w.WriteRow(new object?[] { i, $"r{i}" });
        var ex = Assert.Throws<RowLimitExceededException>(() => w.WriteRow(new object?[] { 10, "r10" }));
        w.Dispose();
        Assert.Equal(11, ex.RowNumber);
        Assert.Equal(10, ex.MaxRows);
        Assert.Equal(1, SheetCount(path));
        Assert.Equal(10, CountRows(path, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Spill_OverLimit_CreatesSecondSheet()
    {
        var path = Tmp();
        using (var w = CreateTestWriter(path, RowLimitExceededMode.SpillToNewSheet, maxRows: 10))
        {
            for (int i = 0; i < 15; i++)
                w.WriteRow(new object?[] { i, $"r{i}" });
        }
        Assert.Equal(2, SheetCount(path));
        Assert.Equal(10, CountRows(path, 1));
        Assert.Equal(5, CountRows(path, 2));
        // 文件可被库读回
        var names = Excel.GetSheetNames(path);
        Assert.Equal(new[] { "Sheet1", "Sheet2" }, names.ToArray());
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Spill_ExactlyAtLimit_NoSpill()
    {
        var path = Tmp();
        using (var w = CreateTestWriter(path, RowLimitExceededMode.SpillToNewSheet, maxRows: 10))
        {
            for (int i = 0; i < 10; i++)
                w.WriteRow(new object?[] { i });
        }
        Assert.Equal(1, SheetCount(path));
        Assert.Equal(10, CountRows(path, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Spill_MultipleSpills()
    {
        var path = Tmp();
        using (var w = CreateTestWriter(path, RowLimitExceededMode.SpillToNewSheet, maxRows: 5))
        {
            for (int i = 0; i < 13; i++)
                w.WriteRow(new object?[] { i });
        }
        Assert.Equal(3, SheetCount(path));
        Assert.Equal(5, CountRows(path, 1));
        Assert.Equal(5, CountRows(path, 2));
        Assert.Equal(3, CountRows(path, 3));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void DefaultMode_IsThrow()
    {
        var path = Tmp();
        var w = CreateTestWriter(path, RowLimitExceededMode.Throw, maxRows: 3);
        for (int i = 0; i < 3; i++)
            w.WriteRow(new object?[] { i });
        Assert.Throws<RowLimitExceededException>(() => w.WriteRow(new object?[] { 3 }));
        w.Dispose();
        Assert.Equal(1, SheetCount(path));
        Assert.Equal(3, CountRows(path, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void FacadeCreateWriter_AcceptsModeAndSpills()
    {
        var path = Tmp();
        using (var w = Excel.CreateWriter(path, RowLimitExceededMode.SpillToNewSheet))
        {
            w.MaxRowsPerSheet = 4;
            for (int i = 0; i < 10; i++)
                w.WriteRow(new object?[] { i });
        }
        Assert.Equal(3, SheetCount(path));
        Assert.Equal(4, CountRows(path, 1));
        Assert.Equal(4, CountRows(path, 2));
        Assert.Equal(2, CountRows(path, 3));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Throw_Mode_LeavesPartialFileReadable()
    {
        // 异常后 using/Dispose 仍能正常 Close，文件结构完整
        var path = Tmp();
        try
        {
            var w = CreateTestWriter(path, RowLimitExceededMode.Throw, maxRows: 5);
            for (int i = 0; i < 5; i++)
                w.WriteRow(new object?[] { i });
            Assert.Throws<RowLimitExceededException>(() => w.WriteRow(new object?[] { 5 }));
            w.Dispose();
            // 可被 Excel.Open 读回
            var wb = Excel.Open(path);
            Assert.Equal("Sheet1", wb.Worksheets[0].Name);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Truncate_StopsAtLimit_AndSetsTruncated()
    {
        var path = Tmp();
        XlsxStreamWriter w;
        using (w = CreateTestWriter(path, RowLimitExceededMode.Truncate, maxRows: 10))
        {
            Assert.False(w.Truncated);
            for (int i = 0; i < 15; i++)
                w.WriteRow(new object?[] { i });
            Assert.True(w.Truncated);
        }
        Assert.True(w.Truncated);
        Assert.Equal(1, SheetCount(path));
        Assert.Equal(10, CountRows(path, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Truncate_ExactlyAtLimit_NotTruncated()
    {
        var path = Tmp();
        var w = CreateTestWriter(path, RowLimitExceededMode.Truncate, maxRows: 10);
        for (int i = 0; i < 10; i++)
            w.WriteRow(new object?[] { i });
        Assert.False(w.Truncated);
        w.Dispose();
        Assert.Equal(10, CountRows(path, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Spill_WithHeader_WritesHeaderOnEverySheet()
    {
        var path = Tmp();
        using (var w = XlsxStreamWriter.Create(path, RowLimitExceededMode.SpillToNewSheet, spillHeader: new object?[] { "ID", "Name" }))
        {
            w.MaxRowsPerSheet = 4; // 含表头，每表 1 表头 + 3 数据
            for (int i = 0; i < 8; i++)
                w.WriteRow(new object?[] { i, "row" + i });
        }
        // 8 数据行，每表 3 数据 -> 3 表（3+3+2）
        Assert.Equal(3, SheetCount(path));
        // 每表首行应为表头 "ID"
        Assert.Equal("ID", ReadCell(path, 1, 1, 1)); // sheet1 A1
        Assert.Equal("ID", ReadCell(path, 2, 1, 1)); // sheet2 A1
        Assert.Equal("ID", ReadCell(path, 3, 1, 1)); // sheet3 A1
        // sheet1 第 2 行起是数据 0
        Assert.Equal(0.0, CellNum(path, 1, 2, 1));
        Assert.Equal(7.0, CellNum(path, 3, 3, 1)); // sheet3 数据第 2 行 = i=7
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void Spill_WithoutHeader_SpilledSheetStartsWithData()
    {
        var path = Tmp();
        using (var w = CreateTestWriter(path, RowLimitExceededMode.SpillToNewSheet, maxRows: 4))
        {
            for (int i = 0; i < 6; i++)
                w.WriteRow(new object?[] { i });
        }
        Assert.Equal(2, SheetCount(path));
        // 无表头：sheet2 首行是数据（i=4）
        Assert.Equal(4.0, CellNum(path, 2, 1, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void SpillHeader_IgnoredInThrowMode()
    {
        var path = Tmp();
        var w = XlsxStreamWriter.Create(path, RowLimitExceededMode.Throw, spillHeader: new object?[] { "ID" });
        w.MaxRowsPerSheet = 5;
        // Throw 模式下 spillHeader 被忽略：首行应是调用方写的第 1 行
        w.WriteRow(new object?[] { 0 });
        w.Dispose();
        Assert.Equal(0.0, CellNum(path, 1, 1, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    [Fact]
    public void SpillHeader_IgnoredInTruncateMode()
    {
        var path = Tmp();
        var w = XlsxStreamWriter.Create(path, RowLimitExceededMode.Truncate, spillHeader: new object?[] { "ID" });
        w.MaxRowsPerSheet = 5;
        for (int i = 0; i < 3; i++)
            w.WriteRow(new object?[] { i });
        w.Dispose();
        // Truncate 模式下表头未写入：A1 是 0
        Assert.Equal(0.0, CellNum(path, 1, 1, 1));
        if (File.Exists(path)) File.Delete(path);
    }

    private static object? ReadCell(string path, int sheetIndex, int row, int col)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry($"xl/worksheets/sheet{sheetIndex}.xml");
        if (entry is null) return null;
        using var s = entry.Open();
        using var sr = new StreamReader(s);
        var xml = sr.ReadToEnd();
        var rowPat = $"<row r=\"{row}\"";
        var ri = xml.IndexOf(rowPat);
        if (ri < 0) return null;
        int idx = ri;
        for (int c = 0; c < col; c++)
        {
            idx = xml.IndexOf("<c ", idx + 1);
            if (idx < 0) return null;
        }
        var cEnd = xml.IndexOf("</c>", idx);
        if (cEnd < 0) cEnd = xml.IndexOf("/>", idx);
        var seg = xml.Substring(idx, cEnd - idx + 4);
        var vMatch = Regex.Match(seg, "<v>([^<]+)</v>");
        if (vMatch.Success && double.TryParse(vMatch.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num))
            return num;
        var tMatch = Regex.Match(seg, "<t>([^<]*)</t>");
        return tMatch.Success ? tMatch.Groups[1].Value : null;
    }

    private static double CellNum(string path, int sh, int row, int col)
        => (double)ReadCell(path, sh, row, col)!;
}
