using LiteExcel;
using System.IO.Compression;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 回归：手术式编辑（xlsb，仅改单元格内容）的字节补丁必须结构正确。
/// 曾有两处缺陷：替换与新增分两趟应用导致坐标错位、新行多单元格重复写行头。
/// 这里用结构断言（记录可完整解析、每行仅一条行头、行号升序）锁定，二者在 Excel 中都表现为文件损坏。
/// </summary>
public class SurgicalEditCellPatchTests
{
    private const int BrtRowHdr = 0x0000;

    private static string GetTempFile() =>
        Path.Combine(Path.GetTempPath(), $"surgcell_{Guid.NewGuid():N}.xlsb");

    private static int ReadVarInt(byte[] b, ref int pos)
    {
        int v = 0;
        for (int i = 0; i < 4; i++) { byte x = b[pos++]; v |= (x & 0x7F) << (7 * i); if ((x & 0x80) == 0) break; }
        return v;
    }

    private static byte[] ReadPart(string path, string entry)
    {
        using var zip = ZipFile.OpenRead(path);
        var e = zip.GetEntry(entry);
        Assert.NotNull(e);
        using var s = e!.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>解析 sheetN.bin，返回 (记录类型, 行号) 序列；记录越界即失败，乱码记录会在此暴露。</summary>
    private static List<(int Rt, int Row)> ParseRecords(byte[] part)
    {
        var list = new List<(int, int)>();
        int pos = 0;
        while (pos < part.Length)
        {
            int rt = ReadVarInt(part, ref pos);
            int cb = ReadVarInt(part, ref pos);
            Assert.True(cb >= 0 && pos + cb <= part.Length, $"记录越界：rt=0x{rt:X4} cb={cb} @{pos}/{part.Length}");
            int row = int.MinValue;
            if (rt == BrtRowHdr && cb >= 4) row = BitConverter.ToInt32(part, pos);
            list.Add((rt, row));
            pos += cb;
        }
        Assert.Equal(part.Length, pos);
        return list;
    }

    /// <summary>每行只能有一条行头，且行号严格升序。</summary>
    private static void AssertRowsWellFormed(List<(int Rt, int Row)> recs)
    {
        var rows = recs.Where(r => r.Rt == BrtRowHdr).Select(r => r.Row).ToList();
        Assert.Equal(rows.Count, rows.Distinct().Count());
        for (int i = 1; i < rows.Count; i++)
            Assert.True(rows[i] > rows[i - 1], $"行号未严格升序：{rows[i - 1]} → {rows[i]}");
    }

    /// <summary>构造只有一个工作表、已有两行数据的文件。</summary>
    private static string BuildTwoRowFile()
    {
        var spec = new XlsbTestFile.WorkbookSpec();
        var sheet = new XlsbTestFile.SheetSpec { Name = "S0" };
        sheet.Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "a1" } } });
        sheet.Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "a2" } } });
        spec.Sheets.Add(sheet);
        return XlsbTestFile.Build(spec);
    }

    [Fact]
    public void SurgicalEdit_ModifyExistingAndInsertNewCell_ProducesWellFormedSheet()
    {
        // 替换已有单元格的同时插入新单元格，补丁坐标不能错位。
        var src = BuildTwoRowFile();
        var outPath = GetTempFile();
        try
        {
            var wb = Excel.Open(src);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "changed");   // 已有：替换
            ws.SetValue("B2", "inserted");  // 新增：插入到已有 row1
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            var recs = ParseRecords(ReadPart(outPath, "xl/worksheets/sheet1.bin"));
            AssertRowsWellFormed(recs);

            var read = Excel.Open(outPath).Worksheets[0];
            Assert.Equal("changed", read.Cell("A1").Text);
            Assert.Equal("inserted", read.Cell("B2").Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void SurgicalEdit_WriteNewRowBlock_ProducesOneRowHdrPerRow()
    {
        // 一次写入整块新行，每行只应有一条行头。
        var src = BuildTwoRowFile();
        var outPath = GetTempFile();
        try
        {
            var wb = Excel.Open(src);
            var ws = wb.Worksheets[0];
            for (int r = 100; r <= 103; r++)
                for (int c = 1; c <= 3; c++)
                    ws.SetValue(r, c, $"{r}_{c}");
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            var recs = ParseRecords(ReadPart(outPath, "xl/worksheets/sheet1.bin"));
            AssertRowsWellFormed(recs);
            Assert.Equal(4, recs.Count(r => r.Rt == BrtRowHdr && r.Row >= 99)); // 新行各一条 RowHdr（A100→row0=99）

            var read = Excel.Open(outPath).Worksheets[0];
            Assert.Equal("100_1", read.Cell("A100").Text);
            Assert.Equal("103_3", read.Cell("C103").Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }

    [Fact]
    public void SurgicalEdit_ModifyCellThatShortensRecord_KeepsFollowingRowsIntact()
    {
        // 替换为更短的文本后再写后续新行，确认长度变化不会让后续插入错位。
        var spec = new XlsbTestFile.WorkbookSpec();
        var sheet = new XlsbTestFile.SheetSpec { Name = "S0" };
        sheet.Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "a-much-longer-header-value" } } });
        sheet.Rows.Add(new XlsbTestFile.RowSpec { Cells = { new XlsbTestFile.CellSpec { Col = 0, Text = "row2" } } });
        spec.Sheets.Add(sheet);
        var src = XlsbTestFile.Build(spec);
        var outPath = GetTempFile();
        try
        {
            var wb = Excel.Open(src);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "x");        // 缩短已有文本
            ws.SetValue("C3", "tail");     // 新增新行单元格
            wb.AllowFeatureLossOnSave = true;
            wb.SaveAs(outPath, ExcelFormat.Xlsb);

            var recs = ParseRecords(ReadPart(outPath, "xl/worksheets/sheet1.bin"));
            AssertRowsWellFormed(recs);

            var read = Excel.Open(outPath).Worksheets[0];
            Assert.Equal("x", read.Cell("A1").Text);
            Assert.Equal("row2", read.Cell("A2").Text);
            Assert.Equal("tail", read.Cell("C3").Text);
        }
        finally
        {
            if (File.Exists(src)) File.Delete(src);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }
}
