using LiteExcel;
using System.IO;

namespace LiteExcel.Tests;

/// <summary>
/// 2.4.6+ 批 A：门面补齐（Append / ws.AutoColumnWidths / ReadWithProgress / GetSheetNames(stream)）。
/// </summary>
public class FacadeGapsTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"gap_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Excel_Append_DelegatesToWriter()
    {
        var file = GetTempFile();
        try
        {
            // 先写 2 行（表头 + 1 数据）
            var wb = Excel.Create("数据");
            wb.Worksheets[0].SetValue("A1", "ID");
            wb.Worksheets[0].SetValue("A2", 1);
            wb.SaveAs(file);

            Excel.Append(file, new SheetData
            {
                SheetName = "数据",
                Headers = new() { "ID" },
                Rows = new() { new Cell[] { Cell.FromNumber(2) } },
            });

            var read = Excel.Read<IdRow>(file);
            Assert.Equal(2, read.Count);
            Assert.Equal(1, read[0].Id);
            Assert.Equal(2, read[1].Id);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    private sealed class IdRow
    {
        [LiteColumn(Name = "ID")]
        public int Id { get; set; }
    }

    [Fact]
    public void Worksheet_AutoColumnWidths_FillsColumnWidths()
    {
        var ws = Excel.Create().Worksheets[0];
        ws.SetValue("A1", "姓名");           // 短
        ws.SetValue("B1", "这是一个非常非常非常长的列标题");  // 长
        ws.SetValue("A2", "张三"); ws.SetValue("B2", "北京");

        ws.AutoColumnWidths();

        Assert.NotNull(ws.ColumnWidths);
        Assert.True(ws.ColumnWidths!.ContainsKey(0));
        Assert.True(ws.ColumnWidths.ContainsKey(1));
        // B 列长文本应宽于 A 列
        Assert.True(ws.ColumnWidths[1] > ws.ColumnWidths[0]);
    }

    [Fact]
    public void Excel_ReadWithProgress_InvokesCallback()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("S");
            for (int i = 1; i <= 4; i++) { wb.Worksheets[0].SetValue($"A{i}", i); }
            wb.SaveAs(file);

            int calls = 0;
            int? lastCur = null; int? lastTotal = null;
            Excel.ReadWithProgress(file, 0, (cur, total) => { calls++; lastCur = cur; lastTotal = total; });
            Assert.True(calls >= 1, $"callback calls: {calls}");
            Assert.Equal(3, lastTotal);   // 4 行 - 第一行表头
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Excel_GetSheetNames_FromStream()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create("第一张");
            wb.Worksheets.Add("第二张");
            wb.SaveAs(file);
            using var fs = File.OpenRead(file);
            var names = Excel.GetSheetNames(fs);
            Assert.Equal(new[] { "第一张", "第二张" }, names);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
