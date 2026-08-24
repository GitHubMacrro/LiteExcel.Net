using LiteExcel;
using System.Data;
using System.IO;

namespace LiteExcel.Tests;

public class CreateWithDataTests
{
    private class Item
    {
        [LiteColumn(Name = "名称", Order = 0)]
        public string Name { get; set; } = "";

        [LiteColumn(Name = "数量", Order = 1)]
        public int Qty { get; set; }

        [LiteColumn(Ignore = true)]
        public string Internal { get; set; } = "";
    }

    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"cwd_{Guid.NewGuid():N}.xlsx");

    // ── Excel.Create<T> ──

    [Fact]
    public void Create_WithList_BuildsWorkbook()
    {
        var wb = Excel.Create(new List<Item>
        {
            new() { Name = "苹果", Qty = 3, Internal = "x" },
            new() { Name = "香蕉", Qty = 5, Internal = "y" },
        }, "水果");

        Assert.Equal("水果", wb.Worksheets[0].Name);
        Assert.Equal(1, wb.Worksheets.Count);

        var file = GetTempFile();
        try
        {
            wb.SaveAs(file);
            var rb = Excel.Open(file);
            var ws = rb.Worksheets[0];
            Assert.Equal("名称", ws.Cell("A1").Text);
            Assert.Equal("数量", ws.Cell("B1").Text);
            Assert.Equal("苹果", ws.Cell("A2").Text);
            Assert.Equal(3d, ws.Cell("B2").Number);
            Assert.Equal("香蕉", ws.Cell("A3").Text);
            Assert.DoesNotContain("Internal", ws.Cell("A1").Text + "|" + ws.Cell("B1").Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Create_WithList_CanMixWorkbookLevelFeatures()
    {
        var wb = Excel.Create(new List<Item> { new() { Name = "苹果", Qty = 3 } });
        wb.Worksheets[0].FreezeRows = 1;
        wb.Worksheets[0].HeaderStyle = new CellStyle { Bold = true };
        wb.Worksheets[0].ConditionalFormats.Add(new ConditionalFormat
        {
            Type = ConditionalFormatType.CellIs,
            Sqref = "B2:B100",
            Operator = ConditionalOperator.GreaterThan,
            Formula = "5",
            Style = new CellStyle { FontColor = "#FF0000" },
        });

        var file = GetTempFile();
        try
        {
            wb.SaveAs(file);
            var rb = Excel.Open(file);
            var ws = rb.Worksheets[0];
            Assert.Equal(1, ws.FreezeRows);
            Assert.Single(ws.ConditionalFormats);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Create_WithList_FluentConfigApplied()
    {
        var wb = Excel.Create(new List<Item> { new() { Name = "苹果", Qty = 3 } }, "S", configure: opt =>
        {
            opt.Column(x => x.Name, "品名");
            opt.Ignore(x => x.Qty);
        });
        Assert.Equal("品名", wb.Worksheets[0].Cell("A1").Text);
        Assert.True(string.IsNullOrEmpty(wb.Worksheets[0].Cell("B1").Text));
    }

    // ── Excel.Create(DataTable) ──

    [Fact]
    public void Create_WithDataTable_UsesTableName()
    {
        var dt = new DataTable("库存表");
        dt.Columns.Add("货号", typeof(string));
        dt.Columns.Add("金额", typeof(decimal));
        dt.Rows.Add("A1", 12.5m);

        var wb = Excel.Create(dt);
        Assert.Equal("库存表", wb.Worksheets[0].Name);

        var file = GetTempFile();
        try
        {
            wb.SaveAs(file);
            var rb = Excel.Open(file);
            var ws = rb.Worksheets[0];
            Assert.Equal("货号", ws.Cell("A1").Text);
            Assert.Equal("A1", ws.Cell("A2").Text);
            Assert.Equal(12.5d, ws.Cell("B2").Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Create_WithDataTable_EmptyTableName_FallsBackToSheet1()
    {
        var dt = new DataTable();
        dt.Columns.Add("C", typeof(int));
        var wb = Excel.Create(dt);
        Assert.Equal("Sheet1", wb.Worksheets[0].Name);
    }

    // ── Worksheet.ImportData ──

    [Fact]
    public void ImportData_ReplacesExistingContent()
    {
        var wb = Excel.Create("初始");
        var ws = wb.Worksheets[0];
        ws.Cell("A1").SetValue("旧数据");
        ws.Cell("B2").SetValue(99d);

        ws.ImportData(new List<Item> { new() { Name = "新数据", Qty = 1 } });

        // 清空语义：旧内容被替换，从 A1 重建（含表头 + 数据）
        Assert.Equal("名称", ws.Cell("A1").Text);
        Assert.Equal("新数据", ws.Cell("A2").Text);
        Assert.Equal(1d, ws.Cell("B2").Number);
    }

    [Fact]
    public void ImportData_DataTable_WithoutHeader()
    {
        var dt = new DataTable();
        dt.Columns.Add("A", typeof(string));
        dt.Rows.Add("v1");
        dt.Rows.Add("v2");

        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.ImportData(dt, includeHeader: false);

        Assert.Equal("v1", ws.Cell("A1").Text);
        Assert.Equal("v2", ws.Cell("A2").Text);
    }

    [Fact]
    public void ImportData_ThenStyleChange_TakesEffect()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets[0];
        ws.ImportData(new List<Item> { new() { Name = "苹果", Qty = 3 } });

        // 导入后修改单元格样式应生效（验证 RebindOwners 绑定了 Owner）
        ws.Cell("B2").Style = new CellStyle { Bold = true, FontColor = "#FF0000" };

        var file = GetTempFile();
        try
        {
            wb.SaveAs(file);
            var rb = Excel.Open(file);
            var cell = rb.Worksheets[0].Cell("B2");
            Assert.NotNull(cell.Style);
            Assert.True(cell.Style.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── WorksheetCollection.Add<T> ──

    [Fact]
    public void Add_WithList_BuildsSheet()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets.Add("第二张", new List<Item> { new() { Name = "苹果", Qty = 3 } });
        Assert.Equal("第二张", ws.Name);
        Assert.Equal("苹果", ws.Cell("A2").Text);
        Assert.Equal(2, wb.Worksheets.Count);
    }

    [Fact]
    public void Add_WithDataTable_BuildsSheet()
    {
        var dt = new DataTable();
        dt.Columns.Add("C", typeof(int));
        dt.Rows.Add(42);

        var wb = Excel.Create();
        var ws = wb.Worksheets.Add("表", dt);
        Assert.Equal("表", ws.Name);
        Assert.Equal("C", ws.Cell("A1").Text);
        Assert.Equal(42d, ws.Cell("A2").Number);
    }

    // ── 参数校验 ──

    [Fact]
    public void Create_WithNullData_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => Excel.Create((IEnumerable<Item>)null!));
    }

    [Fact]
    public void Create_WithNullDataTable_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => Excel.Create((DataTable)null!));
    }

    [Fact]
    public void ImportData_WithNull_Throws()
    {
        var ws = Excel.Create().Worksheets[0];
        Assert.Throws<ArgumentNullException>(() => ws.ImportData((IEnumerable<Item>)null!));
    }
}
