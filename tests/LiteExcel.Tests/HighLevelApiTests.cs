using LiteExcel;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

public class HighLevelApiTests
{
    private static string GetTempFile(string ext = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_hi_{Guid.NewGuid():N}{ext}");

    // ── Excel 门面 / Workbook ──

    [Fact]
    public void Create_AddSheet_SetValue_SaveAs_Open_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "姓名");
            ws.SetValue("B1", "年龄");
            ws.SetValue("A2", "张三");
            ws.SetValue("B2", 25);
            ws.SetValue("A3", "李四");
            ws.SetValue("B3", 30);
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Single(opened.Worksheets);
            var sheet = opened.Worksheets[0];
            Assert.Equal("姓名", sheet.Cell("A1").GetString());
            Assert.Equal("年龄", sheet.Cell("B1").GetString());
            Assert.Equal("张三", sheet.Cell("A2").GetString());
            Assert.Equal(25.0, sheet.Cell("B2").GetDouble());
            Assert.Equal(3, sheet.RowCount);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Open_ModifyCell_Save_Persists()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "原值");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("A1", "修改后");
            opened.Save();

            var reopened = Excel.Open(file);
            Assert.Equal("修改后", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SaveAs_UpdatesCurrentPath()
    {
        var file1 = GetTempFile();
        var file2 = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.SaveAs(file1);
            Assert.Equal(file1, wb.CurrentPath);

            wb.SaveAs(file2);
            Assert.Equal(file2, wb.CurrentPath);

            wb.Worksheets["Sheet1"].SetValue("A1", "x");
            wb.Save(); // 应保存到 file2
            Assert.True(File.Exists(file2));
            var reopened = Excel.Open(file2);
            Assert.Equal("x", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally
        {
            if (File.Exists(file1)) File.Delete(file1);
            if (File.Exists(file2)) File.Delete(file2);
        }
    }

    [Fact]
    public void Save_WithoutPath_Throws()
    {
        var wb = Excel.Create();
        Assert.Throws<LiteExcelException>(() => wb.Save());
    }

    [Fact]
    public void Create_UnsupportedFormat_Throws()
    {
        Assert.Throws<NotSupportedException>(() => Excel.Create((ExcelFormat)999));
    }

    [Fact]
    public void Open_InvalidXlsb_Throws()
    {
        var file = GetTempFile(".xlsb");
        try
        {
            File.WriteAllBytes(file, new byte[] { 0xD0, 0xCF, 0x11, 0xE0 });
            Assert.ThrowsAny<Exception>(() => Excel.Open(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── Worksheet / Cell / Cells / ExcelRange ──

    [Fact]
    public void Cell_ByAddress_And_ByIndex_Agree()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.SetValue("B3", "值");

        Assert.Equal("值", ws.Cell(3, 2).GetString());
        Assert.Equal("值", ws.Cell("B3").GetString());
        Assert.Equal("值", ws.Cells[3, 2].GetString());
        Assert.Equal("值", ws.Cells["B3"].GetString());
    }

    [Fact]
    public void Cell_OutOfRangeRead_ReturnsEmpty()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.SetValue("A1", "x");

        var cell = ws.Cell("Z99");
        Assert.True(cell.IsEmpty);
        Assert.Null(cell.GetString());
    }

    [Fact]
    public void SetValue_OutOfRange_WritesBack()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.SetValue("E10", 42);
        Assert.Equal(42.0, ws.Cell(10, 5).GetDouble());
        Assert.Equal(10, ws.RowCount);
        Assert.Equal(5, ws.MaxColumn);
    }

    [Fact]
    public void Cells_Enumeration_And_Clear()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.SetValue("A1", 1);
        ws.SetValue("B2", 2);

        // 使用区域包含 A1、B1（填充占位）、B2
        Assert.Equal(3, ws.Cells.Count());

        ws.Cells.Clear();
        Assert.All(ws.Cells, c => Assert.True(c.IsEmpty));
    }

    [Fact]
    public void Cells_Indexer_Set()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Cells[1, 1] = Cell.FromText("索引器");
        Assert.Equal("索引器", ws.Cell(1, 1).GetString());
    }

    [Fact]
    public void Range_Fill_And_ToValues()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        var range = ws.Range("A1:C3");
        range.Fill(7);

        var values = range.ToValues();
        Assert.Equal(3, values.GetLength(0));
        Assert.Equal(3, values.GetLength(1));
        Assert.All(values.Cast<object?>(), v => Assert.Equal(7.0, v));

        Assert.Equal("A1:C3", range.Address);
        Assert.Equal(3, range.RowCount);
        Assert.Equal(3, range.ColumnCount);
    }

    [Fact]
    public void Range_Enumeration_CountsCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        var range = ws.Range("A1:B2");
        Assert.Equal(4, range.Count());
    }

    [Fact]
    public void Range_Clear_EmptiesCells()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Range("A1:C3").Fill("x");
        ws.Range("A1:B2").Clear();

        Assert.True(ws.Cell("A1").IsEmpty);
        Assert.True(ws.Cell("B2").IsEmpty);
        Assert.Equal("x", ws.Cell("C3").GetString());
    }

    [Fact]
    public void Merge_And_Unmerge()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.Merge("A1:B2");
        Assert.Single(ws.MergedRanges);

        ws.Unmerge("A1:B2");
        Assert.Empty(ws.MergedRanges);
    }

    [Fact]
    public void WorksheetCollection_Operations()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("员工");
        wb.Worksheets.Add("工资");
        Assert.Equal(3, wb.Worksheets.Count);
        Assert.True(wb.Worksheets.Contains("工资"));
        Assert.True(wb.Worksheets.Contains("员工"));

        wb.Worksheets.Move(2, 0);
        Assert.Equal("工资", wb.Worksheets[0].Name);

        Assert.True(wb.Worksheets.Remove("员工"));
        Assert.False(wb.Worksheets.Contains("员工"));
        Assert.Throws<LiteExcelException>(() => wb.Worksheets["不存在"]);
    }

    // ── 便利 API 对称性 ──

    [Fact]
    public void ReadAsDataTable_And_Write_DataTable_Symmetry()
    {
        var file = GetTempFile();
        try
        {
            var dt = new System.Data.DataTable();
            dt.Columns.Add("名称", typeof(string));
            dt.Columns.Add("数量", typeof(int));
            dt.Rows.Add("苹果", 3);
            dt.Rows.Add("香蕉", 5);

            Excel.Write(file, dt, "水果");

            var read = Excel.ReadAsDataTable(file, "水果");
            Assert.Equal(2, read.Columns.Count);
            Assert.Equal(2, read.Rows.Count);
            Assert.Equal("苹果", read.Rows[0][0]);
            Assert.Equal("香蕉", read.Rows[1][0]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void GetSheetNames_Works()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets.Add("甲");
            wb.Worksheets.Add("乙");
            wb.SaveAs(file);

            var names = Excel.GetSheetNames(file);
            Assert.Equal(3, names.Count);
            Assert.Contains("甲", names);
            Assert.Contains("乙", names);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── Cell 便利方法 ──

    [Fact]
    public void Cell_Convenience_Getters()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        ws.SetValue("A1", "文本");
        ws.SetValue("A2", 3.14);
        ws.SetValue("A3", new DateTime(2024, 5, 6));
        ws.SetValue("A4", true);

        Assert.Equal("文本", ws.Cell("A1").GetString());
        Assert.True(ws.Cell("A2").TryGetDouble(out var d));
        Assert.Equal(3.14, d);
        Assert.True(ws.Cell("A3").TryGetDateTime(out var dt));
        Assert.Equal(new DateTime(2024, 5, 6), dt);
        Assert.True(ws.Cell("A4").GetBoolean());
        Assert.Equal("3.14", ws.Cell("A2").GetString());

        Assert.Throws<InvalidCastException>(() => ws.Cell("A1").GetDouble());
        Assert.False(ws.Cell("A1").TryGetDouble(out _));
    }

    [Fact]
    public void Cell_SetValue_Types()
    {
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        var cell = ws.Cell("A1");
        cell.SetValue(123);
        Assert.Equal(CellType.Number, cell.Type);
        Assert.Equal(123.0, cell.Number);

        cell.SetValue("abc");
        Assert.Equal(CellType.Text, cell.Type);
        Assert.Equal("abc", cell.Text);

        cell.SetValue(null);
        Assert.True(cell.IsEmpty);
    }

    [Fact]
    public void FillMergedCells_Option_ExpandsValues()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "合并值");
            ws.Merge("A1:B2");
            wb.SaveAs(file);

            var opened = Excel.Open(file, new ExcelReadOptions { FillMergedCells = true });
            Assert.Equal("合并值", opened.Worksheets[0].Cell("B2").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StreamRows_Works()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "头");
            ws.SetValue("A2", 1);
            ws.SetValue("A3", 2);
            wb.SaveAs(file);

            var rows = new List<IReadOnlyList<Cell>>();
            Excel.StreamRows(file, "Sheet1", rows.Add);
            Assert.Equal(2, rows.Count); // 跳过首行，剩 2 行数据
            Assert.Equal(1.0, rows[0][0].Number);
            Assert.Equal(2.0, rows[1][0].Number);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── CSV ──

    [Fact]
    public void Csv_Write_Open_RoundTrip()
    {
        var file = GetTempFile(".csv");
        try
        {
            var wb = Excel.Create(ExcelFormat.Csv);
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "名称");
            ws.SetValue("B1", "数量");
            ws.SetValue("A2", "苹果");
            ws.SetValue("B2", 3);
            wb.SaveAs(file);

            Assert.True(File.Exists(file));
            var text = File.ReadAllText(file);
            Assert.Contains("名称,数量", text);
            Assert.Contains("苹果,3", text);

            var opened = Excel.Open(file);
            var sheet = opened.Worksheets[0];
            Assert.Equal("苹果", sheet.Cell("A2").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Csv_QuotedField_RoundTrip()
    {
        var file = GetTempFile(".csv");
        try
        {
            var wb = Excel.Create(ExcelFormat.Csv);
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "含,逗号");
            ws.SetValue("A2", "含\"引号\"");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal("含,逗号", opened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal("含\"引号\"", opened.Worksheets[0].Cell("A2").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Csv_Write_MultiSheet_Throws()
    {
        var wb = Excel.Create(ExcelFormat.Csv);
        wb.Worksheets.Add("第二张");
        Assert.Throws<NotSupportedException>(() => wb.SaveAs(GetTempFile(".csv")));
    }

    // ── 流式写入 ──

    [Fact]
    public void StreamWriter_WritesAndReadsBack()
    {
        var file = GetTempFile();
        try
        {
            using (var writer = Excel.CreateWriter(file))
            {
                writer.WriteRow(new object?[] { "姓名", "年龄" });
                writer.WriteRow(new object?[] { "张三", 25 });
                writer.WriteRow(new object?[] { "李四", 30 });
            }

            var opened = Excel.Open(file);
            Assert.Equal("姓名", opened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(25.0, opened.Worksheets[0].Cell("B2").GetDouble());
            Assert.Equal(3, opened.Worksheets[0].RowCount);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StreamWriter_LargeRowCount_NoHeaderSplit()
    {
        var file = GetTempFile();
        try
        {
            using (var writer = Excel.CreateWriter(file))
            {
                for (int i = 0; i < 1000; i++)
                    writer.WriteRow(new object?[] { i, $"row{i}" });
            }

            var opened = Excel.Open(file);
            Assert.Equal(1000, opened.Worksheets[0].RowCount);
            Assert.Equal(0.0, opened.Worksheets[0].Cell("A1").GetDouble());
            Assert.Equal("row999", opened.Worksheets[0].Cell("B1000").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StreamWriter_WritesCorrectRowReferences()
    {
        var file = GetTempFile();
        try
        {
            using (var writer = Excel.CreateWriter(file))
            {
                writer.WriteRow(new object?[] { "r1", 1 });
                writer.WriteRow(new object?[] { "r2", 2 });
            }

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            using var stream = entry.Open();
            var doc = XDocument.Load(stream);
            var ns = doc.Root!.GetDefaultNamespace();
            var rows = doc.Root.Element(ns + "sheetData")!.Elements(ns + "row").ToList();

            Assert.Equal(2, rows.Count);
            Assert.Equal("1", rows[0].Attribute("r")?.Value);
            Assert.Equal("2", rows[1].Attribute("r")?.Value);
            var cellRefs = rows.SelectMany(r => r.Elements(ns + "c"))
                .Select(c => c.Attribute("r")?.Value).ToList();
            Assert.Equal(new[] { "A1", "B1", "A2", "B2" }, cellRefs);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void StreamWriter_WriteRow_AfterClose_Throws()
    {
        var file = GetTempFile();
        try
        {
            var writer = Excel.CreateWriter(file);
            writer.Close();
            Assert.Throws<InvalidOperationException>(() => writer.WriteRow(new object?[] { 1 }));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── 高级能力桥接（P8） ──

    [Fact]
    public void Formula_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", 1);
            ws.SetValue("A2", 2);
            ws.Cell("A3").SetValue(Cell.FromFormula("SUM(A1:A2)"));
            ws.Cell("A3").Number = 3; // 缓存值
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            var formulaCell = opened.Worksheets[0].Cell("A3");
            Assert.True(formulaCell.IsFormula);
            Assert.Equal("SUM(A1:A2)", formulaCell.Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Style_OnCell_And_Range()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.Cell("A1").Style = new CellStyle { Bold = true, FillColor = "#FFFF00" };
            ws.Range("B2:C3").Style = new CellStyle { Italic = true };
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.True(opened.Worksheets[0].Cell("A1").Style?.Bold);
            Assert.Equal("#FFFF00", opened.Worksheets[0].Cell("A1").Style?.FillColor);
            Assert.True(opened.Worksheets[0].Cell("C3").Style?.Italic);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Comments_And_Validations_And_Filter_Bridge()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "姓名");
            ws.SetValue("A2", "张三");
            ws.Comments = new Dictionary<string, string> { ["A2"] = "这是批注" };
            ws.Validations = new List<DataValidation>
            {
                new() { Type = DataValidationType.List, Sqref = "B2:B10", Formula1 = "\"男,女\"" },
            };
            ws.Filter = new AutoFilter { Range = "A1:B10" };
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            var read = opened.Worksheets[0];
            Assert.Equal("这是批注", read.Comments?["A2"]);
            Assert.Single(read.Validations!);
            Assert.Equal("B2:B10", read.Validations![0].Sqref);
            Assert.Equal("A1:B10", read.Filter?.Range);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FreezeHeader_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].FreezeHeader = true;
            wb.Worksheets["Sheet1"].SetValue("A1", "头");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.True(opened.Worksheets[0].FreezeHeader);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
