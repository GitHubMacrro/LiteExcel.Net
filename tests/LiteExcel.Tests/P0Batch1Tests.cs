using LiteExcel;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

/// <summary>
/// 第一个施工包测试：P0-1/2（列宽）、P0-8（公式缓存值）、P0-13（重复表名）。
/// 第二批：P0-6（definedNames/bookViews 保留）、P0-12（calcChain 不透传 + fullCalcOnLoad）。
/// </summary>
public class P0Batch1Tests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_p0_{Guid.NewGuid():N}{ext}");

    // ── P0-1：xlsx 列宽读取回填 ──

    [Fact]
    public void P0_1_Xlsx_ColumnWidths_RoundTrip()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var sheet = new SheetData
            {
                SheetName = "T",
                Headers = new() { "A", "B" },
                Rows = new() { new Cell[] { Cell.FromNumber(1), Cell.FromNumber(2) } },
                ColumnWidths = new List<double> { 15, 25 },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.ColumnWidths);
            Assert.Equal(15.0, read.ColumnWidths![0], 2);
            Assert.Equal(25.0, read.ColumnWidths[1], 2);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_1_Xlsx_NoColumnWidths_LeavesNull()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var sheet = new SheetData
            {
                SheetName = "T",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromNumber(1) } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            // 无自定义列宽时，ColumnWidths 应为 null（不写 <cols>）
            Assert.Null(read.ColumnWidths);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-2：稀疏列宽索引不错位 ──

    [Fact]
    public void P0_2_SparseColumnWidth_IndexPreserved()
    {
        // 高层：只设第 5 列（0-based key=4）
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            wb.Worksheets[0].ColumnWidths = new Dictionary<int, double> { [4] = 30 };
            wb.SaveAs(file);

            // 读回：第 5 列应为 30，其余为默认（0 哨兵或不存在）
            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.ColumnWidths);
            Assert.True(read.ColumnWidths!.Count >= 5);
            Assert.Equal(30.0, read.ColumnWidths[4], 2);
            // 前 4 列不应被误设为 30（旧 bug：丢 key 后 30 落到 index 0）
            Assert.NotEqual(30.0, read.ColumnWidths[0]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_2_MultipleSparseWidths_OrderStable()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            // 非连续、非顺序插入（Dictionary 枚举顺序无保证）
            wb.Worksheets[0].ColumnWidths = new Dictionary<int, double> { [7] = 40, [2] = 20, [5] = 30 };
            wb.SaveAs(file);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.ColumnWidths);
            Assert.Equal(20.0, read.ColumnWidths![2], 2);
            Assert.Equal(30.0, read.ColumnWidths[5], 2);
            Assert.Equal(40.0, read.ColumnWidths[7], 2);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-8：公式缓存值不被公式串覆盖 ──

    [Fact]
    public void P0_8_NumberFormula_CachedValuePreserved()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var cell = new Cell { Type = CellType.Number, Number = 42, Formula = "A1*2", IsFormula = true };
            var sheet = new SheetData
            {
                SheetName = "T",
                Headers = new() { "R" },
                Rows = new() { new Cell[] { cell } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var c = read.Rows[0][0];
            Assert.True(c.IsFormula);
            Assert.Equal("A1*2", c.Formula);
            // 缓存值保留在 Number，未被公式串覆盖
            Assert.Equal(42.0, c.GetDouble(), 9);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_8_FromFormula_TextIsNull_FormulaHoldsExpression()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var sheet = new SheetData
            {
                SheetName = "T",
                Headers = new() { "R" },
                Rows = new() { new Cell[] { Cell.FromFormula("SUM(A1:A3)") } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var c = read.Rows[0][0];
            Assert.True(c.IsFormula);
            Assert.Equal("SUM(A1:A3)", c.Formula);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_8_LegacyStyle_FormulaInText_StillWrites()
    {
        // 兼容垫片：旧代码 IsFormula=true + 公式存于 Text，写出仍应产出 <f>
        var file = GetTempFile(".xlsx");
        try
        {
            var sheet = new SheetData
            {
                SheetName = "T",
                Headers = new() { "R" },
                Rows = new() { new Cell[] { new Cell { Type = CellType.Text, Text = "A1+1", IsFormula = true } } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            var c = read.Rows[0][0];
            Assert.True(c.IsFormula);
            Assert.Equal("A1+1", c.Formula);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-13：低层 API 校验重复工作表名 ──

    [Fact]
    public void P0_13_DuplicateSheetNames_LowLevel_Throws()
    {
        var sheets = new List<SheetData>
        {
            new() { SheetName = "Dup", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromNumber(1) } } },
            new() { SheetName = "Dup", Headers = new() { "B" }, Rows = new() { new Cell[] { Cell.FromNumber(2) } } },
        };
        var file = GetTempFile(".xlsx");
        try
        {
            var ex = Assert.Throws<LiteExcelException>(() => XlsxWriter.Write(file, sheets));
            Assert.Contains("重复", ex.Message);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_13_DistinctSheetNames_LowLevel_OK()
    {
        var sheets = new List<SheetData>
        {
            new() { SheetName = "S1", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromNumber(1) } } },
            new() { SheetName = "S2", Headers = new() { "B" }, Rows = new() { new Cell[] { Cell.FromNumber(2) } } },
        };
        var file = GetTempFile(".xlsx");
        try
        {
            XlsxWriter.Write(file, sheets);
            var all = XlsxReader.ReadAll(file);
            Assert.Equal(2, all.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-6：definedNames / bookViews 原样回写 ──

    private static void InjectWorkbookElements(string file, string afterSheetsClose, string beforeSheetsOpen)
    {
        // 直接改写 xl/workbook.xml：注入 definedNames（</sheets> 后）与 bookViews（<sheets> 前）
        using var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update);
        var entry = zip.GetEntry("xl/workbook.xml")!;
        string xml;
        using (var s = entry.Open())
        using (var r = new StreamReader(s, Encoding.UTF8))
            xml = r.ReadToEnd();
        if (!string.IsNullOrEmpty(beforeSheetsOpen))
            xml = xml.Replace("<sheets>", beforeSheetsOpen + "<sheets>");
        if (!string.IsNullOrEmpty(afterSheetsClose))
            xml = xml.Replace("</sheets>", "</sheets>" + afterSheetsClose);
        entry.Delete();
        var newEntry = zip.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
        using (var s = newEntry.Open())
        using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            w.Write(xml);
    }

    [Fact]
    public void P0_6_DefinedNames_Preserved_OnSave()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            wb.SaveAs(file);

            InjectWorkbookElements(file,
                afterSheetsClose: "<definedNames><definedName name=\"MyRange\">Sheet1!$A$1:$A$10</definedName></definedNames>",
                beforeSheetsOpen: "");

            var opened = Excel.Open(file);
            opened.Worksheets[0].SetValue("B1", "y");
            opened.Save();

            // 读回 workbook.xml 验证 definedNames 保留
            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            using var s = zip.GetEntry("xl/workbook.xml")!.Open();
            using var r = new StreamReader(s, Encoding.UTF8);
            var xml = r.ReadToEnd();
            Assert.Contains("definedNames", xml);
            Assert.Contains("MyRange", xml);
            Assert.Contains("Sheet1!$A$1:$A$10", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void P0_6_BookViews_Preserved_OnSave()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", "x");
            wb.SaveAs(file);

            InjectWorkbookElements(file,
                afterSheetsClose: "",
                beforeSheetsOpen: "<bookViews><workbookView activeTab=\"0\"/></bookViews>");

            var opened = Excel.Open(file);
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            using var s = zip.GetEntry("xl/workbook.xml")!.Open();
            using var r = new StreamReader(s, Encoding.UTF8);
            var xml = r.ReadToEnd();
            Assert.Contains("bookViews", xml);
            Assert.Contains("activeTab", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ── P0-12：calcChain 不透传 + fullCalcOnLoad ──

    private static void InjectCalcChain(string file)
    {
        using var zip = new ZipArchive(File.Open(file, FileMode.Open, FileAccess.ReadWrite), ZipArchiveMode.Update);
        var cc = zip.CreateEntry("xl/calcChain.xml", CompressionLevel.Optimal);
        using (var s = cc.Open())
        using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><c r=\"A3\" i=\"1\"/></calcChain>");
        // 注入 workbook rel
        var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels")!;
        string rels;
        using (var s = relsEntry.Open())
        using (var r = new StreamReader(s, Encoding.UTF8))
            rels = r.ReadToEnd();
        rels = rels.Replace("</Relationships>",
            "<Relationship Id=\"rIdCC\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" Target=\"calcChain.xml\"/></Relationships>");
        relsEntry.Delete();
        var newRels = zip.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Optimal);
        using (var s = newRels.Open())
        using (var w = new StreamWriter(s, new UTF8Encoding(false)))
            w.Write(rels);
    }

    [Fact]
    public void P0_12_CalcChain_Dropped_And_FullCalcOnLoad_Written()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets[0].SetValue("A1", 1);
            wb.Worksheets[0].Cell("A2").SetValue(Cell.FromFormula("A1+1"));
            wb.SaveAs(file);
            InjectCalcChain(file);

            var opened = Excel.Open(file);
            opened.Save();

            using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
            // calcChain.xml 不应存在
            Assert.Null(zip.GetEntry("xl/calcChain.xml"));
            // workbook.xml 应含 fullCalcOnLoad
            using var s = zip.GetEntry("xl/workbook.xml")!.Open();
            using var r = new StreamReader(s, Encoding.UTF8);
            var xml = r.ReadToEnd();
            Assert.Contains("fullCalcOnLoad", xml);
            // calcChain rel 也不应存在
            using var rs = zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open();
            using var rr = new StreamReader(rs, Encoding.UTF8);
            var rels = rr.ReadToEnd();
            Assert.DoesNotContain("calcChain", rels);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
