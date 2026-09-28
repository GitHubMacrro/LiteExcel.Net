using System.IO.Compression;
using System.Text;
using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// C1-C3 修复后回归测试补充（针对上一轮只读验收发现的盲区）。
/// 分类约定：
///  - Safe：行为与修复前契约一致 / 正确往返。
///  - Confirmed Regression：本次 C1 绝对网格修复引入的错位（测试断言“正确行为”，预期失败）。
///  - Existing Limitation：并非本次修改造成，部分结构读取侧本就未恢复。
///  - Behavior Change / Heuristic：行为确实变化或为启发式推断，测试固定“当前行为”并要求文档决策。
/// </summary>
public class BugFixAcceptanceTests
{
    private const string PkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static string Tmp(string ext = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"bugacc_{Guid.NewGuid():N}{ext}");

    private static void WriteEntry(ZipArchive zip, string entry, string content)
    {
        var e = zip.GetEntry(entry) ?? zip.CreateEntry(entry);
        using var s = e.Open();
        s.SetLength(0);
        var bytes = Encoding.UTF8.GetBytes(content);
        s.Write(bytes, 0, bytes.Length);
    }

    private static string ReadEntry(string file, string entry)
    {
        using var z = ZipFile.OpenRead(file);
        var e = z.GetEntry(entry);
        if (e is null) return "<missing>";
        using var r = new StreamReader(e.Open());
        return r.ReadToEnd();
    }

    /// <summary>用原始 OOXML 构造一个最小 xlsx（可控稀疏行 / hidden / ht / 样式）。</summary>
    private static void BuildRaw(string file, string sheetDataXml, string? stylesXml = null)
    {
        using var zip = new ZipArchive(File.Open(file, FileMode.Create, FileAccess.ReadWrite), ZipArchiveMode.Update);
        WriteEntry(zip, "xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<worksheet xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\">" + sheetDataXml + "</worksheet>");
        WriteEntry(zip, "xl/sharedStrings.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<sst xmlns=\"{MainNs}\" count=\"8\" uniqueCount=\"8\">" +
            "<si><t>V2</t></si><si><t>V3</t></si><si><t>V4</t></si><si><t>V5</t></si>" +
            "<si><t>V6</t></si><si><t>V7</t></si></sst>");
        WriteEntry(zip, "xl/styles.xml", stylesXml ??
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<styleSheet xmlns=\"{MainNs}\"><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>");
        WriteEntry(zip, "xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\">" +
            "<sheets><sheet name=\"S\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
        WriteEntry(zip, "xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<Relationships xmlns=\"{PkgRelNs}\">" +
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
            $"<Relationship Id=\"rId2\" Type=\"{OfficeRelNs}/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
            $"<Relationship Id=\"rId3\" Type=\"{OfficeRelNs}/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>");
        WriteEntry(zip, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            $"<Relationships xmlns=\"{PkgRelNs}\">" +
            $"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");
        WriteEntry(zip, "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
            "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
            "</Types>");
    }

    // ─────────────────────────────────────────────────────────────────────
    // 必须测试 1：Filter.HiddenRows + FirstRow 偏移
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// [Confirmed Regression] 前导偏移文件：row5 hidden。经对象模型 Open→Save 后，
    /// hidden 被错误地移动到 row3（偏移 FirstRowNumber-1=2 未对 Filter.HiddenRows 做转换）。
    /// 本测试断言“正确行为”（hidden 仍在 row5），预期失败，用于锁定回归。
    /// </summary>
    [Fact]
    public void C1_Filter_HiddenRows_LeadingOffset_PreservedAfterObjectModelSave()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            // 数据行 3..7（前导 2 空行），row5 隐藏，filter A3:A7
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>1</v></c></row>" +
                "<row r=\"4\"><c r=\"A4\" t=\"s\"><v>2</v></c></row>" +
                "<row r=\"5\" hidden=\"1\"><c r=\"A5\" t=\"s\"><v>3</v></c></row>" +
                "<row r=\"6\"><c r=\"A6\" t=\"s\"><v>4</v></c></row>" +
                "<row r=\"7\"><c r=\"A7\" t=\"s\"><v>5</v></c></row>" +
                "</sheetData>" +
                "<autoFilter ref=\"A3:A7\"></autoFilter>");

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            Assert.NotNull(ws.Filter);
            // 绝对 0-based：隐藏行 row5 -> 索引 4
            Assert.Contains(4, ws.Filter!.HiddenRows);

            rb.SaveAs(file2);
            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"5\" hidden=\"1\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    /// <summary>
    /// [Confirmed Regression] 中间存在空行（数据行 2,4,7）：row4 设行高。
    /// 经对象模型 Open→Save 后，RowHeights 的 key 只加了前导偏移（mergeOffset），未用 RowNumbers
    /// 处理内部间隙，导致行高落到空行 3 而丢失。断言“正确行为”（row4 仍有 ht），预期失败。
    /// </summary>
    [Fact]
    public void C1_RowHeights_InteriorGap_PreservedAfterObjectModelSave()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>" +
                "<row r=\"4\" ht=\"33\" customHeight=\"1\"><c r=\"A4\" t=\"s\"><v>2</v></c></row>" +
                "<row r=\"7\"><c r=\"A7\" t=\"s\"><v>5</v></c></row>" +
                "</sheetData>");

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            // 绝对 0-based：row4 -> 索引 3
            Assert.NotNull(ws.RowHeights);
            Assert.True(ws.RowHeights!.ContainsKey(3), $"RowHeights keys=[{string.Join(",", ws.RowHeights.Keys)}]，期望含绝对索引 3(row4)");

            rb.SaveAs(file2);
            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"4\" ht=\"33\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    /// <summary>
    /// [Safe] 仅前导偏移（数据行 3,4,5 连续）：RowHeights 经对象模型往返后仍对应正确绝对行。
    /// </summary>
    [Fact]
    public void C1_RowHeights_LeadingOffset_RoundTripsCorrectly()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            var sd = new SheetData
            {
                SheetName = "S",
                FirstRowNumber = 3,
                Rows = new()
                {
                    new Cell[] { Cell.FromText("r3") },
                    new Cell[] { Cell.FromText("r4") },
                    new Cell[] { Cell.FromText("r5") },
                },
                RowHeights = new() { { 0, 20.0 }, { 2, 40.0 } }, // 行3=20, 行5=40
            };
            XlsxWriter.Write(file, sd);

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            Assert.NotNull(ws.RowHeights);
            // 绝对 0-based：row3 -> 2, row5 -> 4
            Assert.Equal(20.0, ws.RowHeights![2], 3);
            Assert.Equal(40.0, ws.RowHeights[4], 3);

            rb.SaveAs(file2);
            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"3\" ht=\"20\"", saved);
            Assert.Contains("<row r=\"5\" ht=\"40\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    /// <summary>
    /// [Safe] 内部稀疏多行高度：数据行 2,4,7，row2 与 row7 设不同行高。
    /// 修复后：RowNumbers 精确映射，各高度落到正确绝对行（1 和 6），互不干扰。
    /// </summary>
    [Fact]
    public void C1_RowHeights_InteriorGap_MultipleRows_MapToCorrectAbsoluteRows()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"2\" ht=\"11\" customHeight=\"1\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>" +
                "<row r=\"4\"><c r=\"A4\" t=\"s\"><v>2</v></c></row>" +
                "<row r=\"7\" ht=\"77\" customHeight=\"1\"><c r=\"A7\" t=\"s\"><v>5</v></c></row>" +
                "</sheetData>");

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            Assert.NotNull(ws.RowHeights);
            Assert.Equal(11.0, ws.RowHeights![1], 3); // 绝对 0-based：row2 -> 1
            Assert.Equal(77.0, ws.RowHeights[6], 3);  // 绝对 0-based：row7 -> 6
            Assert.False(ws.RowHeights.ContainsKey(3), "空行 row4 不应有高度");

            rb.SaveAs(file2);
            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"2\" ht=\"11\"", saved);
            Assert.Contains("<row r=\"7\" ht=\"77\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    /// <summary>
    /// [Safe] 连续数据（行 1,2,3）内隐藏行 row2：RowNumbers 映射与旧的连续语义一致，
    /// 绝对索引 = 行号 - 1，且 Filter 行为不受本次修复影响。
    /// </summary>
    [Fact]
    public void C1_Filter_HiddenRows_ConsecutiveRows_Unchanged()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>1</v></c></row>" +
                "<row r=\"2\" hidden=\"1\"><c r=\"A2\" t=\"s\"><v>2</v></c></row>" +
                "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>3</v></c></row>" +
                "</sheetData>" +
                "<autoFilter ref=\"A1:A3\"></autoFilter>");

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            Assert.NotNull(ws.Filter);
            Assert.Contains(1, ws.Filter!.HiddenRows); // row2 -> 绝对索引 1

            rb.SaveAs(file2);
            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"2\" hidden=\"1\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    // ─────────────────────────────────────────────────────────────────────
    // 必须测试 2：RowStyles 对象模型往返
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// [Existing Limitation] 中间存在空行时，Filter.HiddenRows 丢失（reader 以数据相对索引存储，
    /// 而 writer/对象模型按行索引消费）。修复前同样丢失（紧凑模型下索引不匹配），故非本次 C1 引入。
    /// 本测试固定当前限制行为。
    /// </summary>
    [Fact]
    public void C1_Filter_HiddenRows_InteriorGap_DocumentedLimitation()
    {
        var file = Tmp();
        var file2 = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>0</v></c></row>" +
                "<row r=\"4\"><c r=\"A4\" t=\"s\"><v>2</v></c></row>" +
                "<row r=\"7\" hidden=\"1\"><c r=\"A7\" t=\"s\"><v>5</v></c></row>" +
                "</sheetData>" +
                "<autoFilter ref=\"A2:A7\"></autoFilter>");

            var rb = Excel.Open(file);
            var ws = rb.Worksheets["S"];
            Assert.NotNull(ws.Filter);
            rb.SaveAs(file2);

            var saved = ReadEntry(file2, "xl/worksheets/sheet1.xml");
            // 当前限制：内部空行场景隐藏标记未被保留（row7 不再 hidden）
            Assert.DoesNotContain("r=\"7\" hidden=\"1\"", saved);
        }
        finally { if (File.Exists(file)) File.Delete(file); if (File.Exists(file2)) File.Delete(file2); }
    }

    /// <summary>
    /// [Safe] 对象模型 RowStyles 写出侧：key 为绝对行索引（行号-1），样式应用到正确绝对行。
    /// </summary>
    [Fact]
    public void C1_RowStyles_WriteSide_AbsoluteRowConsistent()
    {
        var file = Tmp();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            for (int r = 1; r <= 5; r++)
                ws.SetValue(r, 1, $"v{r}");
            // 绝对行索引 3 -> 绝对第 4 行
            ws.RowStyles = new() { { 3, new CellStyle { Bold = true } } };
            wb.SaveAs(file);

            var saved = ReadEntry(file, "xl/worksheets/sheet1.xml");
            Assert.Contains("<c r=\"A4\" s=", saved);      // 行4 应用了样式
            Assert.DoesNotContain("<c r=\"A3\" s=", saved); // 行3 未应用
            Assert.DoesNotContain("<c r=\"A5\" s=", saved); // 行5 未应用
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>
    /// [Existing Limitation] XlsxReader 从不恢复 sheet.RowStyles（行级样式被展开到单元格 xf），
    /// 因此对象模型 Open 后 Worksheet.RowStyles 恒为 null。并非本次 C1 修改造成。
    /// </summary>
    [Fact]
    public void C1_RowStyles_NotRestored_OnObjectModelOpen()
    {
        var file = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"3\" s=\"1\" customFormat=\"1\"><c r=\"A3\" t=\"s\"><v>1</v></c></row>" +
                "<row r=\"4\" s=\"2\" customFormat=\"1\"><c r=\"A4\" t=\"s\"><v>2</v></c></row>" +
                "</sheetData>",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<styleSheet xmlns=\"{MainNs}\"><cellXfs count=\"3\">" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"center\"/></xf>" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"><alignment horizontal=\"right\"/></xf>" +
                "</cellXfs></styleSheet>");

            var ws = Excel.Open(file).Worksheets["S"];
            // 行级样式不会恢复到 Worksheet.RowStyles（既有行为，非本次回归）
            Assert.Null(ws.RowStyles);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>
    /// [Design Limitation · API 契约] 对象模型 RowStyles 的完整往返契约：
    /// Save 有效（写出时行级样式已展开到该行各单元格的 xf），但 Open 后**不恢复** <see cref="Worksheet.RowStyles"/>
    /// ——Reader 把行级样式效果展开为 <c>Cell.Style</c>，不重建 sheet 级 RowStyles。
    /// 因此目标绝对行的单元格样式效果保留，而行级抽象恢复为 null。这是当前正式设计行为，非 Reader Bug。
    /// 说明：本用例只在 RowStyles 上设样式、**不设** Cell.Style，以隔离样式来源。
    /// </summary>
    [Fact]
    public void C1_RowStyles_ObjectModelRoundTrip_PreservesCellStyle_ButNotRowStyle()
    {
        var file = Tmp();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            for (int r = 1; r <= 5; r++)
                ws.SetValue(r, 1, $"v{r}");

            // 仅通过 RowStyles 施加样式：绝对行索引 3 -> 绝对第 4 行。
            ws.RowStyles = new() { { 3, new CellStyle { FillColor = "#FF0000", Bold = true } } };
            Assert.Null(ws.Cell("A4").Style);   // 单元格本身没有显式 Cell.Style

            wb.SaveAs(file);

            var reopened = Excel.Open(file).Worksheets["S"];
            // 断言 A：行级抽象不在 Open 后恢复（正式设计契约）。
            Assert.Null(reopened.RowStyles);
            // 断言 B：RowStyle 的视觉样式效果经 Writer 展开到单元格、由 Reader 恢复为 Cell.Style。
            Assert.Equal("#FF0000", reopened.Cell("A4").Style?.FillColor);
            Assert.True(reopened.Cell("A4").Style?.Bold);
            // 其他行未受影响（证明是行级、非表级效果）。
            Assert.Null(reopened.Cell("A3").Style);
            Assert.Null(reopened.Cell("A5").Style);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ─────────────────────────────────────────────────────────────────────
    // 必须测试 3：Worksheet.FirstRowNumber setter
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// [API 契约 · 最终] 对象模型保存路径不使用 <c>FirstRowNumber</c> 进行行位移（_grid 已按绝对行号稠密布局，
    /// ToSheetData 固定以第 1 行起导出）。setter 已标记 <see cref="ObsoleteAttribute"/>（error:false）；
    /// getter 仍反映打开文件的原始首行号。本测试锁定该最终契约：新建簿 → 赋值 setter → 数据仍写在第 1 行。
    /// </summary>
    [Fact]
    public void C1_FirstRowNumberSetter_IgnoredOnObjectModelSave_BehaviorChange()
    {
        var file = Tmp();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
#pragma warning disable CS0618 // 故意调用已废弃的 setter 以锁定最终 API 契约。
            ws.FirstRowNumber = 5;
#pragma warning restore CS0618
            ws.SetValue("A1", "x");
            Assert.Equal(5, ws.FirstRowNumber);                    // getter 保留赋值（属性本身可暂时保存）
            Assert.Equal(1, ws.ToSheetData().FirstRowNumber);      // 写出值固定为 1（保存不位移）
            wb.SaveAs(file);

            var saved = ReadEntry(file, "xl/worksheets/sheet1.xml");
            Assert.Contains("<row r=\"1\">", saved);               // 数据被写在第 1 行（绝对行号决定的落点）
            Assert.DoesNotContain("r=\"5\"", saved);

            var reopened = Excel.Open(file).Worksheets["S"];
            Assert.Equal(1, reopened.FirstRowNumber);              // 重新打开后 getter 反映实际文件首行号 1
            Assert.Equal("x", reopened.Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    // ─────────────────────────────────────────────────────────────────────
    // 必须测试 4/5：C3 启发式误判 + 对象模型 HeaderStyle 限制
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// [Heuristic false positive] 未设置 DefaultStyle，但所有数据单元格恰好使用同一样式 A；
    /// RecoverSheetLevelStyles 会把 A 推断为 DefaultStyle。固定该启发式契约（非严格 round-trip）。
    /// </summary>
    [Fact]
    public void C3_DefaultStyle_HeuristicFalsePositive()
    {
        var file = Tmp();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            var s = new CellStyle { Bold = true, FontColor = "#123456" };
            ws.SetValue("A1", "h"); ws.Cell("A1").Style = s;
            ws.SetValue("A2", "x"); ws.Cell("A2").Style = s;
            ws.SetValue("B2", "y"); ws.Cell("B2").Style = s;
            Assert.Null(ws.DefaultStyle); // 用户并未设置 sheet 级 DefaultStyle
            wb.SaveAs(file);

            var r = Excel.Open(file).Worksheets["S"];
            // 当前实现：所有单元格样式一致 -> 被推断为 DefaultStyle（启发式，可能误判）
            Assert.NotNull(r.DefaultStyle);
            Assert.True(r.DefaultStyle!.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>
    /// [Heuristic false positive] 未设置 HeaderStyle，但表头行所有单元格恰好使用同一样式 A
    /// （且与数据区默认不同）；低层读取时会把 A 推断为 HeaderStyle。固定该启发式契约。
    /// </summary>
    [Fact]
    public void C3_HeaderStyle_HeuristicFalsePositive()
    {
        var file = Tmp();
        try
        {
            BuildRaw(file,
                "<sheetData>" +
                "<row r=\"1\"><c r=\"A1\" s=\"1\" t=\"s\"><v>0</v></c><c r=\"B1\" s=\"1\" t=\"s\"><v>1</v></c></row>" +
                "<row r=\"2\"><c r=\"A2\" s=\"2\" t=\"s\"><v>2</v></c></row>" +
                "</sheetData>",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<styleSheet xmlns=\"{MainNs}\"><fonts count=\"3\">" +
                "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
                "<font><sz val=\"11\"/><name val=\"Calibri\"/><b/><color rgb=\"FFFF0000\"/></font>" +
                "<font><sz val=\"11\"/><name val=\"Calibri\"/><i/></font>" +
                "</fonts><cellXfs count=\"3\">" +
                "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
                "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"1\"/>" +
                "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" applyFont=\"1\"/>" +
                "</cellXfs></styleSheet>");

            // 低层读取（firstRowIsHeader=true）：首行进入 Headers；未显式设 HeaderStyle
            var read = XlsxReader.Read(file, 0);
            // 当前实现：表头单元格式一致且与数据默认不同 -> 被推断为 HeaderStyle（启发式）
            Assert.NotNull(read.HeaderStyle);
            Assert.True(read.HeaderStyle!.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    /// <summary>
    /// [Design Limitation] 普通对象模型路径（Excel.Open → ReadAllRaw，firstRowIsHeader=false）没有
    /// Headers 行概念；写出时 Worksheet.HeaderStyle 不会落到任何可逆的文件级单元格样式，
    /// 因此保存文件无法恢复 HeaderStyle。本测试固定当前限制。
    /// </summary>
    [Fact]
    public void C3_ObjectModel_HeaderStyle_NotAppliedOnWrite_DesignLimitation()
    {
        var file = Tmp();
        try
        {
            var wb = Excel.Create("S");
            var ws = wb.Worksheets["S"];
            ws.SetValue("A1", "Header");
            ws.SetValue("A2", "x");
            ws.HeaderStyle = new CellStyle { Bold = true, FillColor = "#FF0000" };
            wb.SaveAs(file);

            // 写出侧：对象模型 HeaderStyle 未作用于任何单元格（A1 单元格无 s= 样式引用）
            var saved = ReadEntry(file, "xl/worksheets/sheet1.xml");
            Assert.DoesNotContain("<c r=\"A1\" s=", saved);

            var r = Excel.Open(file).Worksheets["S"];
            Assert.Null(r.HeaderStyle); // 无法从保存文件恢复
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
