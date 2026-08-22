using LiteExcel.Internal;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace LiteExcel;
public static partial class XlsxWriter
{
    internal const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>写出单表 </summary>
    public static void Write(string path, SheetData data)
    {
        Write(path, new[] { data }, null);
    }

    /// <summary>写出单表，并携带文档属性 </summary>
    public static void Write(string path, SheetData data, WorkbookProperties? properties)
    {
        Write(path, new[] { data }, properties);
    }

    /// <summary>写出多表 </summary>
    public static void Write(string path, IReadOnlyList<SheetData> sheets)
    {
        Write(path, sheets, null);
    }

    /// <summary>写出多表，并携带文档属性。目标为 .xlsm 时写出 macroEnabled 主文档类型 </summary>
    public static void Write(string path, IReadOnlyList<SheetData> sheets, WorkbookProperties? properties)
    {
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        Write(fs, sheets, properties, null, true, macroEnabled: IsXlsmPath(path));
    }

    private static bool IsXlsmPath(string path) =>
        string.Equals(Path.GetExtension(path), ".xlsm", StringComparison.OrdinalIgnoreCase);

    /// <summary>写出单表到流 </summary>
    public static void Write(Stream stream, SheetData data)
    {
        Write(stream, new[] { data }, null);
    }

    /// <summary>写出单表到流，并携带文档属性 </summary>
    public static void Write(Stream stream, SheetData data, WorkbookProperties? properties)
    {
        Write(stream, new[] { data }, properties);
    }

    /// <summary>写出多表到流 </summary>
    public static void Write(Stream stream, IReadOnlyList<SheetData> sheets)
    {
        Write(stream, sheets, null);
    }

    /// <summary>写出多表到流，并携带文档属性 </summary>
    public static void Write(Stream stream, IReadOnlyList<SheetData> sheets, WorkbookProperties? properties)
    {
        Write(stream, sheets, properties, null, true);
    }

    /// <summary>
    /// 写出多表到流，并携带文档属性与打开时捕获的保留部件。
    /// <paramref name="preserved"/> 为打开工作簿时捕获的未重建 OOXML 部件（宏/主题/绘图等）；
    /// <paramref name="mergeSheetRels"/> 为 false 时丢弃工作表级保留 rels（工作表结构已变化）。
    /// </summary>
    internal static void Write(Stream stream, IReadOnlyList<SheetData> sheets, WorkbookProperties? properties,
        OoxmlPreservedParts? preserved, bool mergeSheetRels, bool macroEnabled = false, bool date1904 = false,
        string? fileSharingHash = null, string? fileSharingSalt = null, int? fileSharingSpin = null,
        bool fileSharingReadOnlyRecommended = false)
    {
        if (sheets is null || sheets.Count == 0)
            throw new ArgumentException("至少需要一张工作表", nameof(sheets));

        // 0. Sheet 名校验（入口拦截，不影响写出逻辑）
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in sheets)
        {
            ValidateSheetName(sheet?.SheetName);
            // P0-13: 低层 API 也校验重复表名，与高层 WorksheetCollection.Add 行为一致
            if (!seenNames.Add(sheet!.SheetName!))
                throw new LiteExcelException($"工作表名重复：{sheet.SheetName}");
        }

        // 0. Sheet 名校验（入口拦截，不影响写出逻辑）
        foreach (var sheet in sheets)
        {
            ValidateSheetName(sheet?.SheetName);
        }

        //   收集共享字符串和样式（跨所有表）
        var sharedStrings = new List<string>();
        var sharedIndex = new Dictionary<string, int>();
        var stylesheet = new Stylesheet();

        // 预扫描：注册所有字符串和样式
        foreach (var sheet in sheets)
        {
            if (sheet.Headers is not null)
            {
                foreach (var h in sheet.Headers)
                    RegisterSharedString(h, sharedStrings, sharedIndex);
            }
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row)
                {
                    if (cell.Type == CellType.Text && !cell.IsFormula && cell.Text is not null)
                        RegisterSharedString(cell.Text, sharedStrings, sharedIndex);
                }
            }
        }

        //   构建 zip
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);

        // 预计算哪些 sheet 有批注
        var sheetsWithComments = new List<int>();
        for (int i = 0; i < sheets.Count; i++)
        {
            if (sheets[i].Comments is { Count: > 0 })
                sheetsWithComments.Add(i);
        }

        // 图片规划：分配 media 序号、生成 drawing/richData 部件
        var imagePlan = ImagePlan.Create(sheets, preserved);

        // 先写保留部件（blob），再写重建部件，避免重名时重建优先
        // 浮动图片 drawing 与 InCell richData 部件会在后续整体合并/新建，若与保留部件重名则跳过保留（避免 zip 重名异常，P0-11）
        if (preserved is not null)
        {
            var imageEntries = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
            foreach (var (entry, _) in imagePlan.FloatingDrawingParts(preserved))
                imageEntries.Add(entry);
            foreach (var (entry, _) in imagePlan.InCellEntries())
                imageEntries.Add(entry);
            foreach (var kv in preserved.Parts)
            {
                if (imageEntries.Contains(kv.Key)) continue;
                WriteEntry(zip, kv.Key, kv.Value);
            }
        }

        WriteXmlEntry(zip, "[Content_Types].xml", ContentTypesXml(sheets.Count, sheetsWithComments, properties is not null, preserved, macroEnabled, imagePlan));
        WriteXmlEntry(zip, "_rels/.rels", RootRelsXml(properties is not null, preserved));
        WriteXmlEntry(zip, "xl/workbook.xml", WorkbookXml(sheets, preserved, date1904, fileSharingHash, fileSharingSalt, fileSharingSpin, fileSharingReadOnlyRecommended));
        WriteXmlEntry(zip, "xl/_rels/workbook.xml.rels", WorkbookRelsXml(sheets.Count, preserved, imagePlan));

        // 图片 media
        foreach (var (entry, bytes) in imagePlan.MediaEntries())
            WriteEntry(zip, entry, bytes);

        // 浮动图片 drawing 部件（含既有 drawing 合并 + rels）
        foreach (var (entry, xml) in imagePlan.FloatingDrawingParts(preserved))
            WriteEntry(zip, entry, xml);

        // InCell richData 部件
        foreach (var (entry, xml) in imagePlan.InCellEntries())
            WriteXmlEntry(zip, entry, xml);

        // 文档属性（文件属性对话框信息）
        if (properties is not null)
        {
            WriteXmlEntry(zip, "docProps/core.xml", CorePropsXml(properties));
            WriteXmlEntry(zip, "docProps/app.xml", AppPropsXml(properties, sheets));
        }

        for (int i = 0; i < sheets.Count; i++)
        {
            var hyperlinks = new List<(string Ref, string Target, string? Tooltip, bool IsInternal)>();
            var inCellVm = imagePlan.InCellVmBySheet(i);
            bool hasDrawing = imagePlan.FloatingBySheet[i].Count > 0;
            string drawingRelId = hasDrawing ? imagePlan.DrawingTargetFor(i, preserved).RelId : "";
            var sheetXml = BuildSheetXml(sheets[i], sharedIndex, stylesheet, date1904, hyperlinks, inCellVm, hasDrawing, drawingRelId);
            WriteXmlEntry(zip, $"xl/worksheets/sheet{i + 1}.xml", sheetXml);

            // 批注：每张有批注的 sheet 对应一个 comments 文件
            bool hasComments = sheets[i].Comments is { Count: > 0 };
            if (hasComments)
            {
                WriteXmlEntry(zip, $"xl/comments{i + 1}.xml", CommentsXml(sheets[i].Comments!));
            }

            // 工作表 rels：合并保留的绘图/超链接等 rel（工作表结构未变时），追加新建超链接/批注/drawing
            var sheetRels = MergeSheetRels(i + 1, hasComments, preserved, mergeSheetRels, hyperlinks, hasDrawing, imagePlan);
            if (sheetRels is not null)
            {
                WriteXmlEntry(zip, $"xl/worksheets/_rels/sheet{i + 1}.xml.rels", sheetRels);
            }
        }

        WriteXmlEntry(zip, "xl/sharedStrings.xml", SharedStringsXml(sharedStrings));
        WriteXmlEntry(zip, "xl/styles.xml", stylesheet.BuildStylesXml());
    }

    /// <summary>
    /// 追加数据到已有文件 同名 sheet 合并列后追加行；不同名则作为新 sheet 加入 
    /// 文件不存在时直接创建 
    /// </summary>
    public static void Append(string path, SheetData? newData, WorkbookProperties? updateProperties = null)
    {
        if (newData is null || newData.Rows is null || newData.Rows.Count == 0)
            return;

        // P0-24: 仅 xlsx/xlsm 支持追加，其他格式显式报错而非误导的 zip 解析异常
        var format = Excel.DetectFormat(path);
        if (format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsm)
            throw new LiteExcelException($"该格式不支持追加：{format}。仅支持 xlsx/xlsm。");

        if (!File.Exists(path))
        {
            Write(path, newData, updateProperties);
            return;
        }

        var allSheets = XlsxReader.ReadAll(path);
        var properties = XlsxReader.ReadProperties(path);

        // P0-25: 捕获保留部件（宏/主题/绘图/图表/表格等），追加时透传，避免 xlsm 丢宏、xlsx 丢图表
        OoxmlPreservedParts preserved;
        using (var readFs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(readFs, ZipArchiveMode.Read))
        {
            preserved = OoxmlPreservedParts.Capture(zip, allSheets.Count);
            preserved.WorkbookCodeName = XlsxReader.WorkbookCodeNameSnapshot;
        }

        ApplyPropertyUpdates(properties, updateProperties);
        // Preserve existing metadata while recording this append operation.
        properties.Modified = DateTime.Now;
        var name = newData.SheetName ?? "";
        int idx = allSheets.FindIndex(s => s.SheetName == name);

        if (idx >= 0)
        {
            // 同名 sheet：合并列头，追加行
            AppendToSheet(allSheets[idx], newData);
        }
        else
        {
            allSheets.Add(newData);
        }

        // 追加不改变既有 sheet 顺序，工作表级保留 rels 可继续复用
        using var outFs = new FileStream(path, FileMode.Create, FileAccess.Write);
        Write(outFs, allSheets, properties, preserved, true, macroEnabled: IsXlsmPath(path));
    }

    private static void ApplyPropertyUpdates(WorkbookProperties target, WorkbookProperties? updates)
    {
        if (updates is null) return;

        if (updates.Creator is not null) target.Creator = updates.Creator;
        if (updates.LastModifiedBy is not null) target.LastModifiedBy = updates.LastModifiedBy;
        if (updates.Title is not null) target.Title = updates.Title;
        if (updates.Subject is not null) target.Subject = updates.Subject;
        if (updates.Application is not null) target.Application = updates.Application;
        if (updates.Created is not null) target.Created = updates.Created;
    }

    private static void AppendToSheet(SheetData existing, SheetData newData)
    {
        //   合并 Headers
        var mergedHeaders = new List<string>(existing.Headers);
        if (newData.Headers is not null)
        {
            foreach (var h in newData.Headers)
            {
                if (!mergedHeaders.Contains(h))
                    mergedHeaders.Add(h);
            }
        }
        existing.Headers = mergedHeaders;

        //   构建列名映射
        var newHeaderMap = new Dictionary<string, int>();
        if (newData.Headers is not null)
        {
            for (int i = 0; i < newData.Headers.Count; i++)
            {
                if (!newHeaderMap.ContainsKey(newData.Headers[i]))
                    newHeaderMap[newData.Headers[i]] = i;
            }
        }

        int colCount = mergedHeaders.Count;

        //   追加行
        foreach (var row in newData.Rows)
        {
            var padded = new List<Cell>(colCount);
            for (int c = 0; c < colCount; c++)
            {
                string colName = mergedHeaders[c];
                if (newHeaderMap.TryGetValue(colName, out int srcIdx) && srcIdx < row.Count)
                {
                    padded.Add(row[srcIdx]);
                }
                else
                {
                    padded.Add(Cell.Empty);
                }
            }
            existing.Rows.Add(padded);
        }
    }

    // ── Sheet 名校验 ──

    private static readonly char[] InvalidSheetNameChars = new[] { '\\', '/', '?', '*', '[', ']', ':' };

    private static void ValidateSheetName(string? sheetName)
    {
        bool invalid = string.IsNullOrEmpty(sheetName)
            || sheetName!.Length > 31
            || sheetName.IndexOfAny(InvalidSheetNameChars) >= 0;

        if (invalid)
        {
            throw new InvalidSheetNameException(
                sheetName ?? "",
                "Sheet 名不能为空/超过 31 字符/包含 \\ / ? * [ ] : 字符");
        }
    }

    // ── 共享字符串管理 ──

    private static int RegisterSharedString(string? s, List<string> shared, Dictionary<string, int> index)
    {
        if (string.IsNullOrEmpty(s)) return -1;
        if (index.TryGetValue(s, out var i)) return i;
        i = shared.Count;
        shared.Add(s);
        index[s] = i;
        return i;
    }

    // ── 工作表 XML 构建 ──

    private static string BuildSheetXml(SheetData sheet,
        Dictionary<string, int> sharedIndex, Stylesheet stylesheet, bool date1904,
        List<(string Ref, string Target, string? Tooltip, bool IsInternal)>? hyperlinks = null,
        Dictionary<string, int>? inCellVm = null, bool hasDrawing = false, string drawingRelId = "rIdD1")
    {
        var sb = new StringBuilder(4096);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<worksheet xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\"");
        sb.Append(">");

        // 工作表宿主 VBA 代码名：schema 要求 sheetPr 为 worksheet 第一个子元素，位于 sheetViews 之前
        if (!string.IsNullOrEmpty(sheet.CodeName))
            sb.Append($"<sheetPr codeName=\"{XmlEscape(sheet.CodeName)}\"/>");

        // sheetView（冻结窗格）
        int freezeRows = sheet.FreezeRows;
        int freezeCols = sheet.FreezeColumns;
        if (sheet.FreezeHeader) freezeRows = Math.Max(freezeRows, 1);
        if (freezeRows > 0 || freezeCols > 0)
        {
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\">");
            sb.Append($"<pane");
            if (freezeRows > 0) sb.Append($" ySplit=\"{freezeRows}\"");
            if (freezeCols > 0) sb.Append($" xSplit=\"{freezeCols}\"");
            if (freezeRows > 0 && freezeCols > 0)
                sb.Append($" topLeftCell=\"{CellRef.ToString(freezeRows, freezeCols)}\" activePane=\"bottomRight\"");
            else if (freezeRows > 0)
                sb.Append($" topLeftCell=\"{CellRef.ToString(freezeRows, 0)}\" activePane=\"bottomLeft\"");
            else
                sb.Append($" topLeftCell=\"{CellRef.ToString(0, freezeCols)}\" activePane=\"topRight\"");
            sb.Append(" state=\"frozen\"/>");
            sb.Append("</sheetView></sheetViews>");
        }
        else
        {
            sb.Append("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>");
        }

        // 列宽
        if (sheet.ColumnWidths is { Count: > 0 })
        {
            sb.Append("<cols>");
            for (int i = 0; i < sheet.ColumnWidths.Count; i++)
            {
                sb.Append($"<col min=\"{i + 1}\" max=\"{i + 1}\" width=\"{FormatDouble(sheet.ColumnWidths[i])}\" customWidth=\"1\"/>");
            }
            sb.Append("</cols>");
        }

        sb.Append("<sheetData>");

        int rowIndex = 1;
        int headerStyleId = stylesheet.GetOrCreateXfId(sheet.HeaderStyle);

        // 表头行
        if (sheet.Headers is { Count: > 0 })
        {
            int headerRow = rowIndex++;
            sb.Append($"<row r=\"{headerRow}\">");
            for (int col = 0; col < sheet.Headers.Count; col++)
            {
                // 表头样式优先级: HeaderStyle > ColumnStyles > DefaultStyle
                var headerCellStyle = sheet.HeaderStyle
                    ?? (sheet.ColumnStyles is not null && sheet.ColumnStyles.TryGetValue(col, out var cs) ? cs : null)
                    ?? sheet.DefaultStyle;
                int headerCellStyleId = stylesheet.GetOrCreateXfId(headerCellStyle);
                WriteTextCell(sb, headerRow, col, sheet.Headers[col], sharedIndex, headerCellStyleId);
            }
            sb.Append("</row>");
        }

        // 计算筛选 hidden 行（如果有筛选条件但没手动设 HiddenRows）
        var hiddenRows = sheet.Filter?.HiddenRows;
        if (sheet.Filter is not null && sheet.Filter.Columns.Count > 0 && hiddenRows is not null && hiddenRows.Count == 0)
        {
            hiddenRows = FilterEvaluator.EvaluateHiddenRows(sheet);
        }

        // 数据行
        int dataRowIdx = 0;
        foreach (var row in sheet.Rows)
        {
            int currentRow = rowIndex++;
            int maxCol = row.Count - 1;
            if (maxCol < 0) { dataRowIdx++; continue; }

            // 行级样式
            CellStyle? rowStyle = null;
            if (sheet.RowStyles is not null && sheet.RowStyles.TryGetValue(dataRowIdx, out var rs))
                rowStyle = rs;

            bool isHidden = hiddenRows is not null && hiddenRows.Contains(dataRowIdx);
            string hiddenAttr = isHidden ? " hidden=\"1\"" : "";
            string heightAttr = "";
            if (sheet.RowHeights is not null && sheet.RowHeights.TryGetValue(dataRowIdx, out var ht))
            {
                heightAttr = $" ht=\"{FormatDouble(ht)}\" customHeight=\"1\"";
            }
            sb.Append($"<row r=\"{currentRow}\"{hiddenAttr}{heightAttr}>");
            for (int col = 0; col <= maxCol; col++)
            {
                var cell = col < row.Count ? row[col] : Cell.Empty;

                // 样式优先级: Cell.Style > RowStyle > ColumnStyle > DefaultStyle
                CellStyle? resolvedStyle = cell.Style;
                if (resolvedStyle is null)
                {
                    resolvedStyle = rowStyle;
                    if (resolvedStyle is null && sheet.ColumnStyles is not null && sheet.ColumnStyles.TryGetValue(col, out var cs))
                        resolvedStyle = cs;
                    if (resolvedStyle is null)
                        resolvedStyle = sheet.DefaultStyle;
                }

                if (hyperlinks is not null && cell.Hyperlink is not null)
                {
                    hyperlinks.Add((CellRef.ToString(currentRow - 1, col), cell.Hyperlink.Target, cell.Hyperlink.Tooltip, cell.Hyperlink.IsInternal));
                }

                // InCell 图片单元格：t="e" vm 指向 metadata 中的 richData 记录
                if (inCellVm is not null && inCellVm.TryGetValue(CellRef.ToString(currentRow - 1, col), out int vm))
                {
                    sb.Append($"<c r=\"{CellRef.ToString(currentRow - 1, col)}\" t=\"e\" vm=\"{vm}\"><v>#VALUE!</v></c>");
                    continue;
                }

                WriteCell(sb, currentRow, col, cell, sharedIndex, stylesheet, resolvedStyle, date1904);
            }
            sb.Append("</row>");
            dataRowIdx++;
        }

        // 补充 InCell 图片单元格：所在行可能不在数据网格中（如空行或超出网格），需按行号补齐
        if (inCellVm is { Count: > 0 })
        {
            int lastRow = rowIndex - 1;
            var rowGroups = new SortedDictionary<int, List<(int Col, int Vm)>>();
            foreach (var kv in inCellVm)
            {
                var (r, c) = CellRef.Parse(kv.Key);
                int row1 = r + 1, col1 = c + 1;
                if (row1 > lastRow)
                {
                    if (!rowGroups.TryGetValue(row1, out var list))
                    {
                        list = new List<(int, int)>();
                        rowGroups[row1] = list;
                    }
                    list.Add((col1, kv.Value));
                }
            }
            foreach (var grp in rowGroups)
            {
                for (int r = lastRow + 1; r < grp.Key; r++)
                    sb.Append($"<row r=\"{r}\"/>");
                sb.Append($"<row r=\"{grp.Key}\">");
                foreach (var (col1, vm) in grp.Value)
                    sb.Append($"<c r=\"{CellRef.ToString(grp.Key - 1, col1 - 1)}\" t=\"e\" vm=\"{vm}\"><v>#VALUE!</v></c>");
                sb.Append("</row>");
                lastRow = grp.Key;
            }
        }

        sb.Append("</sheetData>");

        // 超链接（外部 r:id 与 sheet rels 编号对应，从 rIdH1 起；内部链接用 location）
        if (hyperlinks is { Count: > 0 })
        {
            sb.Append("<hyperlinks>");
            int extIndex = 0;
            for (int h = 0; h < hyperlinks.Count; h++)
            {
                var (ref_, target, tooltip, isInternal) = hyperlinks[h];
                string tooltipAttr = string.IsNullOrEmpty(tooltip) ? "" : $" tooltip=\"{XmlEscape(tooltip)}\"";
                if (isInternal)
                {
                    sb.Append($"<hyperlink ref=\"{ref_}\" location=\"{XmlEscape(NormalizeInternalLocation(target))}\"{tooltipAttr}/>");
                }
                else
                {
                    extIndex++;
                    sb.Append($"<hyperlink ref=\"{ref_}\" r:id=\"rIdH{extIndex}\"{tooltipAttr}/>");
                }
            }
            sb.Append("</hyperlinks>");
        }

        // 数据验证
        if (sheet.Validations is { Count: > 0 })
        {
            sb.Append($"<dataValidations count=\"{sheet.Validations.Count}\">");
            foreach (var dv in sheet.Validations)
            {
                string typeAttr = dv.Type switch
                {
                    DataValidationType.List => "list",
                    DataValidationType.WholeNumber => "whole",
                    DataValidationType.Decimal => "decimal",
                    DataValidationType.Date => "date",
                    _ => "list",
                };
                string allowBlankAttr = dv.AllowBlank ? " allowBlank=\"1\"" : "";
                string promptTitleAttr = dv.PromptTitle is not null ? $" promptTitle=\"{XmlEscape(dv.PromptTitle)}\"" : "";
                string promptAttr = dv.Prompt is not null ? $" prompt=\"{XmlEscape(dv.Prompt)}\"" : "";

                sb.Append($"<dataValidation type=\"{typeAttr}\" sqref=\"{XmlEscape(dv.Sqref)}\"{allowBlankAttr}{promptTitleAttr}{promptAttr}>");
                sb.Append($"<formula1>{XmlEscape(dv.Formula1)}</formula1>");
                if (dv.Formula2 is not null)
                {
                    sb.Append($"<formula2>{XmlEscape(dv.Formula2)}</formula2>");
                }
                sb.Append("</dataValidation>");
            }
            sb.Append("</dataValidations>");
        }

        // 自动筛选
        if (sheet.Filter is not null)
        {
            string filterRange = sheet.Filter.Range;
            if (string.IsNullOrEmpty(filterRange) && sheet.Headers.Count > 0)
            {
                // 自动计算范围：表头行到最后一数据行
                int totalRows = sheet.Headers.Count + sheet.Rows.Count;
                filterRange = $"{CellRef.ToString(0, 0)}:{CellRef.ToString(totalRows - 1, sheet.Headers.Count - 1)}";
            }

            if (!string.IsNullOrEmpty(filterRange))
            {
                sb.Append($"<autoFilter ref=\"{filterRange}\">");
                foreach (var col in sheet.Filter.Columns)
                {
                    sb.Append($"<filterColumn colId=\"{col.ColumnIndex}\">");
                    sb.Append(BuildFilterColumnXml(col));
                    sb.Append("</filterColumn>");
                }
                sb.Append("</autoFilter>");
            }
        }

        // 合并单元格
        if (sheet.MergedRanges is { Count: > 0 })
        {
            // 表头占 1 行（如果有 Headers），MergedRanges 行号是相对于 Rows 的，需要偏移
            int headerOffset = (sheet.Headers is { Count: > 0 }) ? 1 : 0;
            sb.Append("<mergeCells count=\"" + sheet.MergedRanges.Count + "\">");
            foreach (var range in sheet.MergedRanges)
            {
                var from = CellRef.ToString(range.FirstRow + headerOffset, range.FirstCol);
                var to = CellRef.ToString(range.LastRow + headerOffset, range.LastCol);
                sb.Append($"<mergeCell ref=\"{from}:{to}\"/>");
            }
            sb.Append("</mergeCells>");
        }

        // 条件格式（2.4.3：cellIs/expression/colorScale/dataBar）
        if (sheet.ConditionalFormats is { Count: > 0 })
        {
            int priority = 1;
            foreach (var cf in sheet.ConditionalFormats)
            {
                if (string.IsNullOrEmpty(cf.Sqref)) continue;
                int dxfId = cf.Style is not null ? stylesheet.GetOrCreateDxfId(cf.Style) : -1;
                string dxfAttr = dxfId >= 0 ? $" dxfId=\"{dxfId}\"" : "";
                int prio = cf.Priority > 0 ? cf.Priority : priority++;

                sb.Append($"<conditionalFormatting sqref=\"{XmlEscape(cf.Sqref)}\">");
                switch (cf.Type)
                {
                    case ConditionalFormatType.CellIs:
                    {
                        sb.Append($"<cfRule type=\"cellIs\"{dxfAttr} priority=\"{prio}\" operator=\"{OperatorToString(cf.Operator)}\">");
                        AppendCfFormula(sb, cf.Formula);
                        if (cf.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
                            AppendCfFormula(sb, cf.Formula2);
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.Expression:
                    {
                        sb.Append($"<cfRule type=\"expression\"{dxfAttr} priority=\"{prio}\">");
                        AppendCfFormula(sb, cf.Formula);
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.ColorScale:
                    {
                        var cs = cf.ColorScale ?? new ColorScaleInfo();
                        sb.Append($"<cfRule type=\"colorScale\" priority=\"{prio}\">");
                        sb.Append("<colorScale>");
                        // cfvo 类型须为 ST_CfvoType 合法值：min/max/num/percent/percentile/formula
                        sb.Append(cs.MidColor is not null
                            ? "<cfvo type=\"min\"/><cfvo type=\"percent\" val=\"50\"/><cfvo type=\"max\"/>"
                            : "<cfvo type=\"min\"/><cfvo type=\"max\"/>");
                        if (cs.MidColor is null)
                        {
                            sb.Append($"<color rgb=\"FF{NormalizeColorRgb(cs.LowColor)}\"/><color rgb=\"FF{NormalizeColorRgb(cs.HighColor)}\"/>");
                        }
                        else
                        {
                            sb.Append($"<color rgb=\"FF{NormalizeColorRgb(cs.LowColor)}\"/><color rgb=\"FF{NormalizeColorRgb(cs.MidColor)}\"/><color rgb=\"FF{NormalizeColorRgb(cs.HighColor)}\"/>");
                        }
                        sb.Append("</colorScale>");
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.DataBar:
                    {
                        var db = cf.DataBar ?? new DataBarInfo();
                        sb.Append($"<cfRule type=\"dataBar\" priority=\"{prio}\">");
                        sb.Append($"<dataBar minLength=\"{db.MinLengthPercent}\" maxLength=\"{db.MaxLengthPercent}\" showValue=\"{(db.ShowValue ? 1 : 0)}\">");
                        // dataBar 的 cfvo 必须为 min/max（ST_CfvoType 无 "auto"，否则 Excel 报 XML 错误并丢弃规则）
                        sb.Append("<cfvo type=\"min\"/><cfvo type=\"max\"/>");
                        sb.Append($"<color rgb=\"FF{NormalizeColorRgb(db.Color)}\"/>");
                        sb.Append("</dataBar>");
                        sb.Append("</cfRule>");
                        break;
                    }
                    // ── 2.4.4 长尾类型（文本/时间周期/空值/错误/唯一重复/前N/平均线） ──
                    case ConditionalFormatType.ContainsText:
                    case ConditionalFormatType.BeginsWith:
                    case ConditionalFormatType.EndsWith:
                    case ConditionalFormatType.NotContainsText:
                    {
                        var text = XmlEscape(cf.Text ?? "");
                        sb.Append($"<cfRule type=\"{CfTypeToExcel(cf.Type)}\"{dxfAttr} priority=\"{prio}\" text=\"{text}\">");
                        AppendCfFormula(sb, TextCfFormula(cf));
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.TextLength:
                    {
                        sb.Append($"<cfRule type=\"lengthIs\"{dxfAttr} priority=\"{prio}\" operator=\"{OperatorToString(cf.Operator)}\">");
                        AppendCfFormula(sb, cf.Formula);
                        if (cf.Operator is ConditionalOperator.Between or ConditionalOperator.NotBetween)
                            AppendCfFormula(sb, cf.Formula2);
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.TimePeriod:
                    {
                        string tp = string.IsNullOrEmpty(cf.TimePeriod) ? "thisMonth" : cf.TimePeriod;
                        sb.Append($"<cfRule type=\"timePeriod\"{dxfAttr} priority=\"{prio}\" timePeriod=\"{tp}\">");
                        AppendCfFormula(sb, "TODAY()");
                        sb.Append("</cfRule>");
                        break;
                    }
                    case ConditionalFormatType.Blanks:
                    case ConditionalFormatType.NoBlanks:
                    case ConditionalFormatType.Errors:
                    case ConditionalFormatType.NoErrors:
                    case ConditionalFormatType.Unique:
                    case ConditionalFormatType.Duplicate:
                    {
                        // ST_CfType 合法值：uniqueValues/duplicateValues/containsBlanks/notContainsBlanks/containsErrors/notContainsErrors
                        string excelType = cf.Type switch
                        {
                            ConditionalFormatType.Blanks => "containsBlanks",
                            ConditionalFormatType.NoBlanks => "notContainsBlanks",
                            ConditionalFormatType.Errors => "containsErrors",
                            ConditionalFormatType.NoErrors => "notContainsErrors",
                            ConditionalFormatType.Unique => "uniqueValues",
                            _ => "duplicateValues",
                        };
                        sb.Append($"<cfRule type=\"{excelType}\"{dxfAttr} priority=\"{prio}\"/>");
                        break;
                    }
                    case ConditionalFormatType.Top10:
                    {
                        sb.Append($"<cfRule type=\"top10\"{dxfAttr} priority=\"{prio}\" rank=\"{cf.Rank}\" percent=\"{(cf.Percent ? 1 : 0)}\"/>");
                        break;
                    }
                    case ConditionalFormatType.AboveAverage:
                    case ConditionalFormatType.BelowAverage:
                    {
                        // 合格分两种：type="aboveAverage" + aboveAverage="1|0"（无 belowAverage 枚举）
                        string excelType = "aboveAverage";
                        sb.Append($"<cfRule type=\"{excelType}\"{dxfAttr} priority=\"{prio}\" aboveAverage=\"{(cf.Type == ConditionalFormatType.AboveAverage ? 1 : 0)}\"/>");
                        break;
                    }
                }
                sb.Append("</conditionalFormatting>");
            }
        }

        // 浮动图片 drawing 引用
        if (hasDrawing)
        {
            sb.Append($"<drawing r:id=\"{drawingRelId}\"/>");
        }

        sb.Append("</worksheet>");
        return sb.ToString();
    }

    private static string BuildFilterColumnXml(FilterColumn col)
    {
        var sb = new StringBuilder();
        switch (col.Type)
        {
            case FilterType.Equals:
                sb.Append("<filters>");
                foreach (var v in col.Values)
                    sb.Append($"<filter val=\"{XmlEscape(v)}\"/>");
                sb.Append("</filters>");
                break;

            case FilterType.Blank:
                if (col.Values.Count == 0)
                    sb.Append("<filters><blank/></filters>");
                else
                    sb.Append("<filters/>");
                break;

            case FilterType.Compare:
                if (col.Operator == FilterOperator.Between && col.MinValue is not null && col.MaxValue is not null)
                {
                    sb.Append("<customFilters and=\"1\">");
                    sb.Append($"<customFilter operator=\"greaterThanOrEqual\" val=\"{XmlEscape(col.MinValue)}\"/>");
                    sb.Append($"<customFilter operator=\"lessThanOrEqual\" val=\"{XmlEscape(col.MaxValue)}\"/>");
                    sb.Append("</customFilters>");
                }
                else
                {
                    sb.Append("<customFilters>");
                    for (int i = 0; i < col.Values.Count; i++)
                    {
                        string op = col.Operator switch
                        {
                            FilterOperator.GreaterThan => "greaterThan",
                            FilterOperator.GreaterThanOrEqual => "greaterThanOrEqual",
                            FilterOperator.LessThan => "lessThan",
                            FilterOperator.LessThanOrEqual => "lessThanOrEqual",
                            _ => "equal",
                        };
                        sb.Append($"<customFilter operator=\"{op}\" val=\"{XmlEscape(col.Values[i])}\"/>");
                    }
                    sb.Append("</customFilters>");
                }
                break;

            case FilterType.Contains:
                sb.Append("<customFilters>");
                foreach (var v in col.Values)
                    sb.Append($"<customFilter operator=\"equal\" val=\"*{XmlEscape(v)}*\"/>");
                sb.Append("</customFilters>");
                break;

            case FilterType.BeginsWith:
                sb.Append("<customFilters>");
                foreach (var v in col.Values)
                    sb.Append($"<customFilter operator=\"equal\" val=\"{XmlEscape(v)}*\"/>");
                sb.Append("</customFilters>");
                break;

            case FilterType.EndsWith:
                sb.Append("<customFilters>");
                foreach (var v in col.Values)
                    sb.Append($"<customFilter operator=\"equal\" val=\"*{XmlEscape(v)}\"/>");
                sb.Append("</customFilters>");
                break;
        }
        return sb.ToString();
    }

    // ── 条件格式辅助 ──

    private static void AppendCfFormula(StringBuilder sb, string? formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        sb.Append($"<formula>{XmlEscape(formula)}</formula>");
    }

    private static string OperatorToString(ConditionalOperator op) => op switch
    {
        ConditionalOperator.LessThan => "lessThan",
        ConditionalOperator.LessThanOrEqual => "lessThanOrEqual",
        ConditionalOperator.Equal => "equal",
        ConditionalOperator.NotEqual => "notEqual",
        ConditionalOperator.GreaterThan => "greaterThan",
        ConditionalOperator.GreaterThanOrEqual => "greaterThanOrEqual",
        ConditionalOperator.Between => "between",
        ConditionalOperator.NotBetween => "notBetween",
        _ => "greaterThan",
    };

    private static string NormalizeColorRgb(string color)
    {
        if (color.StartsWith("#")) return color.Substring(1).ToUpperInvariant();
        return color.ToUpperInvariant();
    }

    /// <summary>长尾文本类条件格式的 cfRule type 属性值 </summary>
    private static string CfTypeToExcel(ConditionalFormatType t) => t switch
    {
        ConditionalFormatType.ContainsText => "containsText",
        ConditionalFormatType.BeginsWith => "beginsWith",
        ConditionalFormatType.EndsWith => "endsWith",
        ConditionalFormatType.NotContainsText => "notContainsText",
        _ => "containsText",
    };

    /// <summary>长尾文本类条件格式的 <formula> 内容（Excel 约定，ref 指范围左上角） </summary>
    private static string TextCfFormula(ConditionalFormat cf)
    {
        var text = XmlEscape(cf.Text ?? "");
        // 若用户显式提供 Formula 则直接用；否则按类型生成标准 Excel 公式
        if (!string.IsNullOrEmpty(cf.Formula)) return cf.Formula;
        string refCell = FirstCellOfSqref(cf.Sqref);
        return cf.Type switch
        {
            ConditionalFormatType.ContainsText => $"NOT(ISERROR(SEARCH(\"{text}\",{refCell})))",
            ConditionalFormatType.BeginsWith => $"LEFT({refCell},LEN(\"{text}\"))=\"{text}\"",
            ConditionalFormatType.EndsWith => $"RIGHT({refCell},LEN(\"{text}\"))=\"{text}\"",
            ConditionalFormatType.NotContainsText => $"ISERROR(SEARCH(\"{text}\",{refCell}))",
            _ => $"NOT(ISERROR(SEARCH(\"{text}\",{refCell})))",
        };
    }

    /// <summary>从 sqref 取第一个单元格（如 "A1:A10" → "A1"），无则回退 "A1" </summary>
    private static string FirstCellOfSqref(string sqref)
    {
        var first = sqref.Split(' ', ';')[0];
        var colon = first.IndexOf(':');
        if (colon > 0) first = first.Substring(0, colon);
        return string.IsNullOrEmpty(first) ? "A1" : first;
    }

    private static void WriteTextCell(StringBuilder sb, int row1Based, int col, string text,
        Dictionary<string, int> sharedIndex, int styleId = 0)
    {
        var styleAttr = styleId > 0 ? $" s=\"{styleId}\"" : "";
        if (string.IsNullOrEmpty(text))
        {
            sb.Append($"<c r=\"{CellRef.ToString(row1Based - 1, col)}\"{styleAttr}/>");
            return;
        }
        var idx = sharedIndex.TryGetValue(text, out var i) ? i : -1;
        if (idx >= 0)
        {
            sb.Append($"<c r=\"{CellRef.ToString(row1Based - 1, col)}\"{styleAttr} t=\"s\"><v>{idx}</v></c>");
        }
        else
        {
            sb.Append($"<c r=\"{CellRef.ToString(row1Based - 1, col)}\"{styleAttr} t=\"inlineStr\"><is><t>{XmlEscape(text)}</t></is></c>");
        }
    }

    private static void WriteCell(StringBuilder sb, int row, int col, Cell cell,
        Dictionary<string, int> sharedIndex, Stylesheet stylesheet, CellStyle? resolvedStyle = null, bool date1904 = false)
    {
        var cellRef = CellRef.ToString(row - 1, col);
        // 使用解析后的样式（优先级已在外部处理），或 cell 自带的样式
        int styleId = stylesheet.GetOrCreateXfId(resolvedStyle ?? cell.Style, cell.NumberFormat);
        var styleAttr = styleId > 0 ? $" s=\"{styleId}\"" : "";

        // 公式单元格：写 <f> 公式文本 + <v> 缓存值（不做公式计算）
        // P0-8 兼容垫片：优先读 Cell.Formula；旧代码（IsFormula=true 且公式存于 Text）仍可用
        var formulaText = cell.Formula ?? (cell.IsFormula ? cell.Text : null);
        if (!string.IsNullOrEmpty(formulaText))
        {
            var fEsc = XmlEscape(formulaText);
            switch (cell.Type)
            {
                case CellType.Number:
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr}><f>{fEsc}</f><v>{FormatDouble(cell.Number)}</v></c>");
                    break;
                case CellType.Date:
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr}><f>{fEsc}</f><v>{FormatDouble(FormatDetector.DateToSerial(cell.Date, date1904))}</v></c>");
                    break;
                case CellType.Boolean:
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"b\"><f>{fEsc}</f><v>{(cell.Boolean ? 1 : 0)}</v></c>");
                    break;
                default:
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr}><f>{fEsc}</f></c>");
                    break;
            }
            return;
        }

        switch (cell.Type)
        {
            case CellType.Empty:
                sb.Append($"<c r=\"{cellRef}\"{styleAttr}/>");
                break;

            case CellType.Text:
                if (string.IsNullOrEmpty(cell.Text))
                {
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr}/>");
                }
                else if (sharedIndex.TryGetValue(cell.Text, out var idx))
                {
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"s\"><v>{idx}</v></c>");
                }
                else
                {
                    sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"inlineStr\"><is><t>{XmlEscape(cell.Text)}</t></is></c>");
                }
                break;

            case CellType.Number:
                sb.Append($"<c r=\"{cellRef}\"{styleAttr}><v>{FormatDouble(cell.Number)}</v></c>");
                break;

            case CellType.Date:
                var serial = FormatDetector.DateToSerial(cell.Date, date1904);
                sb.Append($"<c r=\"{cellRef}\"{styleAttr}><v>{FormatDouble(serial)}</v></c>");
                break;

            case CellType.Boolean:
                sb.Append($"<c r=\"{cellRef}\"{styleAttr} t=\"b\"><v>{(cell.Boolean ? 1 : 0)}</v></c>");
                break;
        }
    }

    // ── OOXML 部件构建 ──

    private static string ContentTypesXml(int sheetCount, IReadOnlyList<int> sheetsWithComments, bool hasProps, OoxmlPreservedParts? preserved, bool macroEnabled = false, ImagePlan? imagePlan = null)
    {
        var defaults = new List<(string Ext, string Ct)>();
        var overrides = new List<(string Part, string Ct)>();

        // 写入器固有声明；xlsm 的主文档类型必须为 macroEnabled，否则 Excel 拒绝打开
        defaults.Add(("rels", "application/vnd.openxmlformats-package.relationships+xml"));
        defaults.Add(("xml", "application/xml"));
        overrides.Add(("/xl/workbook.xml", macroEnabled
            ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
        for (int i = 1; i <= sheetCount; i++)
        {
            overrides.Add(($"/xl/worksheets/sheet{i}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
        }
        overrides.Add(("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"));
        overrides.Add(("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
        foreach (var sheetIdx in sheetsWithComments)
        {
            overrides.Add(($"/xl/comments{sheetIdx + 1}.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"));
        }
        if (hasProps)
        {
            overrides.Add(("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"));
            overrides.Add(("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"));
        }

        // 图片：media 类型 Default + drawing / richData Override
        if (imagePlan is { Any: true })
        {
            var imgSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "rels", "xml" };
            foreach (var img in imagePlan.All)
            {
                if (imgSeen.Add(img.EffectiveExtension))
                {
                    string ct = img.EffectiveExtension switch
                    {
                        "png" => "image/png",
                        "jpg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        _ => "application/octet-stream",
                    };
                    defaults.Add((img.EffectiveExtension, ct));
                }
            }
            for (int i = 0; i < imagePlan.FloatingBySheet.Count; i++)
            {
                if (imagePlan.FloatingBySheet[i].Count > 0)
                    overrides.Add(($"/{imagePlan.DrawingTargetFor(i, preserved).Entry}", "application/vnd.openxmlformats-officedocument.drawing+xml"));
            }
            if (imagePlan.HasInCell)
            {
                overrides.Add(("/xl/metadata.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"));
                overrides.Add(("/xl/richData/richValueRel.xml", "application/vnd.ms-excel.richvaluerel+xml"));
                overrides.Add(("/xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml"));
                overrides.Add(("/xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml"));
                overrides.Add(("/xl/richData/rdRichValueTypes.xml", "application/vnd.ms-excel.rdrichvaluetypes+xml"));
            }
        }

        // 保留的声明（排除与重建部件冲突的）
        if (preserved is not null)
        {
            foreach (var d in preserved.DefaultTypes)
            {
                if (d.Extension == "rels" || d.Extension == "xml") continue;
                defaults.Add(d);
            }
            var rebuiltEntries = OoxmlPreservedParts.BuildRebuiltEntries(sheetCount);
            foreach (var o in preserved.OverrideTypes)
            {
                if (rebuiltEntries.Contains(o.PartName.TrimStart('/'))) continue;
                overrides.Add(o);
            }
        }

        // 去重后输出
        var seenExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var seenPart = new HashSet<string>(StringComparer.Ordinal);
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
        foreach (var d in defaults)
        {
            if (seenExt.Add(d.Ext))
                sb.Append($"<Default Extension=\"{d.Ext}\" ContentType=\"{d.Ct}\"/>");
        }
        foreach (var o in overrides)
        {
            if (seenPart.Add(o.Part))
                sb.Append($"<Override PartName=\"{o.Part}\" ContentType=\"{o.Ct}\"/>");
        }
        sb.Append("</Types>");
        return sb.ToString();
    }

    // ── rels 合并（保留部件） ──

    internal sealed class RelInfo
    {
        public string Id = "";
        public string Type = "";
        public string Target = "";
        public string TargetMode = "";
    }

    /// <summary>合并工作表级 rels：保留未重建目标（绘图/超链接等），追加新建超链接与批注 rel。返回 null 表示无需写出 </summary>
    private static string? MergeSheetRels(int sheetNumber, bool hasComments, OoxmlPreservedParts? preserved,
        bool mergeSheetRels, List<(string Ref, string Target, string? Tooltip, bool IsInternal)>? hyperlinks = null,
        bool hasDrawing = false, XlsxWriter.ImagePlan? imagePlan = null)
    {
        string original = "";
        if (mergeSheetRels && preserved is not null
            && preserved.Rels.TryGetValue($"xl/worksheets/_rels/sheet{sheetNumber}.xml.rels", out var r))
        {
            original = r;
        }

        // 重建 rels（完整 XML）：批注 + 超链接 + drawing
        var relParts = new List<RelInfo>();
        if (hasComments)
        {
            relParts.Add(new RelInfo
            {
                Id = "rId1",
                Type = $"{OfficeRelNs}/comments",
                Target = $"../comments{sheetNumber}.xml",
            });
        }
        if (hyperlinks is { Count: > 0 })
        {
            int extIndex = 0;
            for (int h = 0; h < hyperlinks.Count; h++)
            {
                if (hyperlinks[h].IsInternal) continue; // 内部链接不需要 rels
                extIndex++;
                relParts.Add(new RelInfo
                {
                    Id = $"rIdH{extIndex}",
                    Type = $"{OfficeRelNs}/hyperlink",
                    Target = hyperlinks[h].Target,
                    TargetMode = "External",
                });
            }
        }
        if (hasDrawing)
        {
            // 既有 drawing 时其 rel 已由 MergeRelsXml 保留，无需新增；新建 drawing 才加 rIdD1
            var (drawingEntry, _) = imagePlan is not null
                ? imagePlan.DrawingTargetFor(sheetNumber - 1, preserved)
                : ($"xl/drawings/drawing{sheetNumber}.xml", "rIdD1");
            bool hasExistingDrawingRel = false;
            if (!string.IsNullOrEmpty(original))
            {
                foreach (var rel in ParseRels(original))
                {
                    if (rel.Type.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase))
                    {
                        var abs = ResolveRelsTarget("xl/worksheets", rel.Target);
                        if (abs == drawingEntry) { hasExistingDrawingRel = true; break; }
                    }
                }
            }
            if (!hasExistingDrawingRel)
            {
                relParts.Add(new RelInfo
                {
                    Id = "rIdD1",
                    Type = $"{OfficeRelNs}/drawing",
                    Target = $"../drawings/{drawingEntry.Substring(drawingEntry.LastIndexOf('/') + 1)}",
                });
            }
        }

        string rebuilt = RelsXml(relParts);
        var rebuiltTargets = new HashSet<string> { $"xl/comments{sheetNumber}.xml" };
        return MergeRelsXml(original, "xl/worksheets", rebuiltTargets, rebuilt);
    }

    /// <summary>把 rel 列表序列化为完整 &lt;Relationships&gt; XML </summary>
    internal static string RelsXml(List<RelInfo> rels)
    {
        if (rels.Count == 0) return "";
        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        foreach (var rel in rels)
        {
            sb.Append($"<Relationship Id=\"{rel.Id}\" Type=\"{XmlEscape(rel.Type)}\" Target=\"{XmlEscape(rel.Target)}\"");
            if (rel.TargetMode.Length > 0)
                sb.Append($" TargetMode=\"{XmlEscape(rel.TargetMode)}\"");
            sb.Append("/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    /// <summary>
    /// 合并 rels：保留原始 rels 中不指向重建部件的条目（外部链接始终保留），
    /// 追加写入器生成的重建 rels。保留条目的 rId 重新编号以避免冲突。
    /// 全部为空时返回 null。
    /// </summary>
    internal static string? MergeRelsXml(string originalRelsXml, string baseDir, HashSet<string> rebuiltTargets, string rebuiltRelsXml)
    {
        var kept = new List<RelInfo>();
        if (!string.IsNullOrEmpty(originalRelsXml))
        {
            foreach (var rel in ParseRels(originalRelsXml))
            {
                if (rel.TargetMode == "External")
                {
                    kept.Add(rel);
                    continue;
                }
                var abs = ResolveRelsTarget(baseDir, rel.Target);
                if (rebuiltTargets.Contains(abs)) continue;
                kept.Add(rel);
            }
        }

        var rebuilt = ParseRels(rebuiltRelsXml);
        if (kept.Count == 0 && rebuilt.Count == 0) return null;

        var usedIds = new HashSet<string>(rebuilt.Select(x => x.Id), StringComparer.Ordinal);
        int next = 1;
        foreach (var rel in kept)
        {
            string id;
            do { id = "rId" + next++; } while (usedIds.Contains(id));
            rel.Id = id;
            usedIds.Add(id);
        }

        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        foreach (var rel in kept.Concat(rebuilt))
        {
            sb.Append($"<Relationship Id=\"{rel.Id}\" Type=\"{XmlEscape(rel.Type)}\" Target=\"{XmlEscape(rel.Target)}\"");
            if (rel.TargetMode.Length > 0)
                sb.Append($" TargetMode=\"{XmlEscape(rel.TargetMode)}\"");
            sb.Append("/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    internal static List<RelInfo> ParseRels(string xml)
    {
        var list = new List<RelInfo>();
        if (string.IsNullOrEmpty(xml)) return list;
        var doc = XDocument.Parse(xml);
        if (doc.Root is null) return list;
        var ns = doc.Root.GetDefaultNamespace();
        foreach (var el in doc.Root.Elements(ns + "Relationship"))
        {
            list.Add(new RelInfo
            {
                Id = (string?)el.Attribute("Id") ?? "",
                Type = (string?)el.Attribute("Type") ?? "",
                Target = (string?)el.Attribute("Target") ?? "",
                TargetMode = (string?)el.Attribute("TargetMode") ?? "",
            });
        }
        return list;
    }

    internal static string ResolveRelsTarget(string baseDir, string target)
    {
        target = target.Replace('\\', '/');
        if (target.StartsWith("/")) return target.TrimStart('/');

        var combined = (string.IsNullOrEmpty(baseDir) ? "" : baseDir + "/") + target;
        var parts = combined.Split('/');
        var stack = new List<string>();
        foreach (var p in parts)
        {
            if (string.IsNullOrEmpty(p) || p == ".") continue;
            if (p == "..")
            {
                if (stack.Count > 0) stack.RemoveAt(stack.Count - 1);
                continue;
            }
            stack.Add(p);
        }
        return string.Join("/", stack);
    }

    private static string CommentsXml(IReadOnlyDictionary<string, string> comments)
    {
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<comments xmlns=\"{MainNs}\">");
        sb.Append("<authors><author>LiteExcel</author></authors>");
        sb.Append("<commentList>");
        foreach (var kv in comments)
        {
            var refAttr = XmlEscape(kv.Key);
            var text = XmlEscape(kv.Value);
            // 保留前导/尾随空格
            string spaceAttr = "";
            if (kv.Value.Length > 0 && (char.IsWhiteSpace(kv.Value[0]) || char.IsWhiteSpace(kv.Value[kv.Value.Length - 1])))
            {
                spaceAttr = " xml:space=\"preserve\"";
            }
            sb.Append($"<comment ref=\"{refAttr}\" authorId=\"0\"><text><t{spaceAttr}>{text}</t></text></comment>");
        }
        sb.Append("</commentList>");
        sb.Append("</comments>");
        return sb.ToString();
    }

    private static string RootRelsXml(bool hasProps, OoxmlPreservedParts? preserved)
    {
        var rebuiltTargets = new HashSet<string> { "xl/workbook.xml" };
        if (hasProps)
        {
            rebuiltTargets.Add("docProps/core.xml");
            rebuiltTargets.Add("docProps/app.xml");
        }

        string original = "";
        if (preserved is not null && preserved.Rels.TryGetValue("_rels/.rels", out var r))
            original = r;

        return MergeRelsXml(original, "", rebuiltTargets, WriterRootRels(hasProps)) ?? WriterRootRels(hasProps);
    }

    private static string WriterRootRels(bool hasProps)
    {
        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        sb.Append($"<Relationship Id=\"rId1\" Type=\"{OfficeRelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>");
        if (hasProps)
        {
            sb.Append($"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
            sb.Append($"<Relationship Id=\"rId3\" Type=\"{OfficeRelNs}/extended-properties\" Target=\"docProps/app.xml\"/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    internal static string CorePropsXml(WorkbookProperties props)
    {
        var now = DateTime.UtcNow;
        var created = props.Created ?? now;
        var modified = props.Modified ?? now;
        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"");
        sb.Append(" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"");
        sb.Append(" xmlns:dcterms=\"http://purl.org/dc/terms/\"");
        sb.Append(" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"");
        sb.Append(" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        if (!string.IsNullOrEmpty(props.Creator))
            sb.Append($"<dc:creator>{XmlEscape(props.Creator!)}</dc:creator>");
        if (!string.IsNullOrEmpty(props.LastModifiedBy))
            sb.Append($"<cp:lastModifiedBy>{XmlEscape(props.LastModifiedBy!)}</cp:lastModifiedBy>");
        sb.Append($"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{XmlEscape(created.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture))}</dcterms:created>");
        sb.Append($"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{XmlEscape(modified.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture))}</dcterms:modified>");
        if (!string.IsNullOrEmpty(props.Title))
            sb.Append($"<dc:title>{XmlEscape(props.Title!)}</dc:title>");
        if (!string.IsNullOrEmpty(props.Subject))
            sb.Append($"<dc:subject>{XmlEscape(props.Subject!)}</dc:subject>");
        sb.Append("</cp:coreProperties>");
        return sb.ToString();
    }

    internal static string AppPropsXml(WorkbookProperties props, IReadOnlyList<SheetData> sheets)
    {
        // Application 默认取宿主程序集名；显式设置则优先
        string application = !string.IsNullOrEmpty(props.Application)
            ? props.Application!
            : (System.Reflection.Assembly.GetEntryAssembly()?.GetName().Name ?? "LiteExcel");

        var sb = new StringBuilder(512);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"");
        sb.Append(" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
        sb.Append($"<Application>{XmlEscape(application)}</Application>");
        sb.Append("<DocSecurity>0</DocSecurity>");
        sb.Append("<ScaleCrop>false</ScaleCrop>");
        sb.Append($"<HeadingPairs><vt:vector size=\"2\" baseType=\"variant\"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{sheets.Count}</vt:i4></vt:variant></vt:vector></HeadingPairs>");
        sb.Append($"<TitlesOfParts><vt:vector size=\"{sheets.Count}\" baseType=\"lpstr\">");
        foreach (var sheet in sheets)
            sb.Append($"<vt:lpstr>{XmlEscape(sheet.SheetName)}</vt:lpstr>");
        sb.Append("</vt:vector></TitlesOfParts>");
        sb.Append("<Company></Company>");
        sb.Append("<LinksUpToDate>false</LinksUpToDate>");
        sb.Append("<SharedDoc>false</SharedDoc>");
        sb.Append("<HyperlinksChanged>false</HyperlinksChanged>");
        sb.Append("<AppVersion>16.0300</AppVersion>");
        sb.Append("</Properties>");
        return sb.ToString();
    }

    private static string WorkbookXml(IReadOnlyList<SheetData> sheets, OoxmlPreservedParts? preserved, bool date1904,
        string? fileSharingHash = null, string? fileSharingSalt = null, int? fileSharingSpin = null,
        bool fileSharingReadOnlyRecommended = false)
    {
        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{OfficeRelNs}\">");
        // 工作簿宿主 VBA 代码名：schema 要求位于 sheets 之前，缺失会导致 Excel 重排宏文档模块
        // date1904：1904 日期系统（1904-01-01 基准，读取侧 XlsxReader 按此换算）
        var wbCodeName = preserved?.WorkbookCodeName;
        bool hasWbAttr = !string.IsNullOrEmpty(wbCodeName) || date1904;
        if (hasWbAttr)
        {
            sb.Append("<workbookPr");
            if (!string.IsNullOrEmpty(wbCodeName))
                sb.Append($" codeName=\"{XmlEscape(wbCodeName)}\"");
            if (date1904)
                sb.Append(" date1904=\"1\"");
            sb.Append("/>");
        }
        // fileSharing（修改密码 / 写保护）：位于 workbookPr 之后、sheets 之前
        if (!string.IsNullOrEmpty(fileSharingHash))
        {
            sb.Append($"<fileSharing readOnlyRecommended=\"{(fileSharingReadOnlyRecommended ? 1 : 0)}\" " +
                      $"userName=\"Admin\" algorithmName=\"SHA-512\" " +
                      $"hashValue=\"{fileSharingHash}\" saltValue=\"{fileSharingSalt ?? ""}\" " +
                      $"spinCount=\"{fileSharingSpin ?? 100000}\"/>");
        }
        // P0-6: bookViews 原样回写（schema 位于 sheets 之前，保留窗口视图/活动表）
        if (preserved?.BookViewsXml is { Length: > 0 })
            sb.Append(preserved.BookViewsXml);
        sb.Append("<sheets>");
        for (int i = 0; i < sheets.Count; i++)
        {
            var name = XmlEscape(sheets[i].SheetName);
            if (string.IsNullOrEmpty(name)) name = $"Sheet{i + 1}";
            sb.Append($"<sheet name=\"{name}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>");
        }
        sb.Append("</sheets>");
        // P0-6: definedNames 原样回写（schema 位于 sheets 之后，保留命名区域）
        if (preserved?.DefinedNamesXml is { Length: > 0 })
            sb.Append(preserved.DefinedNamesXml);
        // P0-12: 陈旧 calcChain 不透传，写 fullCalcOnLoad 让 Excel 保存时重建计算链
        sb.Append("<calcPr fullCalcOnLoad=\"1\"/>");
        sb.Append("</workbook>");
        return sb.ToString();
    }

    private static string WorkbookRelsXml(int sheetCount, OoxmlPreservedParts? preserved, ImagePlan? imagePlan = null)
    {
        var rebuiltTargets = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 1; i <= sheetCount; i++)
        {
            rebuiltTargets.Add($"xl/worksheets/sheet{i}.xml");
        }
        rebuiltTargets.Add("xl/sharedStrings.xml");
        rebuiltTargets.Add("xl/styles.xml");
        rebuiltTargets.Add("xl/calcChain.xml"); // P0-12: 陈旧 calcChain 不透传，其 rel 一并丢弃
        if (imagePlan is { HasInCell: true })
        {
            rebuiltTargets.Add("xl/metadata.xml");
            rebuiltTargets.Add("xl/richData/richValueRel.xml");
            rebuiltTargets.Add("xl/richData/rdrichvalue.xml");
            rebuiltTargets.Add("xl/richData/rdrichvaluestructure.xml");
            rebuiltTargets.Add("xl/richData/rdRichValueTypes.xml");
        }

        string original = "";
        if (preserved is not null && preserved.Rels.TryGetValue("xl/_rels/workbook.xml.rels", out var r))
            original = r;

        return MergeRelsXml(original, "xl", rebuiltTargets, WriterWorkbookRels(sheetCount, imagePlan)) ?? WriterWorkbookRels(sheetCount, imagePlan);
    }

    private static string WriterWorkbookRels(int sheetCount, ImagePlan? imagePlan = null)
    {
        var sb = new StringBuilder(256);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<Relationships xmlns=\"{RelNs}\">");
        for (int i = 1; i <= sheetCount; i++)
        {
            sb.Append($"<Relationship Id=\"rId{i}\" Type=\"{OfficeRelNs}/worksheet\" Target=\"worksheets/sheet{i}.xml\"/>");
        }
        sb.Append($"<Relationship Id=\"rId{sheetCount + 1}\" Type=\"{OfficeRelNs}/sharedStrings\" Target=\"sharedStrings.xml\"/>");
        sb.Append($"<Relationship Id=\"rId{sheetCount + 2}\" Type=\"{OfficeRelNs}/styles\" Target=\"styles.xml\"/>");

        // InCell richData 部件（workbook 级关系）
        int rid = sheetCount + 3;
        if (imagePlan is { HasInCell: true })
        {
            sb.Append($"<Relationship Id=\"rId{rid++}\" Type=\"{OfficeRelNs}/sheetMetadata\" Target=\"metadata.xml\"/>");
            sb.Append($"<Relationship Id=\"rId{rid++}\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue\" Target=\"richData/rdrichvalue.xml\"/>");
            sb.Append($"<Relationship Id=\"rId{rid++}\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure\" Target=\"richData/rdrichvaluestructure.xml\"/>");
            sb.Append($"<Relationship Id=\"rId{rid++}\" Type=\"http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes\" Target=\"richData/rdRichValueTypes.xml\"/>");
            sb.Append($"<Relationship Id=\"rId{rid++}\" Type=\"http://schemas.microsoft.com/office/2022/10/relationships/richValueRel\" Target=\"richData/richValueRel.xml\"/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    private static string SharedStringsXml(List<string> shared)
    {
        var sb = new StringBuilder(shared.Count * 32 + 128);
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append($"<sst xmlns=\"{MainNs}\" count=\"{shared.Count}\" uniqueCount=\"{shared.Count}\">");
        foreach (var s in shared)
        {
            sb.Append("<si><t");
            if (s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[s.Length - 1])))
            {
                sb.Append(" xml:space=\"preserve\"");
            }
            sb.Append(">").Append(XmlEscape(s)).Append("</t></si>");
        }
        sb.Append("</sst>");
        return sb.ToString();
    }

    // ── 工具方法 ──

    private static void WriteXmlEntry(ZipArchive zip, string entryName, string xml)
    {
        var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        var bytes = new UTF8Encoding(false).GetBytes(xml);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteEntry(ZipArchive zip, string entryName, byte[] bytes)
    {
        var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }

    internal static string FormatDouble(double d)
    {
        if (double.IsNaN(d) || double.IsInfinity(d)) return "0";
        if (d == Math.Floor(d) && Math.Abs(d) < 1e15)
        {
            return ((long)d).ToString(CultureInfo.InvariantCulture);
        }
        return d.ToString("R", CultureInfo.InvariantCulture);
    }

    internal static string XmlEscape(string s)
    {
        if (string.IsNullOrEmpty(s)) return s ?? "";
        var sb = new StringBuilder(s.Length + 16);
        foreach (var ch in s)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&apos;"); break;
                default: sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>内部超链接 location 以 '#' 开头时去除（OOXML location 不带 '#'） </summary>
    private static string NormalizeInternalLocation(string target)
    {
        return target.StartsWith("#", StringComparison.Ordinal) ? target.Substring(1) : target;
    }

    // ── 列宽自适应 ──

    /// <summary>
    /// 估算并设置每列的宽度（中文字符算 2，英文/数字算 1），范围 [8, 50] 
    /// 调用此方法后再调用 <see cref="Write(string, SheetData)"/> 
    /// </summary>
    public static void AutoColumnWidths(SheetData sheet)
    {
        if (sheet is null) throw new ArgumentNullException(nameof(sheet));

        int colCount = 0;
        if (sheet.Headers is { Count: > 0 }) colCount = sheet.Headers.Count;
        foreach (var row in sheet.Rows)
        {
            if (row.Count > colCount) colCount = row.Count;
        }
        if (colCount == 0) return;

        var widths = new double[colCount];

        // 表头
        if (sheet.Headers is not null)
        {
            for (int c = 0; c < sheet.Headers.Count; c++)
            {
                double w = EstimateTextWidth(sheet.Headers[c] ?? "");
                if (w > widths[c]) widths[c] = w;
            }
        }

        // 数据行
        foreach (var row in sheet.Rows)
        {
            for (int c = 0; c < row.Count; c++)
            {
                double w = EstimateCellWidth(row[c]);
                if (w > widths[c]) widths[c] = w;
            }
        }

        // 应用最小/最大限制
        var result = new List<double>(colCount);
        for (int i = 0; i < colCount; i++)
        {
            double w = widths[i];
            if (w < 8) w = 8;
            if (w > 50) w = 50;
            result.Add(w);
        }
        sheet.ColumnWidths = result;
    }

    private static double EstimateTextWidth(string s)
    {
        if (string.IsNullOrEmpty(s)) return 0;
        double w = 0;
        foreach (var ch in s)
        {
            w += IsWideChar(ch) ? 2 : 1;
        }
        return w;
    }

    private static double EstimateCellWidth(Cell cell)
    {
        switch (cell.Type)
        {
            case CellType.Text:
                return EstimateTextWidth(cell.Text ?? "");
            case CellType.Number:
                return EstimateTextWidth(FormatDouble(cell.Number));
            case CellType.Date:
                return EstimateTextWidth(cell.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            case CellType.Boolean:
                return cell.Boolean ? 4 : 5; // TRUE / FALSE
            case CellType.Empty:
            default:
                return 0;
        }
    }

    private static bool IsWideChar(char ch)
    {
        // CJK 及全角字符算宽字符
        if (ch >= 0x1100 && ch <= 0x115F) return true;   // Hangul Jamo
        if (ch >= 0x2E80 && ch <= 0x303E) return true;   // CJK Radicals / Kangxi
        if (ch >= 0x3040 && ch <= 0x33BF) return true;   // Hiragana / Katakana / CJK symbols
        if (ch >= 0x3400 && ch <= 0x4DBF) return true;   // CJK Unified Extension A
        if (ch >= 0x4E00 && ch <= 0x9FFF) return true;   // CJK Unified Ideographs
        if (ch >= 0xA000 && ch <= 0xA4CF) return true;   // Yi
        if (ch >= 0xAC00 && ch <= 0xD7A3) return true;   // Hangul Syllables
        if (ch >= 0xF900 && ch <= 0xFAFF) return true;   // CJK Compatibility Ideographs
        if (ch >= 0xFE30 && ch <= 0xFE4F) return true;   // CJK Compatibility Forms
        if (ch >= 0xFF00 && ch <= 0xFF60) return true;   // Fullwidth Forms
        if (ch >= 0xFFE0 && ch <= 0xFFE6) return true;   // Fullwidth Signs
        if (ch >= 0x20000 && ch <= 0x2FFFD) return true; // CJK Extension B-F
        if (ch >= 0x30000 && ch <= 0x3FFFD) return true; // CJK Extension G+
        return false;
    }
}

