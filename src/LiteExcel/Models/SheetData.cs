namespace LiteExcel;

/// <summary>
/// 一张工作表的完整数据 
/// </summary>
public sealed class SheetData
{
    public string SheetName { get; set; } = "Sheet1";
    public List<string> Headers { get; set; } = new();
    public List<IReadOnlyList<Cell>> Rows { get; set; } = new();
    public List<CellRange> MergedRanges { get; set; } = new();
    public AutoFilter? Filter { get; set; }
    public bool FreezeHeader { get; set; }

    /// <summary>冻结行数（0 = 不冻结行） </summary>
    public int FreezeRows { get; set; }

    /// <summary>冻结列数（0 = 不冻结列） </summary>
    public int FreezeColumns { get; set; }
    public List<double>? ColumnWidths { get; set; }
    public CellStyle? HeaderStyle { get; set; }

    /// <summary>全表默认样式（优先级最低） </summary>
    public CellStyle? DefaultStyle { get; set; }

    /// <summary>行级样式（key = 0-based 行索引，对应 Rows） </summary>
    public Dictionary<int, CellStyle>? RowStyles { get; set; }

    /// <summary>列级样式（key = 0-based 列索引） </summary>
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }

    /// <summary>行高（key = 0-based 行索引，对应 Rows） 单位：磅（point） </summary>
    public Dictionary<int, double>? RowHeights { get; set; }

    /// <summary>单元格批注（key = A1 格式单元格引用，value = 批注文本） </summary>
    public Dictionary<string, string>? Comments { get; set; }

    /// <summary>数据验证规则列表 </summary>
    public List<DataValidation>? Validations { get; set; }

    /// <summary>工作表宿主的 VBA 代码名（sheetPr@codeName）。带宏工作簿经保存后仍与 vbaProject 绑定，避免 Excel 重排文档模块 </summary>
    public string? CodeName { get; set; }

    /// <summary>工作表图片（InCell richData / Floating drawing） </summary>
    public List<WorksheetImage>? Images { get; set; }
}