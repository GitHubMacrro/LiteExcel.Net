using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// xls（BIFF8）工作表尺寸上限回归：数据超出 256 列 / 65536 行时必须裁剪，
/// 不得写出 Excel 打开时需修复的文件。见 XlsWriter 的 MaxColBiff8/MaxRowBiff8 裁剪。
/// </summary>
public class XlsSheetSizeTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"xlssize_{Guid.NewGuid():N}{ext}");

    [Fact]
    public void Xls_Write_WideSheetBeyond256Cols_SavesWithoutRepair()
    {
        // xlsx 工作表数据超出 xls 的 256 列上限时，写出 xls 应裁剪而非生成结构非法文件
        var xlsx = GetTempFile(".xlsx");
        var xls = GetTempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            wb.Worksheets[0].SetValue("A1", "in-range");
            wb.Worksheets[0].Cells[1, 300].Value = "out-of-range"; // 超出 256 列
            wb.SaveAs(xlsx);

            var opened = Excel.Open(xlsx);
            opened.SaveAs(xls, ExcelFormat.Xls);
            Assert.True(File.Exists(xls));

            // 范围内数据保留；文件可被读回（结构合法）
            var reopened = Excel.Open(xls);
            Assert.Equal("in-range", reopened.Worksheets[0].Cell("A1").GetString());

            // 越界裁剪应被上报
            bool reported = false;
            foreach (var d in opened.SaveDegradations)
                if (d.Capability == DegradationCapability.SheetSize) reported = true;
            Assert.True(reported);
        }
        finally
        {
            if (File.Exists(xlsx)) File.Delete(xlsx);
            if (File.Exists(xls)) File.Delete(xls);
        }
    }

    [Fact]
    public void Xls_Write_ColumnWidthsBeyond256Cols_SavesWithoutRepair()
    {
        // 列宽声明常超出数据范围（如声明到第 386 列而数据仅到第 12 列）：写 xls 时必须裁剪列宽
        var xlsx = GetTempFile(".xlsx");
        var xls = GetTempFile(".xls");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsx);
            wb.Worksheets[0].SetValue("A1", "v");
            var widths = new Dictionary<int, double>();
            for (int c = 0; c < 400; c++) widths[c] = 10;
            wb.Worksheets[0].ColumnWidths = widths;
            wb.SaveAs(xlsx);

            var opened = Excel.Open(xlsx);
            opened.SaveAs(xls, ExcelFormat.Xls);
            Assert.True(File.Exists(xls));

            var reopened = Excel.Open(xls);
            Assert.Equal("v", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally
        {
            if (File.Exists(xlsx)) File.Delete(xlsx);
            if (File.Exists(xls)) File.Delete(xls);
        }
    }
}
