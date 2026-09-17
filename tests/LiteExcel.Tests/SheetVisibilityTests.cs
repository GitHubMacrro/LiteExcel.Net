using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 批次 T1-A：工作表可见性（SheetVisibility）与标签颜色（TabColor）。
/// 覆盖 xlsx/xlsm/xlsb/xls 可见性往返；tabColor 仅 xlsx/xlsm，xlsb/xls 降级上报。
/// </summary>
public class SheetVisibilityTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"sheetvis_{Guid.NewGuid():N}{ext}");

    [Theory]
    [InlineData(ExcelFormat.Xlsx, ".xlsx")]
    [InlineData(ExcelFormat.Xlsm, ".xlsm")]
    [InlineData(ExcelFormat.Xlsb, ".xlsb")]
    [InlineData(ExcelFormat.Xls, ".xls")]
    public void Visibility_RoundTrip(ExcelFormat format, string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(format);
            wb.Worksheets[0].Name = "V1";
            var hidden = wb.Worksheets.Add("H1");
            hidden.Visible = SheetVisibility.Hidden;
            var very = wb.Worksheets.Add("VH1");
            very.Visible = SheetVisibility.VeryHidden;
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Equal(SheetVisibility.Visible, reopened.Worksheets["V1"].Visible);
            Assert.Equal(SheetVisibility.Hidden, reopened.Worksheets["H1"].Visible);
            Assert.Equal(SheetVisibility.VeryHidden, reopened.Worksheets["VH1"].Visible);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsm")]
    public void TabColor_RoundTrip(string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(ext == ".xlsm" ? ExcelFormat.Xlsm : ExcelFormat.Xlsx);
            wb.Worksheets[0].Name = "S1";
            wb.Worksheets[0].TabColor = "FF0000";
            wb.Worksheets.Add("S2").TabColor = "00FF00";
            wb.SaveAs(file);

            var reopened = Excel.Open(file);
            Assert.Equal("FF0000", reopened.Worksheets["S1"].TabColor);
            Assert.Equal("00FF00", reopened.Worksheets["S2"].TabColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Theory]
    [InlineData(ExcelFormat.Xlsb, ".xlsb")]
    [InlineData(ExcelFormat.Xls, ".xls")]
    public void TabColor_Unsupported_ReportsDegradation(ExcelFormat format, string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(format);
            wb.Worksheets[0].Name = "S1";
            wb.Worksheets[0].TabColor = "FF0000";
            wb.SaveAs(file);

            bool reported = false;
            foreach (var d in wb.SaveDegradations)
                if (d.Capability == DegradationCapability.SheetVisibility) reported = true;
            Assert.True(reported, $"{format} 应上报 tabColor 降级");
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void HidingLastVisibleSheet_Throws()
    {
        var wb = Excel.Create();
        wb.Worksheets.Add("B");
        wb.Worksheets[0].Visible = SheetVisibility.Hidden; // 还有 B 可见，允许
        Assert.Equal(SheetVisibility.Hidden, wb.Worksheets[0].Visible);
        Assert.Throws<LiteExcelException>(() => wb.Worksheets[1].Visible = SheetVisibility.Hidden);
    }

    [Fact]
    public void VisibilityChange_MarksModified_DisablesVerbatim()
    {
        // 打开 xlsb 后仅改可见性，必须触发重建（否则 verbatim 会写回旧的可见性）
        var file = GetTempFile(".xlsb");
        var outPath = GetTempFile(".xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            wb.Worksheets[0].Name = "A";
            wb.Worksheets.Add("B");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            opened.Worksheets["B"].Visible = SheetVisibility.Hidden;
            opened.SaveAs(outPath, ExcelFormat.Xlsb);

            var reopened = Excel.Open(outPath);
            Assert.Equal(SheetVisibility.Hidden, reopened.Worksheets["B"].Visible);
        }
        finally
        {
            if (File.Exists(file)) File.Delete(file);
            if (File.Exists(outPath)) File.Delete(outPath);
        }
    }
}
