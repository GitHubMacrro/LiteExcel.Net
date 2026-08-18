using LiteExcel;

namespace LiteExcel.Tests;

/// <summary>
/// 异常和降级行为统一测试。
/// 覆盖：扩展名推断一致性、宏不静默丢失、不支持能力的明确报错。
/// </summary>
public class DegradationBehaviorTests
{
    [Fact]
    public void ExcelWrite_WithXlsExtension_ConvertsToXls()
    {
        // Excel.Write 的扩展名推断与 DetectFormat 完全一致：.xls 扩展名 → 目标格式为 Xls
        var wb = Excel.Create(ExcelFormat.Xlsx);
        wb.Worksheets["Sheet1"].SetValue("A1", "数据");
        wb.Worksheets["Sheet1"].SetValue("A2", 42);

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-write-convert.xls");
        Excel.Write(path, wb);

        var reopened = Excel.Open(path);
        Assert.Equal(ExcelFormat.Xls, reopened.Format);
        Assert.Equal("数据", reopened.Worksheets[0].Cell("A1").GetString());
        Assert.Equal(42, reopened.Worksheets[0].Cell("A2").GetDouble());
        File.Delete(path);
    }

    [Fact]
    public void ExcelWrite_WithXlsbExtension_ConvertsToXlsb()
    {
        var wb = Excel.Create(ExcelFormat.Xlsx);
        wb.Worksheets["Sheet1"].SetValue("A1", "二进制");
        wb.Worksheets["Sheet1"].SetValue("A2", 3.14);

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-write-convert.xlsb");
        Excel.Write(path, wb);

        var reopened = Excel.Open(path);
        Assert.Equal(ExcelFormat.Xlsb, reopened.Format);
        Assert.Equal("二进制", reopened.Worksheets[0].Cell("A1").GetString());
        Assert.Equal(3.14, reopened.Worksheets[0].Cell("A2").GetDouble());
        File.Delete(path);
    }

    [Fact]
    public void SaveAs_Xls_WithMacro_ThrowsToPreventSilentLoss()
    {
        // 宏不静默丢失：有宏的工作簿写 .xls（不支持宏）必须明确报错
        var wb = Excel.Create(ExcelFormat.Xlsm);
        // 通过 friend 程序集注入假宏字节，模拟打开 xlsm 捕获的宏
        var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        prop!.SetValue(wb, new byte[] { 0x00, 0x01, 0x02 });

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-macro-to-xls.xls");
        if (File.Exists(path)) File.Delete(path);

        var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(path, ExcelFormat.Xls));
        Assert.Contains("宏", ex.Message);
        Assert.False(File.Exists(path)); // 不应生成残缺文件
    }

    [Fact]
    public void SaveAs_Xlsx_WithMacro_ThrowsToPreventSilentLoss()
    {
        // 宏不静默丢失：有宏的工作簿写 .xlsx（不支持宏）必须明确报错，避免生成不一致文件
        var wb = Excel.Create(ExcelFormat.Xlsm);
        var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        prop!.SetValue(wb, new byte[] { 0x00, 0x01, 0x02 });

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-macro-to-xlsx.xlsx");
        if (File.Exists(path)) File.Delete(path);

        var ex = Assert.Throws<LiteExcelException>(() => wb.SaveAs(path, ExcelFormat.Xlsx));
        Assert.Contains("宏", ex.Message);
        Assert.False(File.Exists(path)); // 不应生成残缺文件
    }

    [Fact]
    public void Save_StreamXlsx_WithMacro_ThrowsToPreventSilentLoss()
    {
        // Stream 保存同样受宏保护
        var wb = Excel.Create(ExcelFormat.Xlsm);
        var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        prop!.SetValue(wb, new byte[] { 0x00, 0x01, 0x02 });

        using var ms = new MemoryStream();
        var ex = Assert.Throws<LiteExcelException>(() => wb.Save(ms, ExcelFormat.Xlsx));
        Assert.Contains("宏", ex.Message);
    }

    [Fact]
    public void SaveAs_Xlsm_WithMacro_StillWorks()
    {
        // 有宏的工作簿保存为 xlsm 应正常工作
        var wb = Excel.Create(ExcelFormat.Xlsm);
        var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        prop!.SetValue(wb, new byte[] { 0x00, 0x01, 0x02 });

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-macro-to-xlsm.xlsm");
        try
        {
            wb.SaveAs(path, ExcelFormat.Xlsm);
            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SaveAs_Xlsb_WithMacro_StillWorks()
    {
        // 有宏的工作簿保存为 xlsb 应正常工作
        var wb = Excel.Create(ExcelFormat.Xlsm);
        var prop = typeof(Workbook).GetProperty("VbaProjectBytes",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        prop!.SetValue(wb, new byte[] { 0x00, 0x01, 0x02 });

        var path = Path.Combine(Path.GetTempPath(), "liteexcel-macro-to-xlsb.xlsb");
        try
        {
            wb.SaveAs(path, ExcelFormat.Xlsb);
            Assert.True(File.Exists(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SaveAs_Xls_WithoutMacro_StillWorks()
    {
        // 无宏工作簿转 xls 不受影响
        var wb = Excel.Create(ExcelFormat.Xlsx);
        wb.Worksheets["Sheet1"].SetValue("A1", "无宏");
        var path = Path.Combine(Path.GetTempPath(), "liteexcel-nomacro.xls");
        wb.SaveAs(path, ExcelFormat.Xls);
        Assert.True(File.Exists(path));
        File.Delete(path);
    }

    [Fact]
    public void Save_StreamToCsv_WithMultipleSheets_Throws()
    {
        var wb = Excel.Create(ExcelFormat.Xlsx);
        wb.Worksheets.Add("第二张");
        using var ms = new MemoryStream();
        var ex = Assert.Throws<NotSupportedException>(() => wb.Save(ms, ExcelFormat.Csv));
        Assert.Contains("单工作表", ex.Message);
    }
}
