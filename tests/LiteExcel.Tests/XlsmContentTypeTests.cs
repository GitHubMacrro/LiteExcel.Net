using LiteExcel;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// xlsm 写出的主文档内容类型必须为 macroEnabled，否则 Excel 以"格式或扩展名无效"拒绝打开（见 issue #1）。
/// </summary>
public class XlsmContentTypeTests
{
    private const string SheetMainType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string MacroMainType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_xlsm_{Guid.NewGuid():N}{ext}");

    /// <summary>读取包内 [Content_Types].xml 中 /xl/workbook.xml 的内容类型 </summary>
    private static string GetWorkbookContentType(string file)
    {
        using var zip = new ZipArchive(File.OpenRead(file), ZipArchiveMode.Read);
        var entry = zip.GetEntry("[Content_Types].xml")!;
        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        var ns = doc.Root!.GetDefaultNamespace();
        var overrideEl = doc.Root.Elements(ns + "Override")
            .First(o => (string?)o.Attribute("PartName") == "/xl/workbook.xml");
        return (string)overrideEl.Attribute("ContentType")!;
    }

    [Fact]
    public void SaveAsXlsm_WritesMacroEnabledMainType()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "x");
            wb.SaveAs(file);

            Assert.Equal(MacroMainType, GetWorkbookContentType(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SaveAsXlsx_StillWritesSheetMainType()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "x");
            wb.SaveAs(file);

            Assert.Equal(SheetMainType, GetWorkbookContentType(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void OpenXlsx_SaveAsXlsm_WritesMacroEnabledMainType()
    {
        var src = GetTempFile(".xlsx");
        var dst = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "v");
            wb.SaveAs(src);

            var opened = Excel.Open(src);
            opened.SaveAs(dst, ExcelFormat.Xlsm);

            Assert.Equal(MacroMainType, GetWorkbookContentType(dst));
        }
        finally { if (File.Exists(src)) File.Delete(src); if (File.Exists(dst)) File.Delete(dst); }
    }

    [Fact]
    public void XlsxWriter_WritePath_Xlsm_WritesMacroEnabledMainType()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            var sheet = new SheetData { SheetName = "Sheet1" };
            sheet.Rows.Add(new Cell[] { Cell.FromText("a") });
            XlsxWriter.Write(file, sheet);

            Assert.Equal(MacroMainType, GetWorkbookContentType(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void CreateWriter_Xlsm_WritesMacroEnabledMainType()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            using (var writer = Excel.CreateWriter(file))
            {
                writer.WriteRow(new object?[] { "a", 1 });
            }

            Assert.Equal(MacroMainType, GetWorkbookContentType(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsm_RoundTrip_OpenReadsBack()
    {
        var file = GetTempFile(".xlsm");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsm);
            wb.Worksheets["Sheet1"].SetValue("A1", "roundtrip");
            wb.SaveAs(file);

            var opened = Excel.Open(file);
            Assert.Equal(ExcelFormat.Xlsm, opened.Format);
            Assert.Equal("roundtrip", opened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
