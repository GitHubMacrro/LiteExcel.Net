using LiteExcel;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace LiteExcel.Tests;

public class RichTextTests
{
    private static string GetTempFile(string ext) =>
        Path.Combine(Path.GetTempPath(), $"richtext_{Guid.NewGuid():N}{ext}");

    private static RichText TwoRuns()
    {
        var rich = new RichText();
        rich.Add("粗体", new CellStyle { Bold = true, FontColor = "#FF0000" });
        rich.Add("普通");
        return rich;
    }

    [Fact]
    public void RichText_TextConcatenatesRuns()
    {
        var rich = TwoRuns();
        Assert.Equal("粗体普通", rich.Text);
        Assert.Equal(2, rich.Runs.Count);
    }

    [Fact]
    public void SetValue_RichText_SetsTypeAndText()
    {
        var cell = new Cell();
        cell.SetValue(TwoRuns());
        Assert.Equal(CellType.Text, cell.Type);
        Assert.Equal("粗体普通", cell.Text);
        Assert.NotNull(cell.RichText);
    }

    [Fact]
    public void SetValue_PlainString_ClearsRichText()
    {
        var cell = new Cell();
        cell.SetValue(TwoRuns());
        cell.SetValue("纯文本");
        Assert.Null(cell.RichText);
        Assert.Equal("纯文本", cell.Text);
    }

    [Fact]
    public void Copy_RichTextIsDeepCopied()
    {
        var a = new Cell();
        a.SetValue(TwoRuns());
        var b = new Cell();
        b.SetValue(a);

        Assert.NotNull(b.RichText);
        Assert.NotSame(a.RichText, b.RichText);
        b.RichText!.Runs[0].Text = "改过";
        Assert.Equal("粗体", a.RichText!.Runs[0].Text);
    }

    [Theory]
    [InlineData(ExcelFormat.Xlsx, ".xlsx")]
    [InlineData(ExcelFormat.Xlsm, ".xlsm")]
    [InlineData(ExcelFormat.Xlsb, ".xlsb")]
    [InlineData(ExcelFormat.Xls, ".xls")]
    public void RichText_RoundTrip_AllFormats(ExcelFormat format, string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(format: format);
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(TwoRuns());
            wb.SaveAs(file, format);

            var opened = Excel.Open(file, format);
            var cell = opened.Worksheets[0].Cell("A1");

            Assert.Equal("粗体普通", cell.Text);
            Assert.NotNull(cell.RichText);
            Assert.Equal(2, cell.RichText!.Runs.Count);
            Assert.Equal("粗体", cell.RichText.Runs[0].Text);
            Assert.Equal("普通", cell.RichText.Runs[1].Text);
            Assert.True(cell.RichText.Runs[0].Style?.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_XlsxWritesRunProperties()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(TwoRuns());
            wb.SaveAs(file);

            var xml = ReadSheetXml(file);
            Assert.Contains("<rPr>", xml);
            Assert.Contains("<b/>", xml);
            Assert.Contains("FF0000", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_PlainTextWritesNoRun()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "纯文本");
            wb.SaveAs(file);

            Assert.DoesNotContain("<rPr>", ReadSheetXml(file));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_Superscript_RoundTrip()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var rich = new RichText();
            rich.Add("x");
            rich.Add("2", verticalAlign: TextVerticalAlign.Superscript);

            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(rich);
            wb.SaveAs(file);

            Assert.Contains("vertAlign val=\"superscript\"", ReadSheetXml(file));

            var cell = Excel.Open(file).Worksheets[0].Cell("A1");
            Assert.Equal(TextVerticalAlign.Superscript, cell.RichText!.Runs[1].VerticalAlign);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_SkipsPhoneticRuns()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "漢字");
            wb.SaveAs(file);

            var cell = Excel.Open(file).Worksheets[0].Cell("A1");
            Assert.Equal("漢字", cell.Text);
            Assert.Null(cell.RichText);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_CoexistsWithCellStyle()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var cell = wb.Worksheets["Sheet1"].Cell("A1");
            cell.SetValue(TwoRuns());
            cell.Style = new CellStyle { Italic = true, FillColor = "#FFFF00" };
            wb.SaveAs(file);

            var opened = Excel.Open(file).Worksheets[0].Cell("A1");
            Assert.True(opened.Style?.Italic);
            Assert.True(opened.RichText!.Runs[0].Style?.Bold);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_CellCopy_KeepsRuns()
    {
        var src = new Cell();
        src.SetValue(TwoRuns());
        var dst = new Cell();
        dst.SetValue(src);

        Assert.Equal(2, dst.RichText!.Runs.Count);
        Assert.Equal("粗体", dst.RichText.Runs[0].Text);
        Assert.NotSame(src.RichText, dst.RichText);
    }

    [Fact]
    public void RichText_Csv_ReportsDegradation()
    {
        var file = GetTempFile(".csv");
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(TwoRuns());

            var reported = new List<DegradationCapability>();
            Excel.Write(file, wb, new ExcelWriteOptions
            {
                OnDegradation = d => reported.Add(d.Capability),
            });

            Assert.Contains(DegradationCapability.RichText, reported);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void RichText_StreamWriter_WritesRuns()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            using (var w = Excel.CreateWriter(file))
            {
                w.WriteRow(new[] { new Cell { RichText = TwoRuns() } });
                w.Close();
            }

            Assert.Contains("<rPr>", ReadSheetXml(file));
            var cell = Excel.Open(file).Worksheets[0].Cell("A1");
            Assert.Equal(2, cell.RichText!.Runs.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SettingText_ClearsStaleRichText()
    {
        var file = GetTempFile(".xlsx");
        try
        {
            var wb = Excel.Create();
            var cell = wb.Worksheets["Sheet1"].Cell("A1");
            cell.SetValue(TwoRuns());
            cell.Text = "新文本";

            Assert.Null(cell.RichText);

            wb.SaveAs(file);
            var back = Excel.Open(file).Worksheets[0].Cell("A1");
            Assert.Equal("新文本", back.Text);
            Assert.Null(back.RichText);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void SettingSameText_KeepsRichText()
    {
        var cell = new Cell();
        cell.SetValue(TwoRuns());
        cell.Text = "粗体普通";
        Assert.NotNull(cell.RichText);
        Assert.Equal(2, cell.RichText!.Runs.Count);
    }

    [Theory]
    [InlineData(ExcelFormat.Xlsx, ".xlsx")]
    [InlineData(ExcelFormat.Xlsb, ".xlsb")]
    [InlineData(ExcelFormat.Xls, ".xls")]
    public void OpenModifySave_KeepsRichText(ExcelFormat format, string ext)
    {
        var file = GetTempFile(ext);
        var file2 = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(format: format);
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(TwoRuns());
            wb.SaveAs(file, format);

            var reopened = Excel.Open(file, format);
            reopened.Worksheets[0].SetValue("Z9", "modified");
            reopened.SaveAs(file2, format);

            var cell = Excel.Open(file2, format).Worksheets[0].Cell("A1");
            Assert.Equal("粗体普通", cell.Text);
            Assert.Equal(2, cell.RichText!.Runs.Count);
            Assert.True(cell.RichText.Runs[0].Style?.Bold);
        }
        finally
        {
            foreach (var f in new[] { file, file2 }) if (File.Exists(f)) File.Delete(f);
        }
    }

    [Fact]
    public void LongTextWithManyRuns_RoundTrips()
    {
        var file = GetTempFile(".xls");
        try
        {
            var rich = new RichText();
            rich.Add(new string('A', 2000), new CellStyle { Bold = true });
            rich.Add(new string('B', 2000));

            var wb = Excel.Create(format: ExcelFormat.Xls);
            wb.Worksheets["Sheet1"].Cell("A1").SetValue(rich);
            wb.SaveAs(file, ExcelFormat.Xls);

            var cell = Excel.Open(file, ExcelFormat.Xls).Worksheets[0].Cell("A1");
            Assert.Equal(4000, cell.Text!.Length);
            Assert.Equal(2, cell.RichText!.Runs.Count);
            Assert.Equal(2000, cell.RichText.Runs[0].Text.Length);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    private static string ReadSheetXml(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        using var s = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var r = new StreamReader(s, Encoding.UTF8);
        return r.ReadToEnd();
    }
}
