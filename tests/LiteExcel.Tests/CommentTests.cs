using LiteExcel;

namespace LiteExcel.Tests;

public class CommentTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void Comment_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A", "B" },
                Rows = new() { new Cell[] { Cell.FromText("x"), Cell.FromText("y") } },
                Comments = new() { { "A1", "测试批注" } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Comments);
            Assert.True(read.Comments!.ContainsKey("A1"));
            Assert.Equal("测试批注", read.Comments["A1"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void MultipleComments_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "Name", "Age" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(30) },
                    new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(25) },
                },
                Comments = new()
                {
                    { "A1", "姓名列" },
                    { "B1", "年龄列" },
                    { "A2", "第一个人的名字" },
                    { "B3", "第二个人的年龄" },
                },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Comments);
            Assert.Equal(4, read.Comments!.Count);
            Assert.Equal("姓名列", read.Comments["A1"]);
            Assert.Equal("年龄列", read.Comments["B1"]);
            Assert.Equal("第一个人的名字", read.Comments["A2"]);
            Assert.Equal("第二个人的年龄", read.Comments["B3"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoComments_DoesNotCrash()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Null(read.Comments);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Theory]
    [InlineData("含<尖>括号")]
    [InlineData("含&符号")]
    [InlineData("含\"引号\"")]
    [InlineData("含'单引号'")]
    [InlineData("<>&\"'全部特殊字符")]
    public void SpecialChars_CommentRoundTrip(string comment)
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
                Comments = new() { { "A2", comment } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Comments);
            Assert.True(read.Comments!.ContainsKey("A2"));
            Assert.Equal(comment, read.Comments["A2"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Comments_MultiSheet_Independent()
    {
        var file = GetTempFile();
        try
        {
            var sheets = new List<SheetData>
            {
                new()
                {
                    SheetName = "Sheet1",
                    Headers = new() { "A" },
                    Rows = new() { new Cell[] { Cell.FromText("x") } },
                    Comments = new() { { "A1", "sheet1批注" } },
                },
                new()
                {
                    SheetName = "Sheet2",
                    Headers = new() { "B" },
                    Rows = new() { new Cell[] { Cell.FromText("y") } },
                    Comments = new() { { "A1", "sheet2批注" } },
                },
            };
            XlsxWriter.Write(file, sheets);

            var all = XlsxReader.ReadAll(file);
            Assert.Equal(2, all.Count);
            Assert.NotNull(all[0].Comments);
            Assert.Equal("sheet1批注", all[0].Comments!["A1"]);
            Assert.NotNull(all[1].Comments);
            Assert.Equal("sheet2批注", all[1].Comments!["A1"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Comment_WithLeadingTrailingSpaces()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Test",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
                Comments = new() { { "A1", "  前后有空格  " } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.NotNull(read.Comments);
            Assert.Equal("  前后有空格  ", read.Comments!["A1"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    private static string GetCommentFixturePath(string ext)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Fixtures", $"excel-authored-comments-filter.{ext}");
        Assert.True(File.Exists(path), $"Required fixture is missing: {path}");
        return path;
    }

    [Fact]
    public void Xlsb_Comments_RoundTrip()
    {
        var file = Path.Combine(Path.GetTempPath(), $"litexlsx_xlsb_comment_{Guid.NewGuid():N}.xlsb");
        try
        {
            var wb = Excel.Create(ExcelFormat.Xlsb);
            var ws = wb.Worksheets[0];
            ws.SetValue("A1", "data");
            ws.Comments = new Dictionary<string, string> { { "A1", "xlsb comment text" } };

            var reported = new List<DegradationInfo>();
            Excel.Write(file, wb, new ExcelWriteOptions { OnDegradation = d => reported.Add(d) });

            Assert.DoesNotContain(reported, d => d.Capability == DegradationCapability.Comments);
            var opened = Excel.Open(file);
            Assert.NotNull(opened.Worksheets[0].Comments);
            Assert.True(opened.Worksheets[0].Comments!.ContainsKey("A1"));
            Assert.Equal("xlsb comment text", opened.Worksheets[0].Comments!["A1"]);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Xlsb_RealSample_ReadsComments()
    {
        var wb = Excel.Open(GetCommentFixturePath("xlsb"));
        var ws = wb.Worksheets[0];
        Assert.NotNull(ws.Comments);
        Assert.True(ws.Comments!.Count > 0);
        foreach (var kv in ws.Comments)
        {
            Assert.False(string.IsNullOrEmpty(kv.Value), $"Comment at {kv.Key} should have text");
        }
    }

    [Fact]
    public void Xlsx_RealSample_ReadsComments()
    {
        var wb = Excel.Open(GetCommentFixturePath("xlsx"));
        var ws = wb.Worksheets[0];
        Assert.NotNull(ws.Comments);
        Assert.True(ws.Comments!.Count > 0);
    }

    [Fact]
    public void Xls_RealSample_ReadsComments()
    {
        var wb = Excel.Open(GetCommentFixturePath("xls"));
        var ws = wb.Worksheets[0];
        Assert.NotNull(ws.Comments);
        Assert.True(ws.Comments!.Count > 0);
        foreach (var kv in ws.Comments!)
            Assert.False(string.IsNullOrEmpty(kv.Value), $"Comment at {kv.Key} should have text");
    }
}