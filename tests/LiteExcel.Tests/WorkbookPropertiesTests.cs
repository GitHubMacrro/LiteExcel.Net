using System.IO.Compression;
using LiteExcel;

namespace LiteExcel.Tests;

public class WorkbookPropertiesTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    [Fact]
    public void WriteWithProperties_ReadBack_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Sheet1",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };

            var props = new WorkbookProperties
            {
                Creator = "张三",
                LastModifiedBy = "李四",
                Created = new DateTime(2024, 6, 1, 12, 0, 0),
                Modified = new DateTime(2024, 6, 15, 18, 30, 0),
                Title = "测试标题",
                Subject = "测试主题",
            };

            XlsxWriter.Write(file, sheet, props);

            var read = XlsxReader.ReadProperties(file);
            Assert.NotNull(read);
            Assert.Equal("张三", read.Creator);
            Assert.Equal("李四", read.LastModifiedBy);
            Assert.Equal(new DateTime(2024, 6, 1, 12, 0, 0), read.Created);
            Assert.Equal(new DateTime(2024, 6, 15, 18, 30, 0), read.Modified);
            Assert.Equal("测试标题", read.Title);
            Assert.Equal("测试主题", read.Subject);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void WriteWithoutProps_ReadProperties_ReturnsEmpty()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "Sheet1",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("x") } },
            };

            // 不带 props 的旧重载
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.ReadProperties(file);
            Assert.NotNull(read);
            Assert.Null(read.Creator);
            Assert.Null(read.Title);
            Assert.Null(read.Application);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Stream_WriteWithProps_ReadBack()
    {
        using var ms = new MemoryStream();
        var sheet = new SheetData
        {
            SheetName = "S",
            Headers = new() { "A" },
            Rows = new() { new Cell[] { Cell.FromText("v") } },
        };
        var props = new WorkbookProperties { Creator = "Stream Author" };

        XlsxWriter.Write(ms, sheet, props);
        ms.Position = 0;

        var read = XlsxReader.ReadProperties(ms);
        Assert.Equal("Stream Author", read.Creator);
    }

    [Fact]
    public void Application_DefaultsToHostAssemblyName()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "S", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("v") } } };
            var props = new WorkbookProperties { Creator = "Author" };

            XlsxWriter.Write(file, sheet, props);

            // Application 未显式设置时，默认取宿主程序集名
            var read = XlsxReader.ReadProperties(file);
            Assert.NotNull(read.Application);
            Assert.False(string.IsNullOrEmpty(read.Application));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void Application_ExplicitValue_Preserved()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "S", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("v") } } };
            var props = new WorkbookProperties { Application = "MyCustomApp" };

            XlsxWriter.Write(file, sheet, props);

            var read = XlsxReader.ReadProperties(file);
            Assert.Equal("MyCustomApp", read.Application);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void WriteWithProps_ProducesDocPropsParts()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "S", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("v") } } };
            var props = new WorkbookProperties { Creator = "A", Title = "T" };

            XlsxWriter.Write(file, sheet, props);

            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
            Assert.NotNull(zip.GetEntry("docProps/core.xml"));
            Assert.NotNull(zip.GetEntry("docProps/app.xml"));
            Assert.NotNull(zip.GetEntry("[Content_Types].xml"));
            Assert.NotNull(zip.GetEntry("_rels/.rels"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void WithoutProps_NoDocPropsParts()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "S", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("v") } } };
            XlsxWriter.Write(file, sheet);

            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
            Assert.Null(zip.GetEntry("docProps/core.xml"));
            Assert.Null(zip.GetEntry("docProps/app.xml"));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void WriteWithProperties_MissingDates_AreFilledAutomatically()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData { SheetName = "S", Headers = new() { "A" }, Rows = new() { new Cell[] { Cell.FromText("v") } } };
            XlsxWriter.Write(file, sheet, new WorkbookProperties { Creator = "Author" });

            var read = XlsxReader.ReadProperties(file);
            Assert.Equal("Author", read.Creator);
            Assert.NotNull(read.Created);
            Assert.NotNull(read.Modified);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void AppProperties_ContainsActualSheetNames()
    {
        var file = GetTempFile();
        try
        {
            var sheets = new[]
            {
                new SheetData { SheetName = "One", Headers = new() { "A" } },
                new SheetData { SheetName = "Two", Headers = new() { "B" } },
            };
            XlsxWriter.Write(file, sheets, new WorkbookProperties { Creator = "Author" });

            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
            using var reader = new StreamReader(zip.GetEntry("docProps/app.xml")!.Open());
            var xml = reader.ReadToEnd();
            Assert.Contains("<vt:i4>2</vt:i4>", xml);
            Assert.Contains("<vt:lpstr>One</vt:lpstr>", xml);
            Assert.Contains("<vt:lpstr>Two</vt:lpstr>", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}

public class Gray125FillTests
{
    private static string GetTempFile() => Path.Combine(Path.GetTempPath(), $"litexlsx_{Guid.NewGuid():N}.xlsx");

    private static string ReadStylesXml(string file)
    {
        using var fs = new FileStream(file, FileMode.Open, FileAccess.Read);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry("xl/styles.xml")!;
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    [Fact]
    public void Fills_FirstTwo_AreNoneAndGray125()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "S",
                Headers = new() { "A", "B" },
                Rows = new()
                {
                    new Cell[]
                    {
                        new() { Type = CellType.Text, Text = "x", Style = new CellStyle { FillColor = "#00FF00" } },
                        Cell.FromText("y"),
                    },
                },
            };
            XlsxWriter.Write(file, sheet);

            var xml = ReadStylesXml(file);

            // 前两个 fill 必须是 none 和 gray125（Excel 规范）
            int nonePos = xml.IndexOf("patternType=\"none\"");
            int gray125Pos = xml.IndexOf("patternType=\"gray125\"");
            int solidPos = xml.IndexOf("patternType=\"solid\"");

            Assert.True(nonePos >= 0, "应有 patternType=none");
            Assert.True(gray125Pos >= 0, "应有 patternType=gray125");
            Assert.True(solidPos >= 0, "应有用户颜色的 patternType=solid");

            // none < gray125 < solid（用户颜色在保留填充之后）
            Assert.True(nonePos < gray125Pos, "none 应在 gray125 之前");
            Assert.True(gray125Pos < solidPos, "gray125 应在用户颜色之前");
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void FillColor_RoundTrip_StillWorks()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "S",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { new() { Type = CellType.Text, Text = "x", Style = new CellStyle { FillColor = "#FF0000" } } } },
            };
            XlsxWriter.Write(file, sheet);

            var read = XlsxReader.Read(file, 0);
            Assert.Equal("#FF0000", read.Rows[0][0].Style?.FillColor);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public void NoFillColor_NoSolidFillNeeded()
    {
        var file = GetTempFile();
        try
        {
            var sheet = new SheetData
            {
                SheetName = "S",
                Headers = new() { "A" },
                Rows = new() { new Cell[] { Cell.FromText("plain") } },
            };
            XlsxWriter.Write(file, sheet);

            var xml = ReadStylesXml(file);
            // 无填充色时不应有 solid
            Assert.DoesNotContain("patternType=\"solid\"", xml);
            // 但应有保留的 none + gray125
            Assert.Contains("patternType=\"none\"", xml);
            Assert.Contains("patternType=\"gray125\"", xml);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
