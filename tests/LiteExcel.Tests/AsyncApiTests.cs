using LiteExcel;
using System.Data;
using System.IO;
using Xunit;

namespace LiteExcel.Tests;

/// <summary>
/// 异步 API（Task 包装层）回归：SaveAsAsync/OpenAsync 往返、四格式路由、
/// Write/Read/ReadSheet/ReadAsDataTable/GetSheetNames/StreamRows/Append/ReadWithProgress 异步版、
/// Stream 保存、取消语义、Workbook.SaveAsync。
/// </summary>
public class AsyncApiTests
{
    private static string GetTempFile(string ext = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"litexlsx_async_{Guid.NewGuid():N}{ext}");

    private class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    [Fact]
    public async Task SaveAsAsync_OpenAsync_RoundTrip()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "姓名");
            wb.Worksheets["Sheet1"].SetValue("B1", 25);
            await wb.SaveAsAsync(file);

            var opened = await Excel.OpenAsync(file);
            Assert.Equal("姓名", opened.Worksheets[0].Cell("A1").GetString());
            Assert.Equal(25.0, opened.Worksheets[0].Cell("B1").GetDouble());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Theory]
    [InlineData(".xlsb")]
    [InlineData(".xls")]
    [InlineData(".csv")]
    public async Task OpenAsync_RoutesByExtension_AllFormats(string ext)
    {
        var file = GetTempFile(ext);
        try
        {
            var wb = Excel.Create(format: ExcelFormat.Xlsx);
            wb.Worksheets["Sheet1"].SetValue("A1", "值");
            // WriteAsync 按扩展名推断目标格式（与同步 Write 一致）
            await Excel.WriteAsync(file, wb);

            var opened = await Excel.OpenAsync(file);
            Assert.Equal("值", opened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task WriteAsync_Workbook_SheetData_DataTable_List_AllOverloads()
    {
        var f1 = GetTempFile(); var f2 = GetTempFile(); var f3 = GetTempFile(); var f4 = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "wb");
            await Excel.WriteAsync(f1, wb);

            var sheet = new SheetData { SheetName = "S" };
            sheet.Headers = new List<string> { "H" };
            sheet.Rows.Add(new Cell[] { Cell.FromText("sheet") });
            await Excel.WriteAsync(f2, sheet);

            var dt = new DataTable();
            dt.Columns.Add("C", typeof(string));
            dt.Rows.Add("table");
            await Excel.WriteAsync(f3, dt, "T");

            var list = new List<Person> { new Person { Name = "list", Age = 1 } };
            await Excel.WriteAsync(f4, list);

            Assert.Equal("wb", (await Excel.OpenAsync(f1)).Worksheets[0].Cell("A1").GetString());
            // SheetData.Headers 作为首行写出：A1=H，A2=sheet
            var s2 = (await Excel.OpenAsync(f2)).Worksheets[0];
            Assert.Equal("H", s2.Cell("A1").GetString());
            Assert.Equal("sheet", s2.Cell("A2").GetString());
            // DataTable 列名作为首行：A1=C，A2=table
            var s3 = (await Excel.OpenAsync(f3)).Worksheets["T"];
            Assert.Equal("C", s3.Cell("A1").GetString());
            Assert.Equal("table", s3.Cell("A2").GetString());
            Assert.Equal("list", (await Excel.OpenAsync(f4)).Worksheets[0].Cell("A2").GetString());
        }
        finally
        {
            foreach (var f in new[] { f1, f2, f3, f4 }) if (File.Exists(f)) File.Delete(f);
        }
    }

    [Fact]
    public async Task ReadAsync_ListMapping()
    {
        var file = GetTempFile();
        try
        {
            await Excel.WriteAsync(file, new List<Person>
            {
                new Person { Name = "张三", Age = 25 },
                new Person { Name = "李四", Age = 30 },
            });

            var list = await Excel.ReadAsync<Person>(file);
            Assert.Equal(2, list.Count);
            Assert.Equal("张三", list[0].Name);
            Assert.Equal(25, list[0].Age);
            Assert.Equal(30, list[1].Age);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task ReadAsync_FromStream()
    {
        var file = GetTempFile();
        try
        {
            await Excel.WriteAsync(file, new List<Person> { new Person { Name = "流", Age = 9 } });
            using var fs = File.OpenRead(file);
            var list = await Excel.ReadAsync<Person>(fs, ExcelFormat.Xlsx);
            Assert.Single(list);
            Assert.Equal("流", list[0].Name);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task ReadSheetAsync_ByIndex_And_ByName()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create(format: ExcelFormat.Xlsx);
            wb.Worksheets.Add("数据");
            wb.Worksheets["数据"].SetValue("A1", "表头");
            wb.Worksheets["数据"].SetValue("A2", "行");
            await wb.SaveAsAsync(file);

            var byIndex = await Excel.ReadSheetAsync(file, 1);
            var byName = await Excel.ReadSheetAsync(file, "数据");
            Assert.Equal("行", byIndex.Rows[0][0].Text);
            Assert.Equal("行", byName.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task ReadSheetAsync_FromStream()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "流表");
            await wb.SaveAsAsync(file);
            using var fs = File.OpenRead(file);
            var sheet = await Excel.ReadSheetAsync(fs, ExcelFormat.Xlsx, 0, firstRowIsHeader: false);
            Assert.Equal("流表", sheet.Rows[0][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task ReadAsDataTableAsync_Path_And_Stream()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "列");
            wb.Worksheets["Sheet1"].SetValue("A2", 1);
            await wb.SaveAsAsync(file);

            var dt1 = await Excel.ReadAsDataTableAsync(file);
            Assert.Equal(1, dt1.Rows.Count);

            using var fs = File.OpenRead(file);
            var dt2 = await Excel.ReadAsDataTableAsync(fs, ExcelFormat.Xlsx);
            Assert.Equal(1, dt2.Rows.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task GetSheetNamesAsync_Path_And_Stream()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets.Add("第二张");
            await wb.SaveAsAsync(file);

            var names = await Excel.GetSheetNamesAsync(file);
            Assert.Equal(new[] { "Sheet1", "第二张" }, names);

            using var fs = File.OpenRead(file);
            var names2 = await Excel.GetSheetNamesAsync(fs);
            Assert.Equal(2, names2.Count);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task StreamRowsAsync_CallbackReceivesDataRows()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "表头");
            for (int i = 1; i <= 3; i++) ws.SetValue($"A{i + 1}", $"行{i}");
            await wb.SaveAsAsync(file);

            int count = 0;
            await Excel.StreamRowsAsync(file, "Sheet1", row =>
            {
                count++;
                Assert.StartsWith("行", row[0].GetString());
            });
            Assert.Equal(3, count); // 自动跳过首行（表头）
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task ReadWithProgressAsync_Completes()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            var ws = wb.Worksheets["Sheet1"];
            ws.SetValue("A1", "表头");
            for (int i = 1; i <= 5; i++) ws.SetValue($"A{i + 1}", i);
            await wb.SaveAsAsync(file);

            int last = 0, total = 0;
            await Excel.ReadWithProgressAsync(file, 0, (current, t) => { last = current; total = t; });
            Assert.Equal(5, total);
            Assert.Equal(5, last);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task AppendAsync_AppendsRows()
    {
        var file = GetTempFile();
        try
        {
            var initial = new SheetData { SheetName = "Data" };
            initial.Headers = new List<string> { "Name" };
            initial.Rows.Add(new Cell[] { Cell.FromText("Alice") });
            await Excel.WriteAsync(file, initial);

            var more = new SheetData { SheetName = "Data" };
            more.Headers = new List<string> { "Name" };
            more.Rows.Add(new Cell[] { Cell.FromText("Bob") });
            await Excel.AppendAsync(file, more);

            var read = await Excel.ReadSheetAsync(file, 0);
            Assert.Equal(2, read.Rows.Count);
            Assert.Equal("Alice", read.Rows[0][0].Text);
            Assert.Equal("Bob", read.Rows[1][0].Text);
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task SaveAsync_Stream_ProducesReadableWorkbook()
    {
        using var ms = new MemoryStream();
        var wb = Excel.Create();
        wb.Worksheets["Sheet1"].SetValue("A1", "流保存");
        await wb.SaveAsync(ms, ExcelFormat.Xlsx);

        ms.Position = 0;
        var opened = await Excel.OpenAsync(ms, ExcelFormat.Xlsx);
        Assert.Equal("流保存", opened.Worksheets[0].Cell("A1").GetString());
    }

    [Fact]
    public async Task Workbook_SaveAsync_UsesCurrentPath()
    {
        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            wb.Worksheets["Sheet1"].SetValue("A1", "第一版");
            await wb.SaveAsAsync(file);

            var opened = await Excel.OpenAsync(file);
            opened.Worksheets[0].SetValue("A1", "第二版");
            await opened.SaveAsync();

            var reopened = await Excel.OpenAsync(file);
            Assert.Equal("第二版", reopened.Worksheets[0].Cell("A1").GetString());
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }

    [Fact]
    public async Task PreCancelledToken_Throws()
    {
        using var cts = new System.Threading.CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => Excel.GetSheetNamesAsync("whatever.xlsx", cts.Token));

        var file = GetTempFile();
        try
        {
            var wb = Excel.Create();
            await wb.SaveAsAsync(file);
            using var fs = File.OpenRead(file);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(
                () => Excel.OpenAsync(fs, ExcelFormat.Xlsx, cancellationToken: cts.Token));
        }
        finally { if (File.Exists(file)) File.Delete(file); }
    }
}
