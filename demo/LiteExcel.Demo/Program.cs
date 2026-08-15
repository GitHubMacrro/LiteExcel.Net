using LiteExcel;

namespace LiteExcel.Demo;

public static class Program
{
    public static void Main()
    {
        var outDir = Path.Combine(AppContext.BaseDirectory, "Output");
        Directory.CreateDirectory(outDir);

        Console.WriteLine("=== LiteExcel Demo ===\n");

        Demo1_BasicWriteRead(outDir);
        Demo2_MultiSheet(outDir);
        Demo3_Styles(outDir);
        Demo4_MergedCells(outDir);
        Demo5_AutoFilter(outDir);
        Demo6_ListMapping(outDir);
        Demo7_DataTable(outDir);
        Demo8_StreamRows(outDir);
        Demo9_SpecialChars(outDir);
        Demo10_RowColStyles(outDir);
        Demo11_StreamReadWrite(outDir);
        Demo12_Comments(outDir);
        Demo13_RowHeightAutoWidth(outDir);
        Demo14_AppendData(outDir);
        Demo15_DataValidation(outDir);
        Demo16_ProgressCallback(outDir);
        Demo17_PublicApi(outDir);

        Console.WriteLine("\n=== All demos completed! ===");
        Console.WriteLine($"Output files in: {outDir}");
    }

    // 1. Basic write and read
    private static void Demo1_BasicWriteRead(string dir)
    {
        Console.WriteLine("[1] Basic Write + Read");

        var sheet = new SheetData
        {
            SheetName = "Employees",
            Headers = new() { "Name", "Age", "Salary", "Birthday", "Active" },
            Rows = new()
            {
                new Cell[]
                {
                    Cell.FromText("Zhang San"),
                    Cell.FromNumber(25),
                    Cell.FromNumber(8500.50, "#,##0.00"),
                    Cell.FromDate(new DateTime(2000, 1, 15)),
                    Cell.FromBoolean(true),
                },
                new Cell[]
                {
                    Cell.FromText("Li Si"),
                    Cell.FromNumber(30),
                    Cell.FromNumber(12000.00, "#,##0.00"),
                    Cell.FromDate(new DateTime(1995, 6, 20)),
                    Cell.FromBoolean(false),
                },
                new Cell[]
                {
                    Cell.FromText("Wang Wu"),
                    Cell.FromNumber(28),
                    Cell.FromNumber(9500.75, "#,##0.00"),
                    Cell.FromDate(new DateTime(1997, 3, 10)),
                    Cell.FromBoolean(true),
                },
            },
            FreezeHeader = true,
            ColumnWidths = new() { 15, 8, 12, 14, 8 },
        };

        var file = Path.Combine(dir, "01_basic.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote {read.Rows.Count} rows to: {Path.GetFileName(file)}");
        Console.WriteLine($"  Headers: {string.Join(", ", read.Headers)}");
        foreach (var row in read.Rows)
        {
            Console.WriteLine($"  {row[0].Text} | {row[1].Number} | {row[2].Number:F2} | {row[3].Date:yyyy-MM-dd} | {row[4].Boolean}");
        }
        Console.WriteLine();
    }

    // 2. Multi-sheet
    private static void Demo2_MultiSheet(string dir)
    {
        Console.WriteLine("[2] Multi-Sheet");

        var sheets = new List<SheetData>
        {
            new()
            {
                SheetName = "Q1",
                Headers = new() { "Month", "Revenue" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Jan"), Cell.FromNumber(45000) },
                    new Cell[] { Cell.FromText("Feb"), Cell.FromNumber(52000) },
                    new Cell[] { Cell.FromText("Mar"), Cell.FromNumber(48000) },
                },
            },
            new()
            {
                SheetName = "Q2",
                Headers = new() { "Month", "Revenue" },
                Rows = new()
                {
                    new Cell[] { Cell.FromText("Apr"), Cell.FromNumber(55000) },
                    new Cell[] { Cell.FromText("May"), Cell.FromNumber(61000) },
                },
            },
        };

        var file = Path.Combine(dir, "02_multisheet.xlsx");
        XlsxWriter.Write(file, sheets);

        var names = XlsxReader.GetSheetNames(file);
        Console.WriteLine($"  Sheets: {string.Join(", ", names)}");

        var all = XlsxReader.ReadAll(file);
        foreach (var s in all)
        {
            Console.WriteLine($"  {s.SheetName}: {s.Rows.Count} data rows");
        }
        Console.WriteLine();
    }

    // 3. Styles
    private static void Demo3_Styles(string dir)
    {
        Console.WriteLine("[3] Styles");

        var headerStyle = new CellStyle
        {
            Bold = true,
            FontColor = "#FFFFFF",
            FillColor = "#4472C4",
            HorizontalAlignment = HorizontalAlignment.Center,
            FontSize = 12,
        };

        var highlightStyle = new CellStyle
        {
            FillColor = "#FFC000",
            Bold = true,
            HorizontalAlignment = HorizontalAlignment.Right,
        };

        var borderStyle = new CellStyle
        {
            Border = new BorderStyle
            {
                Top = new BorderEdge { Style = "thin", Color = "#000000" },
                Bottom = new BorderEdge { Style = "thin", Color = "#000000" },
                Left = new BorderEdge { Style = "thin", Color = "#000000" },
                Right = new BorderEdge { Style = "thin", Color = "#000000" },
            },
        };

        var sheet = new SheetData
        {
            SheetName = "Styled",
            Headers = new() { "Product", "Price", "Status" },
            HeaderStyle = headerStyle,
            Rows = new()
            {
                new Cell[]
                {
                    new() { Type = CellType.Text, Text = "Laptop", Style = borderStyle },
                    new() { Type = CellType.Number, Number = 5999, NumberFormat = "#,##0", Style = borderStyle },
                    new() { Type = CellType.Text, Text = "In Stock", Style = highlightStyle },
                },
                new Cell[]
                {
                    new() { Type = CellType.Text, Text = "Mouse", Style = borderStyle },
                    new() { Type = CellType.Number, Number = 49.9, NumberFormat = "0.00", Style = borderStyle },
                    new() { Type = CellType.Text, Text = "In Stock", Style = borderStyle },
                },
            },
            ColumnWidths = new() { 15, 10, 12 },
        };

        var file = Path.Combine(dir, "03_styles.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote styled sheet: {Path.GetFileName(file)}");
        Console.WriteLine($"  Row 0 Col 2 has fill: {read.Rows[0][2].Style?.FillColor}");
        Console.WriteLine($"  Row 0 Col 0 has border: {read.Rows[0][0].Style?.Border?.Top?.Style}");
        Console.WriteLine();
    }

    // 4. Merged cells
    private static void Demo4_MergedCells(string dir)
    {
        Console.WriteLine("[4] Merged Cells");

        var titleStyle = new CellStyle
        {
            Bold = true,
            FontSize = 14,
            FillColor = "#4472C4",
            FontColor = "#FFFFFF",
            HorizontalAlignment = HorizontalAlignment.Center,
        };

        var sheet = new SheetData
        {
            SheetName = "Merge",
            Headers = new() { "Region", "Q1", "Q2", "Q3", "Q4" },
            Rows = new()
            {
                new Cell[]
                {
                    new() { Type = CellType.Text, Text = "East", Style = titleStyle },
                    Cell.FromNumber(100), Cell.FromNumber(120), Cell.FromNumber(110), Cell.FromNumber(130),
                },
                new Cell[]
                {
                    new() { Type = CellType.Text, Text = "West", Style = titleStyle },
                    Cell.FromNumber(80), Cell.FromNumber(90), Cell.FromNumber(85), Cell.FromNumber(95),
                },
            },
            MergedRanges = new()
            {
                new CellRange(0, 0, 0, 0), // A1 (title cell)
            },
        };

        var file = Path.Combine(dir, "04_merged.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote with {read.MergedRanges.Count} merge ranges: {Path.GetFileName(file)}");
        Console.WriteLine();
    }

    // 5. AutoFilter
    private static void Demo5_AutoFilter(string dir)
    {
        Console.WriteLine("[5] AutoFilter");

        var sheet = new SheetData
        {
            SheetName = "Filter",
            Headers = new() { "Name", "City", "Score" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Zhang"), Cell.FromText("Beijing"), Cell.FromNumber(85) },
                new Cell[] { Cell.FromText("Li"), Cell.FromText("Shanghai"), Cell.FromNumber(72) },
                new Cell[] { Cell.FromText("Wang"), Cell.FromText("Beijing"), Cell.FromNumber(90) },
                new Cell[] { Cell.FromText("Zhao"), Cell.FromText("Guangzhou"), Cell.FromNumber(65) },
                new Cell[] { Cell.FromText("Sun"), Cell.FromText("Beijing"), Cell.FromNumber(78) },
            },
            Filter = new AutoFilter
            {
                Range = "A1:C6",
                Columns = new()
                {
                    new FilterColumn
                    {
                        ColumnIndex = 1,
                        Type = FilterType.Equals,
                        Values = new() { "Beijing" },
                    },
                },
            },
        };

        var file = Path.Combine(dir, "05_filter.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote filter: {Path.GetFileName(file)}");
        Console.WriteLine($"  Filter range: {read.Filter?.Range}");
        Console.WriteLine($"  Hidden rows: {string.Join(", ", read.Filter?.HiddenRows ?? new HashSet<int>())}");
        Console.WriteLine("  (Open in Excel to see filter dropdown)");
        Console.WriteLine();
    }

    // 6. List<T> mapping
    private static void Demo6_ListMapping(string dir)
    {
        Console.WriteLine("[6] List<T> Mapping");

        var products = new List<Product>
        {
            new() { Code = "P001", Name = "Keyboard", Price = 199.50m, CreatedAt = new DateTime(2024, 1, 15), Stock = 100 },
            new() { Code = "P002", Name = "Mouse", Price = 49.99m, CreatedAt = new DateTime(2024, 2, 20), Stock = 200 },
            new() { Code = "P003", Name = "Monitor", Price = 1299.00m, CreatedAt = new DateTime(2024, 3, 10), Stock = 50 },
        };

        var file = Path.Combine(dir, "06_listmapping.xlsx");
        XlsxWriter.Write(file, products, opt =>
        {
            opt.SheetName = "Products";
            opt.FreezeHeader = true;
        });

        var read = XlsxReader.Read<Product>(file);
        Console.WriteLine($"  Wrote {read.Count} products: {Path.GetFileName(file)}");
        foreach (var p in read)
        {
            Console.WriteLine($"  {p.Code} | {p.Name} | {p.Price:C} | {p.CreatedAt:yyyy-MM-dd} | Stock={p.Stock}");
        }
        Console.WriteLine();
    }

    // 7. DataTable
    private static void Demo7_DataTable(string dir)
    {
        Console.WriteLine("[7] DataTable");

        var dt = new DataTable("Orders");
        dt.Columns.Add("OrderID", typeof(int));
        dt.Columns.Add("Customer", typeof(string));
        dt.Columns.Add("Amount", typeof(decimal));
        dt.Columns.Add("OrderDate", typeof(DateTime));

        dt.Rows.Add(1001, "Alice", 599.99m, new DateTime(2024, 6, 1));
        dt.Rows.Add(1002, "Bob", 1299.50m, new DateTime(2024, 6, 15));
        dt.Rows.Add(1003, "Charlie", 49.99m, new DateTime(2024, 7, 1));

        var file = Path.Combine(dir, "07_datatable.xlsx");
        XlsxWriter.Write(file, dt, "Orders");

        var read = XlsxReader.ReadAsDataTable(file, 0);
        Console.WriteLine($"  Wrote {read.Rows.Count} orders: {Path.GetFileName(file)}");
        foreach (DataRow r in read.Rows)
        {
            Console.WriteLine($"  #{r["OrderID"]} | {r["Customer"]} | {r["Amount"]:C} | {((DateTime)r["OrderDate"]):yyyy-MM-dd}");
        }
        Console.WriteLine();
    }

    // 8. Stream rows (large file)
    private static void Demo8_StreamRows(string dir)
    {
        Console.WriteLine("[8] Stream Rows (Large File)");

        var rows = new List<IReadOnlyList<Cell>>(1000);
        for (int i = 0; i < 1000; i++)
        {
            rows.Add(new Cell[]
            {
                Cell.FromNumber(i + 1),
                Cell.FromText($"Item-{i:D4}"),
                Cell.FromNumber(Math.Round(Random.Shared.NextDouble() * 1000, 2)),
            });
        }
        var sheet = new SheetData
        {
            SheetName = "LargeData",
            Headers = new() { "ID", "Name", "Value" },
            Rows = rows,
        };

        var file = Path.Combine(dir, "08_large.xlsx");
        XlsxWriter.Write(file, sheet);

        Console.WriteLine($"  Wrote 1000 rows: {Path.GetFileName(file)}");

        int count = 0;
        double sum = 0;
        XlsxReader.StreamRows(file, "LargeData", row =>
        {
            count++;
            sum += row[2].Number;
        });
        Console.WriteLine($"  Streamed {count} rows, sum of values = {sum:F2}");
        Console.WriteLine();
    }

    // 9. Special characters
    private static void Demo9_SpecialChars(string dir)
    {
        Console.WriteLine("[9] Special Characters");

        var sheet = new SheetData
        {
            SheetName = "Special",
            Headers = new() { "Text", "Value" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Hello & <World>"), Cell.FromNumber(1) },
                new Cell[] { Cell.FromText("Quote: \"test\""), Cell.FromNumber(2) },
                new Cell[] { Cell.FromText("中文测试"), Cell.FromNumber(3) },
                new Cell[] { Cell.FromText("Emoji: 🎉🚀"), Cell.FromNumber(4) },
                new Cell[] { Cell.FromText("  Leading & trailing  "), Cell.FromNumber(5) },
                new Cell[] { Cell.FromText("Tab\there"), Cell.FromNumber(6) },
                new Cell[] { Cell.FromText("Newline\nhere"), Cell.FromNumber(7) },
            },
        };

        var file = Path.Combine(dir, "09_special.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote {read.Rows.Count} rows with special chars: {Path.GetFileName(file)}");
        foreach (var row in read.Rows)
        {
            Console.WriteLine($"  [{row[1].Number}] {row[0].Text}");
        }
        Console.WriteLine();
    }
    // 10. Row/Column/Default styles
    private static void Demo10_RowColStyles(string dir)
    {
        Console.WriteLine("[10] Row/Column/Default Styles");

        var sheet = new SheetData
        {
            SheetName = "RowColStyle",
            Headers = new() { "Name", "Score", "Grade" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95), Cell.FromText("A") },
                new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(72), Cell.FromText("C") },
                new Cell[]
                {
                    Cell.FromText("Charlie"),
                    Cell.FromNumber(55),
                    new() { Type = CellType.Text, Text = "F", Style = new CellStyle { Bold = true, FontColor = "#FF0000" } },
                },
            },
            // 全表默认: 11 号字
            DefaultStyle = new CellStyle { FontSize = 11 },
            // 列级: Grade 列居中
            ColumnStyles = new()
            {
                { 2, new CellStyle { HorizontalAlignment = HorizontalAlignment.Center } },
            },
            // 行级: Bob 这行黄底
            RowStyles = new()
            {
                { 1, new CellStyle { FillColor = "#FFFF00" } },
            },
            ColumnWidths = new() { 12, 8, 8 },
        };

        var file = Path.Combine(dir, "10_rowcol_styles.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Wrote: {Path.GetFileName(file)}");
        Console.WriteLine($"  Row 1 (Bob) has yellow fill: {read.Rows[1][0].Style?.FillColor}");
        Console.WriteLine($"  Row 2 Col 2 (Charlie F) has red font: {read.Rows[2][2].Style?.FontColor}");
        Console.WriteLine("  (Open in Excel to see row/column/default styles)");
        Console.WriteLine();
    }

    // 11. Stream read/write
    private static void Demo11_StreamReadWrite(string dir)
    {
        Console.WriteLine("[11] Stream Read/Write");

        var sheet = new SheetData
        {
            SheetName = "Stream",
            Headers = new() { "ID", "Name" },
            Rows = new()
            {
                new Cell[] { Cell.FromNumber(1), Cell.FromText("Alice") },
                new Cell[] { Cell.FromNumber(2), Cell.FromText("Bob") },
            },
        };

        using var ms = new MemoryStream();
        XlsxWriter.Write(ms, sheet);
        ms.Position = 0;
        var read = XlsxReader.Read(ms, 0);

        var file = Path.Combine(dir, "11_stream.xlsx");
        using var fs = File.Create(file);
        XlsxWriter.Write(fs, sheet);

        Console.WriteLine($"  MemoryStream round-trip: {read.Rows.Count} rows");
        Console.WriteLine($"  File: {Path.GetFileName(file)}");
        Console.WriteLine();
    }

    // 12. Comments
    private static void Demo12_Comments(string dir)
    {
        Console.WriteLine("[12] Comments");

        var sheet = new SheetData
        {
            SheetName = "Comments",
            Headers = new() { "A", "B" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
            },
            Comments = new()
            {
                { "A1", "This is a comment on A1" },
                { "B1", "Note for B1 <with special chars>" },
            },
        };

        var file = Path.Combine(dir, "12_comments.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Comments: {read.Comments?.Count ?? 0}");
        Console.WriteLine($"  A1 comment: {read.Comments?["A1"]}");
        Console.WriteLine();
    }

    // 13. Row height + auto column width
    private static void Demo13_RowHeightAutoWidth(string dir)
    {
        Console.WriteLine("[13] Row Height + Auto Column Width");

        var sheet = new SheetData
        {
            SheetName = "RowCol",
            Headers = new() { "Short", "A Longer Header", "中文标题" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("a"), Cell.FromText("hello world"), Cell.FromText("你好") },
                new Cell[] { Cell.FromText("b"), Cell.FromText("hi"), Cell.FromText("世界") },
            },
            RowHeights = new() { { 0, 30.0 } },
        };
        XlsxWriter.AutoColumnWidths(sheet);

        var file = Path.Combine(dir, "13_rowcol.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Row 0 height: {read.RowHeights?[0]}");
        Console.WriteLine($"  Column widths: {string.Join(", ", sheet.ColumnWidths?.Select(w => w.ToString("F1")) ?? Array.Empty<string>())}");
        Console.WriteLine();
    }

    // 14. Append data
    private static void Demo14_AppendData(string dir)
    {
        Console.WriteLine("[14] Append Data");

        var file = Path.Combine(dir, "14_append.xlsx");
        if (File.Exists(file)) File.Delete(file);

        var sheet1 = new SheetData
        {
            SheetName = "Data",
            Headers = new() { "ID" },
            Rows = new()
            {
                new Cell[] { Cell.FromNumber(1) },
                new Cell[] { Cell.FromNumber(2) },
                new Cell[] { Cell.FromNumber(3) },
            },
        };
        XlsxWriter.Write(file, sheet1);

        var appendData = new SheetData
        {
            SheetName = "Data",
            Headers = new() { "ID" },
            Rows = new()
            {
                new Cell[] { Cell.FromNumber(4) },
                new Cell[] { Cell.FromNumber(5) },
            },
        };
        XlsxWriter.Append(file, appendData);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  After append: {read.Rows.Count} rows (expected 5)");
        Console.WriteLine();
    }

    // 15. Data validation (dropdown list)
    private static void Demo15_DataValidation(string dir)
    {
        Console.WriteLine("[15] Data Validation (Dropdown List)");

        var sheet = new SheetData
        {
            SheetName = "Validation",
            Headers = new() { "Name", "Department", "Score" },
            Rows = new()
            {
                new Cell[] { Cell.FromText("Alice"), Cell.FromText("IT"), Cell.FromNumber(85) },
                new Cell[] { Cell.FromText("Bob"), Cell.FromText("HR"), Cell.FromNumber(72) },
            },
            Validations = new()
            {
                new DataValidation
                {
                    Type = DataValidationType.List,
                    Sqref = "B2:B100",
                    Formula1 = "\"IT,HR,Finance,Sales\"",
                    AllowBlank = true,
                    PromptTitle = "Department",
                    Prompt = "Select a department from the list",
                },
                new DataValidation
                {
                    Type = DataValidationType.WholeNumber,
                    Sqref = "C2:C100",
                    Formula1 = "0",
                    Formula2 = "100",
                    AllowBlank = false,
                },
            },
        };

        var file = Path.Combine(dir, "15_validation.xlsx");
        XlsxWriter.Write(file, sheet);

        var read = XlsxReader.Read(file, 0);
        Console.WriteLine($"  Validations: {read.Validations?.Count ?? 0}");
        Console.WriteLine($"  B col dropdown: {read.Validations?[0].Formula1}");
        Console.WriteLine($"  C col type: {read.Validations?[1].Type}");
        Console.WriteLine("  (Open in Excel, click B2 to see dropdown)");
        Console.WriteLine();
    }

    // 16. Progress callback
    private static void Demo16_ProgressCallback(string dir)
    {
        Console.WriteLine("[16] Progress Callback");

        var rows = new List<IReadOnlyList<Cell>>(500);
        for (int i = 0; i < 500; i++)
        {
            rows.Add(new Cell[] { Cell.FromNumber(i + 1), Cell.FromText($"Item-{i:D3}") });
        }
        var sheet = new SheetData
        {
            SheetName = "Progress",
            Headers = new() { "ID", "Name" },
            Rows = rows,
        };

        var file = Path.Combine(dir, "16_progress.xlsx");
        XlsxWriter.Write(file, sheet);

        int lastReported = 0;
        XlsxReader.ReadWithProgress(file, 0, (current, total) =>
        {
            if (current == 1 || current == total || current % 100 == 0)
            {
                Console.WriteLine($"  Progress: {current}/{total} ({current * 100 / total}%)");
                lastReported = current;
            }
        });
        Console.WriteLine($"  Done! Total rows read: {lastReported}");
        Console.WriteLine();
    }

    // 17. 公开 API（Excel 门面 + Workbook/Worksheet/Cell/Range/Cells）
    private static void Demo17_PublicApi(string dir)
    {
        Console.WriteLine("[17] Public API (Excel facade)");

        // 新建工作簿，用自然层级写数据
        var wb = Excel.Create();
        var ws = wb.Worksheets["Sheet1"];
        wb.Properties.Created = DateTime.Now;
        wb.Properties.Title = nameof(Demo17_PublicApi);
        wb.Properties.Application = "自定义";
        wb.Properties.LastModifiedBy = "DemoAdmin";
        wb.Properties.Creator = "JackZ";

        ws.Name = "员工";
        ws.SetValue("A1", "姓名");
        ws.SetValue("B1", "年龄");
        ws.SetValue("C1", "工资");
        ws.SetValue("A2", "张三");
        ws.SetValue("B2", 25);
        ws.SetValue("C2", 8500.50);
        ws.SetValue("A3", "李四");
        ws.SetValue("B3", 30);
        ws.SetValue("C3", 12000.00);

        // Range 批量样式 + 合并
        ws.Range("A1:C1").Style = new CellStyle { Bold = true, FillColor = "#D9E1F2" };
        ws.Merge("A5:C5");
        ws.SetValue("A5", "合计区域示例");
        

        var file = Path.Combine(dir, "17_highlevel.xlsx");
        wb.SaveAs(file);
        Console.WriteLine($"  Written: {file}");

        // 打开并读取
        var opened = Excel.Open(file);
        var sheet = opened.Worksheets["员工"];
        Console.WriteLine($"  Sheet: {sheet.Name}, rows={sheet.RowCount}, cols={sheet.MaxColumn}");
        Console.WriteLine($"  A2 = {sheet.Cell("A2").GetString()}");
        Console.WriteLine($"  B2 = {sheet.Cell(2, 2).GetDouble()}");
        Console.WriteLine($"  C2 = {sheet.Cells["C2"].GetDouble()}");

        // 修改并保存
        sheet.SetValue("B2", 26);
        opened.Save();
        Console.WriteLine($"  Updated B2=26 and saved");

        // 流式写入大文件
        var streamFile = Path.Combine(dir, "17_stream.xlsx");
        using (var writer = Excel.CreateWriter(streamFile))
        {
            writer.WriteRow(new object?[] { "ID", "Name" });
            for (int i = 1; i <= 100; i++)
                writer.WriteRow(new object?[] { i, $"Row{i}" });
        }
        Console.WriteLine($"  StreamWriter: {streamFile}");

        // CSV
        var csvFile = Path.Combine(dir, "17_highlevel.csv");
        var csvWb = Excel.Create(ExcelFormat.Csv);
        csvWb.Worksheets["Sheet1"].SetValue("A1", "名称");
        csvWb.Worksheets["Sheet1"].SetValue("A2", "苹果");
        csvWb.SaveAs(csvFile);
        Console.WriteLine($"  CSV: {csvFile}");

        Console.WriteLine();
    }
}

// Demo model for List<T> mapping
public class Product
{
    [LiteColumn(Name = "Code", Order = 0)]
    public string Code { get; set; } = "";

    [LiteColumn(Name = "Product Name", Order = 1)]
    public string Name { get; set; } = "";

    [LiteColumn(Name = "Price", Order = 2, Format = "#,##0.00")]
    public decimal Price { get; set; }

    [LiteColumn(Name = "Created", Order = 3, Format = "yyyy-MM-dd")]
    public DateTime CreatedAt { get; set; }

    [LiteColumn(Order = 4)]
    public int Stock { get; set; }

    [LiteColumn(Ignore = true)]
    public string? InternalRemark { get; set; }
}
