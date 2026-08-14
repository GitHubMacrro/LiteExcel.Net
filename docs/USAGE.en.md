# LiteExcel Usage Guide

**Version**: 2.1.3  
**Target Frameworks**: net48 + net8.0  
**Dependencies**: Zero third-party dependencies, .NET BCL only

---

## Table of Contents

1. [Installation & Reference](#1-installation--reference)
2. [Quick Start](#2-quick-start)
3. [Cells & Data Types](#3-cells--data-types)
4. [Reading](#4-reading)
5. [Writing](#5-writing)
6. [Styling](#6-styling)
7. [Merged Cells](#7-merged-cells)
8. [Auto Filter](#8-auto-filter)
9. [Row Height & Column Width](#9-row-height--column-width)
10. [Cell Comments](#10-cell-comments)
11. [Data Validation (Dropdown List)](#11-data-validation-dropdown-list)
12. [Appending Data](#12-appending-data)
13. [List<T> Mapping (reflection, not AOT compatible)](#13-listt-mapping-reflection-not-aot-compatible)
14. [DataTable Convenience API (AOT safe)](#14-datatable-convenience-api-aot-safe)
15. [Stream Read/Write](#15-stream-readwrite)
16. [Streaming Read & Progress Callback](#16-streaming-read--progress-callback)
17. [Document Properties (Author/Time/Title)](#17-document-properties-authortimetitle)
18. [Error Handling](#18-error-handling)
19. [AOT Compatibility](#19-aot-compatibility)
20. [Full API Reference](#20-full-api-reference)

---

## 1. Installation & Reference

### NuGet

```powershell
dotnet add package LiteExcel
```

### csproj Reference

```xml
<ItemGroup>
  <PackageReference Include="LiteExcel" Version="2.1.3" />
</ItemGroup>
```

### Namespace

All APIs are in the `LiteExcel` namespace:

```csharp
using LiteExcel;
```

### Target Frameworks

- **net48**: legacy WinForms projects can reference directly
- **net8.0**: new projects, supports AOT
- C# 12 syntax (`<LangVersion>latest</LangVersion>`)

---

## 2. Quick Start

### Minimal Write

```csharp
var sheet = new SheetData
{
    SheetName = "Employees",
    Headers = new() { "Name", "Age", "Birthday" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Zhang San"), Cell.FromNumber(25), Cell.FromDate(new DateTime(2000, 1, 1)) },
        new Cell[] { Cell.FromText("Li Si"), Cell.FromNumber(30), Cell.FromDate(new DateTime(1995, 5, 10)) },
    },
};

XlsxWriter.Write("output.xlsx", sheet);
```

### Minimal Read

```csharp
var sheet = XlsxReader.Read("output.xlsx", sheetIndex: 0);

Console.WriteLine($"Sheet: {sheet.SheetName}");
Console.WriteLine($"Headers: {string.Join(", ", sheet.Headers)}");
foreach (var row in sheet.Rows)
{
    foreach (var cell in row)
    {
        Console.Write($"{cell.Type} ");
    }
    Console.WriteLine();
}
```

### Reading Cell Values

```csharp
var sheet = XlsxReader.Read("output.xlsx", 0);

foreach (var row in sheet.Rows)
{
    string name = row[0].Type == CellType.Text ? row[0].Text! : "";
    int age = row[1].Type == CellType.Number ? (int)row[1].Number : 0;
    DateTime birthday = row[2].Type == CellType.Date ? row[2].Date : DateTime.MinValue;

    Console.WriteLine($"{name}, {age}, {birthday:yyyy-MM-dd}");
}
```

**Determine cell type**: use `cell.Type`, then read the corresponding field.

| `cell.Type` | Valid field |
|---|---|
| `CellType.Text` | `cell.Text` |
| `CellType.Number` | `cell.Number` |
| `CellType.Date` | `cell.Date` |
| `CellType.Boolean` | `cell.Boolean` |
| `CellType.Empty` | (empty cell) |

---

## 3. Cells & Data Types

### Cell Class

`Cell` represents a cell; the `Type` property determines which value field is valid.

```csharp
public sealed class Cell
{
    public CellType Type { get; set; }
    public string? Text { get; set; }          // valid when CellType.Text
    public double Number { get; set; }           // valid when CellType.Number
    public DateTime Date { get; set; }           // valid when CellType.Date
    public bool Boolean { get; set; }            // valid when CellType.Boolean
    public CellStyle? Style { get; set; }        // cell style (optional)
    public string? NumberFormat { get; set; }    // number format (for writing)
    public bool IsEmpty { get; }                 // whether empty
}
```

### Factory Methods

```csharp
// Text
var textCell = Cell.FromText("Hello");
var emptyTextCell = Cell.FromText("");  // → CellType.Empty
var nullTextCell = Cell.FromText(null); // → CellType.Empty

// Number (optional format)
var numCell = Cell.FromNumber(3.14);
var moneyCell = Cell.FromNumber(9999.50, "#,##0.00");

// Date (optional format, default "yyyy-MM-dd")
var dateCell = Cell.FromDate(new DateTime(2024, 6, 1));
var dateCell2 = Cell.FromDate(DateTime.Now, "yyyy/MM/dd HH:mm:ss");

// Boolean
var boolCell = Cell.FromBoolean(true);

// Empty
var emptyCell = Cell.Empty;
```

### Supported Cell Types

| Type | Description |
|---|---|
| `Text` | Text (shared strings or inline strings) |
| `Number` | Number (exact long up to 12 digits) |
| `Date` | Date (stored as number + format code, auto-converted) |
| `Boolean` | Boolean (`t="b"`) |
| `Empty` | Empty cell |

### Number Format Quick Reference

| Format | Effect |
|---|---|
| `"0"` | Integer |
| `"0.00"` | Two decimals |
| `"#,##0"` | Thousands separator, integer |
| `"#,##0.00"` | Thousands separator, two decimals |
| `"0.00%"` | Percent |
| `"yyyy-MM-dd"` | Date (default) |
| `"yyyy/MM/dd"` | Date |
| `"HH:mm:ss"` | Time |
| `"yyyy-MM-dd HH:mm:ss"` | Date time |

### Automatic Date Detection on Read

When reading, the library queries the numFmtId in `styles.xml` and automatically converts date-formatted numbers to `CellType.Date`. Built-in date format IDs (14-22, 27-36, 45-47, 50-58) are recognized as dates.

---

## 4. Reading

### List All Sheet Names

```csharp
var names = XlsxReader.GetSheetNames("file.xlsx");
// ["Employees", "Salaries", "Departments"]
```

### Read Single Sheet by Index

```csharp
// sheetIndex starts at 0, firstRowIsHeader defaults to true
var sheet = XlsxReader.Read("file.xlsx", sheetIndex: 0);

// Do not treat first row as header
var sheet = XlsxReader.Read("file.xlsx", sheetIndex: 0, firstRowIsHeader: false);
```

### Read Single Sheet by Name

```csharp
var sheet = XlsxReader.Read("file.xlsx", sheetName: "Employees");
```

### Read All Worksheets

```csharp
var allSheets = XlsxReader.ReadAll("file.xlsx");
foreach (var sheet in allSheets)
{
    Console.WriteLine($"{sheet.SheetName}: {sheet.Rows.Count} rows");
}
```

### Streaming Read Large Files (No Memory Residency)

```csharp
XlsxReader.StreamRows("bigfile.xlsx", "Sheet1", row =>
{
    // row is IReadOnlyList<Cell>, callback per row
    foreach (var cell in row)
    {
        // process cell
    }
});
```

> **Note**: `StreamRows` automatically skips the header row (first row). Only processes data rows.

### Read with Progress

```csharp
XlsxReader.ReadWithProgress("bigfile.xlsx", sheetIndex: 0, (current, total) =>
{
    Console.WriteLine($"Progress: {current}/{total} ({current * 100 / total}%)");
});
```

> `current` increments from 1 to `total` (data rows, excluding header).

### SheetData Structure of Read Result

```csharp
public sealed class SheetData
{
    public string SheetName { get; set; }            // sheet name
    public List<string> Headers { get; set; }         // headers (filled when firstRowIsHeader=true)
    public List<IReadOnlyList<Cell>> Rows { get; set; } // data rows
    public List<CellRange> MergedRanges { get; set; }   // merged cell ranges
    public AutoFilter? Filter { get; set; }              // auto filter
    public CellStyle? HeaderStyle { get; set; }          // header style
    public CellStyle? DefaultStyle { get; set; }         // default style
    public Dictionary<int, CellStyle>? RowStyles { get; set; }      // row styles
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }  // column styles
    public Dictionary<int, double>? RowHeights { get; set; }       // row heights
    public Dictionary<string, string>? Comments { get; set; }      // cell comments
    public List<DataValidation>? Validations { get; set; }          // data validation
}
```

---

## 5. Writing

### Write Single Sheet

```csharp
var sheet = new SheetData { ... };
XlsxWriter.Write("output.xlsx", sheet);
```

### Write Multiple Sheets

```csharp
var sheets = new List<SheetData>
{
    new() { SheetName = "Sheet1", Headers = new() { "A" }, Rows = ... },
    new() { SheetName = "Sheet2", Headers = new() { "B" }, Rows = ... },
};
XlsxWriter.Write("multi.xlsx", sheets);
```

### Freeze Header

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    FreezeHeader = true,  // freeze first row
};
```

### Column Widths

```csharp
var sheet = new SheetData
{
    Headers = new() { "Name", "Age", "Remark" },
    Rows = ...,
    ColumnWidths = new() { 15, 8, 30 },  // width in Excel character units
};
```

### Auto Column Widths

```csharp
var sheet = new SheetData { ... };

// Estimate best widths (CJK chars count as 2, others 1, range 8~50)
XlsxWriter.AutoColumnWidths(sheet);

XlsxWriter.Write("output.xlsx", sheet);
```

> Call `AutoColumnWidths` before `Write`; it fills `sheet.ColumnWidths`.

### Sheet Name Validation

Sheet names are validated on write; invalid names throw `InvalidSheetNameException`:
- Must not be empty
- Max 31 characters
- Must not contain `\ / ? * [ ] :`

```csharp
try
{
    XlsxWriter.Write("bad.xlsx", new SheetData { SheetName = "test[1]" });
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine(ex.Message);
}
```

---

## 6. Styling

### Style Priority (Override)

```
Cell.Style  >  RowStyles[row]  >  ColumnStyles[col]  >  DefaultStyle
```

> Override: if a cell has its own Style, it fully uses it, not inheriting row/column/sheet defaults.

### Cell Style

```csharp
var style = new CellStyle
{
    FontName = "Segoe UI",
    FontSize = 14,
    Bold = true,
    Italic = false,
    FontColor = "#FF0000",        // red font
    FillColor = "#FFFF00",        // yellow fill
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    WrapText = true,
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
        Bottom = new BorderEdge { Style = "thin", Color = "#000000" },
        Left = new BorderEdge { Style = "medium" },
        Right = new BorderEdge { Style = "medium" },
    },
};

var cell = new Cell { Type = CellType.Text, Text = "styled", Style = style };
```

### Header / Row / Column / Default Styles

```csharp
var sheet = new SheetData
{
    Headers = new() { "Name", "Score", "Grade" },
    Rows = ...,

    // header style
    HeaderStyle = new CellStyle { Bold = true, FillColor = "#4472C4", FontColor = "#FFFFFF" },

    // default style
    DefaultStyle = new CellStyle { FontSize = 11 },

    // column styles (key = 0-based column index)
    ColumnStyles = new()
    {
        { 2, new CellStyle { HorizontalAlignment = HorizontalAlignment.Center } },
    },

    // row styles (key = 0-based row index, corresponding to Rows)
    RowStyles = new()
    {
        { 1, new CellStyle { FillColor = "#FFFF00" } },
    },
};
```

### Border Styles

| Value | Description |
|---|---|
| `"thin"` | thin |
| `"medium"` | medium |
| `"thick"` | thick |
| `"dotted"` | dotted |
| `"dashed"` | dashed |
| `"double"` | double |
| `"none"` | none |

---

## 7. Merged Cells

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C" },
    Rows = ...,
    MergedRanges = new()
    {
        new CellRange(0, 0, 0, 2),  // merge first data row A1:C1
        new CellRange(1, 3, 0, 0),  // merge A2:A4 (3 rows)
    },
};
```

---

## 8. Auto Filter

```csharp
// Method 1: pass filter conditions, library computes hidden rows
var sheet = new SheetData
{
    Headers = new() { "Name", "City" },
    Rows = ...,
    Filter = new AutoFilter
    {
        Range = "A1:B5",
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

// Method 2: pass hidden row numbers directly
sheet.Filter = new AutoFilter
{
    Range = "A1:B5",
    HiddenRows = new() { 1, 3 },
};
```

### Filter Condition Types

| Type | Description |
|---|---|
| `Equals` | equals one of the values (multi-select) |
| `Compare` | compare (`>` `<` `>=` `<=` `between`) |
| `Contains` | text contains |
| `BeginsWith` | begins with |
| `EndsWith` | ends with |
| `Blank` | blank / not blank |

### Between Example

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.Between,
    MinValue = "60",
    MaxValue = "90",
},
```

---

## 9. Row Height & Column Width

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    RowHeights = new()
    {
        { 0, 30.0 },   // first data row height 30pt
    },
};

// auto column widths
XlsxWriter.AutoColumnWidths(sheet);
XlsxWriter.Write("output.xlsx", sheet);
```

---

## 10. Cell Comments

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    Comments = new()
    {
        { "A1", "This is a comment on A1" },
        { "B1", "Comment with special chars < & >" },
    },
};
XlsxWriter.Write("comments.xlsx", sheet);

var read = XlsxReader.Read("comments.xlsx", 0);
Console.WriteLine(read.Comments?["A1"]);
```

---

## 11. Data Validation (Dropdown List)

```csharp
var sheet = new SheetData
{
    Headers = new() { "Name", "Department", "Score" },
    Rows = ...,
    Validations = new()
    {
        // dropdown list
        new DataValidation
        {
            Type = DataValidationType.List,
            Sqref = "B2:B100",
            Formula1 = "\"IT,HR,Finance,Sales\"",
            AllowBlank = true,
            PromptTitle = "Department",
            Prompt = "Select a department from the list",
        },
        // integer range 0-100
        new DataValidation
        {
            Type = DataValidationType.WholeNumber,
            Sqref = "C2:C100",
            Formula1 = "0",
            Formula2 = "100",
        },
    },
};
XlsxWriter.Write("validation.xlsx", sheet);
```

### DataValidationType

| Type | Description | Formula1 | Formula2 |
|---|---|---|---|
| `List` | dropdown list | `"\"A,B,C\""` | not used |
| `WholeNumber` | integer | lower bound | upper bound |
| `Decimal` | decimal | lower bound | upper bound |
| `Date` | date | start date | end date |

---

## 12. Appending Data

```csharp
// Append 2 rows to existing sheet with same name; if name doesn't exist, add new sheet
XlsxWriter.Append("data.xlsx", new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID" },
    Rows = new() { new Cell[] { Cell.FromNumber(4) }, new Cell[] { Cell.FromNumber(5) } },
});
```

> When appending to an existing file, LiteExcel preserves existing document properties (author, title, subject, created time, and so on) and automatically updates the modified time.

> `Append` reconstructs the workbook from the LiteExcel model. Supported data such as styles, merged cells, filters, comments, and data validations is retained. OOXML parts not mapped by LiteExcel, including Excel Tables, themes, pivot tables, and charts, are not guaranteed to be retained.

Pass a third argument to update selected properties:

```csharp
XlsxWriter.Append("data.xlsx", moreRows, new WorkbookProperties
{
    LastModifiedBy = "Editor",
    Title = "Updated report",
});
```

---

## 13. List&lt;T&gt; Mapping (reflection, not AOT compatible)

> For AOT projects, use the SheetData or DataTable overloads.

```csharp
public class Person
{
    [LiteColumn(Name = "Name", Order = 0)]
    public string Name { get; set; }
    [LiteColumn(Name = "Age", Order = 1)]
    public int Age { get; set; }
    [LiteColumn(Ignore = true)]
    public string? InternalId { get; set; }
}

// Write
var list = new List<Person> { new() { Name = "Zhang San", Age = 25 } };
XlsxWriter.Write("people.xlsx", list);

// Read
var read = XlsxReader.Read<Person>("people.xlsx");

// Fluent configuration
XlsxWriter.Write("people.xlsx", list, opt => opt
    .Column(x => x.Name, "Name")
    .Column(x => x.Age, "Age")
    .Ignore(x => x.InternalId));
```

### Three Ways to Customize Columns

1. **`[LiteColumn]` attribute**: `Name` / `Order` / `Format` / `Ignore`
2. **Fluent callback**: `opt.Column(x => x.Name, "Name").Ignore(x => x.Id)`
3. **Dictionary mapping**: `opt.Map(new Dictionary<string,string> { { "Name", "Name" } })`

---

## 14. DataTable Convenience API (AOT safe)

```csharp
var dt = new DataTable();
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Rows.Add("Zhang San", 25);
XlsxWriter.Write("data.xlsx", dt);

var read = XlsxReader.ReadAsDataTable("data.xlsx");
```

---

## 15. Stream Read/Write

All read/write APIs have Stream overloads:

```csharp
using var ms = new MemoryStream();
XlsxWriter.Write(ms, sheet);           // write to stream
ms.Position = 0;
var read = XlsxReader.Read(ms, 0);     // read from stream

using var fs = File.OpenRead("out.xlsx");
var all = XlsxReader.ReadAll(fs);
```

> **Note**: Stream overloads do not close the passed stream (`leaveOpen: true`); the caller owns the stream lifecycle.

---

## 16. Streaming Read & Progress Callback

```csharp
// Stream rows, no memory residency
XlsxReader.StreamRows("big.xlsx", "Sheet1", row =>
{
    foreach (var cell in row) Process(cell);
});

// With progress callback
XlsxReader.ReadWithProgress("big.xlsx", 0, (current, total) =>
{
    Console.WriteLine($"Progress: {current}/{total}");
});
```

---

## 17. Document Properties (Author/Time/Title)

### Write with Properties

```csharp
var props = new WorkbookProperties
{
    Creator = "Zhang San",                                    // author
    LastModifiedBy = "Li Si",                                 // last modified by
    Created = DateTime.Now,                                   // created
    Modified = DateTime.Now,                                  // modified
    Title = "Monthly Report",                                 // title
    Subject = "Finance",                                      // subject
    Application = "MyApp",                                    // application name (optional)
};

XlsxWriter.Write("report.xlsx", sheet, props);
```

> When `Application` is null, it defaults to the host assembly name (`Assembly.GetEntryAssembly().GetName().Name`).

### Read Properties

```csharp
var props = XlsxReader.ReadProperties("report.xlsx");
Console.WriteLine($"Creator: {props.Creator}");
Console.WriteLine($"LastModifiedBy: {props.LastModifiedBy}");
Console.WriteLine($"Created: {props.Created}");
Console.WriteLine($"Modified: {props.Modified}");
Console.WriteLine($"Title: {props.Title}");
Console.WriteLine($"Subject: {props.Subject}");
Console.WriteLine($"Application: {props.Application}");
```

> If the file has no docProps, no exception is thrown; an empty object is returned (all fields null).

### Write without Properties (backward compatible)

```csharp
// No props passed, no docProps generated (same behavior as older versions)
XlsxWriter.Write("output.xlsx", sheet);
```

### WorkbookProperties Fields

| Field | Type | XML |
|---|---|---|
| `Creator` | `string?` | dc:creator |
| `LastModifiedBy` | `string?` | cp:lastModifiedBy |
| `Created` | `DateTime?` | dcterms:created |
| `Modified` | `DateTime?` | dcterms:modified |
| `Title` | `string?` | dc:title |
| `Subject` | `string?` | dc:subject |
| `Application` | `string?` | app.xml Application |

---

## 18. Error Handling

### LiteExcelException

All user-facing errors throw `LiteExcelException`:

```csharp
try
{
    var sheet = XlsxReader.Read("not-an-xlsx.txt", 0);
}
catch (LiteExcelException ex)
{
    Console.WriteLine($"Read failed: {ex.Message}");
    // "This is not a valid xlsx file"
}
```

### InvalidSheetNameException

Invalid sheet name on write:

```csharp
try
{
    XlsxWriter.Write("bad.xlsx", new SheetData { SheetName = "sheet[1]" });
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine(ex.Message);
}
```

### Common Errors

> Exception messages are currently emitted in Chinese. The messages below reflect the current wording.

| Scenario | Exception | Message |
|---|---|---|
| Not an xlsx file | `LiteExcelException` | "这不是有效的 xlsx 文件" |
| Sheet name not found | `LiteExcelException` | "找不到工作表：{name}（共有 {n} 张表）" |
| Invalid sheet name | `InvalidSheetNameException` | details included |
| Sheet index out of range | `ArgumentOutOfRangeException` | "工作表索引超出范围" |
| Empty sheet list | `ArgumentException` | "至少需要一张工作表" |

---

## 19. AOT Compatibility

### AOT-safe APIs (no reflection)

| API | Description |
|---|---|
| `Read(path/stream, ...)` | returns `SheetData` |
| `Write(path/stream, SheetData)` | accepts `SheetData` |
| `ReadAsDataTable(...)` | DataTable has its own schema |
| `Write(path, DataTable)` | DataTable write |
| `StreamRows(...)` | streaming read |
| `GetSheetNames(...)` | list sheet names |
| `Append(...)` | append |
| `AutoColumnWidths(...)` | auto column widths |

### AOT-unsafe APIs (reflection, marked `[RequiresUnreferencedCode]`)

| API | Description |
|---|---|
| `Read<T>(...)` | List<T> read |
| `Write<T>(...)` | List<T> write |

> AOT projects will get `IL3050`/`IL2026` warnings when calling these. Non-AOT projects (net48, net8 normal publish) are unaffected.

---

## 20. Full API Reference

### XlsxReader

| Method | Returns | Description |
|---|---|---|
| `GetSheetNames(string path)` | `List<string>` | list all sheet names |
| `GetSheetNames(Stream stream)` | `List<string>` | list from stream |
| `Read(string path, int sheetIndex, bool firstRowIsHeader = true)` | `SheetData` | read by index |
| `Read(string path, string sheetName, bool firstRowIsHeader = true)` | `SheetData` | read by name |
| `Read(Stream stream, int sheetIndex, bool firstRowIsHeader = true)` | `SheetData` | read by index from stream |
| `Read(Stream stream, string sheetName, bool firstRowIsHeader = true)` | `SheetData` | read by name from stream |
| `ReadAll(string path)` | `List<SheetData>` | read all |
| `ReadAll(Stream stream)` | `List<SheetData>` | read all from stream |
| `StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)` | `void` | stream rows |
| `StreamRows(Stream stream, string sheetName, Action<IReadOnlyList<Cell>> onRow)` | `void` | stream rows from stream |
| `ReadWithProgress(string path, int sheetIndex, Action<int,int> onProgress)` | `void` | read with progress |
| `ReadAsDataTable(string path, int sheetIndex = 0, bool firstRowIsHeader = true)` | `DataTable` | read as DataTable |
| `ReadAsDataTable(string path, string sheetName, bool firstRowIsHeader = true)` | `DataTable` | read by name as DataTable |
| `ReadAsDataTable(Stream stream, int sheetIndex = 0, bool firstRowIsHeader = true)` | `DataTable` | read as DataTable from stream |
| `ReadAsDataTable(Stream stream, string sheetName, bool firstRowIsHeader = true)` | `DataTable` | read by name from stream |
| `Read<T>(string path, int sheetIndex = 0, Action<ReadOptions<T>>? configure = null)` ⚠️ | `List<T>` | read as List<T> (reflection) |
| `Read<T>(string path, string sheetName, Action<ReadOptions<T>>? configure = null)` ⚠️ | `List<T>` | read by name as List<T> |

> ⚠️ Marked `[RequiresUnreferencedCode]`, not AOT compatible.

### XlsxWriter

| Method | Description |
|---|---|
| `Write(string path, SheetData data)` | write single sheet |
| `Write(string path, IReadOnlyList<SheetData> sheets)` | write multiple sheets |
| `Write(Stream stream, SheetData data)` | write single sheet to stream |
| `Write(Stream stream, IReadOnlyList<SheetData> sheets)` | write multiple sheets to stream |
| `Write(string path, DataTable table, string sheetName = "Sheet1")` | write from DataTable |
| `Write<T>(string path, IEnumerable<T> data, Action<WriteOptions<T>>? configure = null)` ⚠️ | write from List<T> (reflection) |
| `Append(string path, SheetData? newData, WorkbookProperties? updateProperties = null)` | append data and optionally update document properties |
| `AutoColumnWidths(SheetData sheet)` | auto estimate column widths |

### CellRef (Utility)

| Method | Description |
|---|---|
| `CellRef.Parse(string cellRef)` | "A1" → (row=0, col=0) |
| `CellRef.ToString(int row, int col)` | (0, 0) → "A1" |
| `CellRef.ColToLetter(int col)` | 0 → "A", 26 → "AA" |
| `CellRef.LetterToCol(string letters)` | "A" → 0, "AA" → 26 |

### Model Classes

| Class | Description |
|---|---|
| `SheetData` | worksheet data |
| `Cell` | cell |
| `CellStyle` / `BorderStyle` / `BorderEdge` | styles |
| `CellRange` | range (merged cells) |
| `AutoFilter` / `FilterColumn` | auto filter |
| `DataValidation` | data validation |
| `WorkbookProperties` | document properties |
| `LiteExcelException` / `InvalidSheetNameException` | exceptions |
| `LiteColumnAttribute` | List<T> column attribute |
| `WriteOptions<T>` / `ReadOptions<T>` | List<T> configuration |

### Enums

| Enum | Values |
|---|---|
| `CellType` | Text, Number, Date, Boolean, Empty |
| `HorizontalAlignment` | General, Left, Center, Right |
| `VerticalAlignment` | Top, Center, Bottom |
| `FilterType` | Equals, Compare, Contains, BeginsWith, EndsWith, Blank |
| `FilterOperator` | GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between |
| `DataValidationType` | List, WholeNumber, Decimal, Date |
