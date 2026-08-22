# LiteExcel Usage Guide

**Version**: 2.4.2  
**Target Frameworks**: net48 + net8.0  
**Dependencies**: Zero third-party dependencies, .NET BCL only

---

## Table of Contents

1. [Installation & Reference](#1-installation--reference)
2. [Object-Model API (Recommended)](#2-object-model-api-recommended)
3. [Quick Start](#3-quick-start)
4. [Cells & Data Types](#4-cells--data-types)
5. [Reading](#5-reading)
6. [Writing](#6-writing)
7. [Styling](#7-styling)
8. [Merged Cells](#8-merged-cells)
9. [Auto Filter](#9-auto-filter)
10. [Row Height & Column Width](#10-row-height--column-width)
11. [Cell Comments](#11-cell-comments)
12. [Hyperlinks](#12-hyperlinks)
13. [Freeze Panes](#13-freeze-panes)
14. [Images](#14-images)
15. [Data Validation (Dropdown List)](#15-data-validation-dropdown-list)
16. [Appending Data](#16-appending-data)
17. [List<T> Mapping (reflection, not AOT compatible)](#17-listt-mapping-reflection-not-aot-compatible)
18. [DataTable Convenience API (AOT safe)](#18-datatable-convenience-api-aot-safe)
19. [Stream Read/Write](#19-stream-readwrite)
20. [Streaming Read & Progress Callback](#20-streaming-read--progress-callback)
21. [Document Properties (Author/Time/Title)](#21-document-properties-authortimetitle)
22. [File-Level Security (Open Password / Modify Password)](#22-file-level-security-open-password--modify-password)
23. [Error Handling](#23-error-handling)
24. [AOT Compatibility](#24-aot-compatibility)
25. [Full API Reference](#25-full-api-reference)

---

## 1. Installation & Reference

### NuGet

```powershell
dotnet add package LiteExcel
```

### csproj Reference

```xml
<ItemGroup>
  <PackageReference Include="LiteExcel" Version="2.4.2" />
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

## 2. Object-Model API (Recommended)

Since `2.2.0`, an intuitive object-model API is provided with a natural hierarchy:

```text
Excel             unified facade (open / create / convenience IO / streaming)
  -> Workbook     workbook (worksheet collection / document properties / save)
      -> Worksheet sheet (cells / ranges / merge / styles)
          -> Cells / Cell / ExcelRange
```

The object-model API is built on the same read/write engine but wraps coordinates, typed access, and save semantics into familiar Excel-style usage. `XlsxReader / XlsxWriter / SheetData / Cell` remain fully supported; old and new APIs can be mixed, and files written by either are readable by the other.

### 2.1 Create a Workbook

```csharp
using LiteExcel;

// Default xlsx, contains one worksheet named "Sheet1"
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
```

Specify the format and initial sheet name:

```csharp
var wbCsv = Excel.Create(ExcelFormat.Csv);              // create a csv workbook
var wb2   = Excel.Create("Employees", ExcelFormat.Xlsx); // create and name the first sheet
```

Supported formats: `Xlsx`, `Xlsm`, `Csv`, `Xls`, `Xlsb` (the latter two are legacy formats with read and write implemented).

### 2.2 Open an Existing File

```csharp
// Auto-detect the format from the extension (.xlsx / .xlsm / .csv)
var opened = Excel.Open("data.xlsx");

// Or force a specific format
var forced = Excel.Open("data.csv", ExcelFormat.Csv);
```

`Excel.DetectFormat(path)` returns the detected format.

Opening an encrypted file requires the password:

```csharp
var opened = Excel.Open("encrypted.xlsx", new ExcelReadOptions
{
    OpenPassword = "your-password",            // open password
    ModifyPassword = "your-modify-password",   // modify password (optional)
});
```

> Full details of file-level security (open/modify password) can be found in [§22 File-Level Security](#22-file-level-security-open-password--modify-password).

### 2.3 Read and Write Cells

Coordinates are **1-based**, so `Cell(1, 1)` is `A1`; A1 addresses are also accepted.

```csharp
var ws = wb.Worksheets["Sheet1"];

// Write
ws.SetValue("A1", "Name");
ws.SetValue("B1", "Age");
ws.SetValue(2, 1, "Zhang San");  // A2
ws.SetValue(2, 2, 25);           // B2

// Read
string name = ws.Cell("A2").GetString();   // "Zhang San"
double age = ws.Cell(2, 2).GetDouble();    // 25
```

Accessor methods:

| Method | Description |
|---|---|
| `GetString()` / `TryGetString(out var v)` | get text |
| `GetDouble()` / `TryGetDouble(out var v)` | get number |
| `GetDateTime()` / `TryGetDateTime(out var v)` | get date |
| `GetBoolean()` / `TryGetBoolean(out var v)` | get boolean |
| `GetValue()` | get as `object?` by type |
| `SetValue(object? value)` | write (string/number/date/boolean/formula) |

`Cell.Style`, `Cell.NumberFormat` read/write cell style and number format; `Cell.IsFormula` tells whether the cell holds a formula.

#### Combined scenario: change value + style + comment + save

After opening an existing file you can read, restyle and comment a **specific cell** in one flow, then save:

```csharp
var wb = Excel.Open("report.xlsx");
var ws = wb.Worksheets["Sheet1"];

// Change the value of A2
ws.Cell("A2").SetValue("Done");

// Change A2's background color and font
ws.Cell("A2").Style = new CellStyle
{
    FillColor = "#FFFF00",              // background
    FontName  = "Microsoft YaHei",      // font name
    FontSize  = 14,
    Bold      = true,
    FontColor = "#FF0000",              // font color
};

// Add a comment to A2
ws.Comments ??= new();
ws.Comments["A2"] = "Needs manual review";

// Batch style for a range (applied to every cell in A2:C3)
ws.Range("A2:C3").Style = new CellStyle { FillColor = "#D9E1F2", Italic = true };

wb.Save();                              // overwrite the original file
```

> See [§7 Styling](#7-styling) and [§11 Cell Comments](#11-cell-comments) for details.

### 2.4 Collection Access via `Cells`

```csharp
var cells = ws.Cells;

var a1 = cells[1, 1];            // 1-based coordinate
var b2 = cells["B2"];            // A1 address
var r  = cells.Range("A1:C10");  // range

foreach (var cell in cells)      // enumerate all stored cells
{
    Console.WriteLine($"{cell.Text} / {cell.Number}");
}

cells.SetValue("D2", "note");
cells.Clear();                   // clear everything
```

### 2.5 Range Operations via `ExcelRange`

`Worksheet.Range(...)` returns an `ExcelRange` (note: the class is `ExcelRange`, not `Range`, to avoid clashing with BCL `System.Range`):

```csharp
var range = ws.Range("A1:C3");            // or ws.Range(1, 1, 3, 3)
range.Fill(0);                            // fill the whole range
range.Fill(new object?[,] { { 1, 2, 3 }, { 4, 5, 6 } });  // fill from a matrix
var values = range.ToValues();            // object?[,]
var cells = range.ToCells();              // Cell[,]
range.Clear();
range.Merge();                            // merge the range
range.Unmerge();
range.Style = new CellStyle { Bold = true };  // range style

foreach (var cell in range) { /* enumerate cells in the range */ }
```

### 2.6 Save and Save As

```csharp
wb.Save();                        // save to the current path (throws LiteExcelException if none)
wb.SaveAs("output.xlsx");         // save as, updates the current path
wb.SaveAs("output.csv", ExcelFormat.Csv);  // cross-format save (depends on backend)
wb.Save(stream, ExcelFormat.Xlsx);         // write to a stream
```

Rules:

- After a successful `SaveAs`, subsequent `Save()` writes to the new path.
- Cross-format save succeeds only if the target backend supports it; `csv` holds tabular data only.
- `Excel.Write(path, workbook)` is equivalent to "save as the given path".

### 2.7 Worksheet Management

```csharp
var wb = Excel.Create();
wb.Worksheets.Add("Sheet2");              // add
wb.Worksheets.Add("Sheet3");
wb.Worksheets.Move(0, 1);                 // move
wb.Worksheets.Remove("Sheet2");           // remove by name
wb.Worksheets.RemoveAt(0);
bool has = wb.Worksheets.Contains("Sheet3");
var names = wb.Worksheets.Names;          // ["Sheet3", ...]
```

### 2.8 Document Properties

```csharp
var props = wb.Properties;
props.Creator = "LiteExcel";
props.Title = "Sample Report";
wb.Save();
```

### 2.9 List\<T\> / DataTable Convenience API

```csharp
// List<T> mapping (reflection, not AOT compatible)
Excel.Write("out.xlsx", new[] { new Person { Name = "Zhang San", Age = 25 } });
var list = Excel.Read<Person>("out.xlsx");

// DataTable (AOT safe)
var dt = Excel.ReadAsDataTable("out.xlsx");
Excel.Write("out2.xlsx", dt);
```

`Excel.GetSheetNames(path)` lists all worksheet names.

### 2.10 Streaming Large Files

```csharp
// Streaming write: row by row, no memory residency
using (var writer = Excel.CreateWriter("large.xlsx"))
{
    writer.WriteRow(new object?[] { "Name", "Age" });
    for (int i = 0; i < 100000; i++)
        writer.WriteRow(new object?[] { $"User{i}", i });
}

// Streaming read
Excel.StreamRows("large.xlsx", "Sheet1", row =>
{
    Console.WriteLine(row[0]?.Text);
});
```

### 2.11 Formulas and Advanced Capabilities

```csharp
ws.SetValue("A1", 1);
ws.SetValue("A2", 2);
ws.Cell("A3").SetValue(Cell.FromFormula("SUM(A1:A2)"));  // write a formula string
bool isFormula = ws.Cell("A3").IsFormula;

ws.Merge("A1:B1");                    // merge
ws.FreezeHeader = true;               // freeze header (equivalent to FreezeRows = 1)
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
```

Styles, merge, comments, data validation, auto filter, row height and column width, hyperlinks, freeze panes and images are all available at the `Worksheet` level, mirroring the `SheetData` capabilities.

### 2.12 Format Support Matrix

| Format | Read | Write | Notes |
|---|---|---|---|
| `xlsx` | ✅ | ✅ | full read/write; 1904 date system read/write; saving a macro workbook to `.xlsx` throws (see "Macros & degradation") |
| `xlsm` | ✅ | ✅ | read/write/save; `vbaProject.bin` macro part and host codeName bindings (`workbookPr`/`sheetPr`) preserved on save |
| `csv` | ✅ | ✅ | tabular data only, no styles/merge |
| `xls` | ✅ | ✅ | read/write (BIFF8, Excel 97+); formulas written as static cached values; saving a macro workbook to `.xls` throws (see "Macros & degradation") |
| `xlsb` | ✅ | ✅ | read/write (BIFF12 binary OOXML); formulas written as static cached values; 1904 date system read/write |

> **xls read scope**: `Excel.Open("file.xls")` reads BIFF8 workbooks: data cells (text/number/date/boolean), shared strings (including cross-CONTINUE continuation), merged cells, column widths, row heights, frozen header. Formula cells return the cached result value and the parsed formula text (common cell references, operators and built-in functions; unsupported formulas such as array/3D references return only the cached value).

> **xls write scope**: `wb.SaveAs("file.xls", ExcelFormat.Xls)` writes BIFF8 workbooks: multiple sheets (Chinese names), text/number/date/boolean cells, merged cells, column widths, row heights, frozen header, custom number formats. Formula cells are written as static cached values (formula text is not preserved). Verified by opening in Excel.

> **xlsb read scope**: `Excel.Open("file.xlsb")` reads the binary OOXML variant: data cells (text/number/date/boolean/error), shared strings, merged cells, column widths, row heights, frozen header, 1904 date system. Formula cells return the cached result value and the parsed formula text.
>
> **xlsb write scope**: `wb.SaveAs("file.xlsb", ExcelFormat.Xlsb)` writes BIFF12 workbooks: multiple sheets (Chinese names), text/number/date/boolean cells, shared strings, number formats, merged cells, column widths, row heights, frozen header. Formula cells are written as static cached values (formula text is not preserved). Verified by opening in Excel (no repair prompt, values consistent after Excel re-save) and cross-checked with SheetJS.

> **Save fidelity**: after `Excel.Open` + modify + save, LiteExcel rebuilds the mapped parts (sheet data, styles, merges, comments, validations, filters, formulas, etc.) and **preserves unmapped OOXML parts as raw bytes** (macro `vbaProject.bin`, themes, drawings, charts, tables, external links, etc.). So an `xlsm` opened → modified → saved keeps its macros. In addition, the VBA host code names (`workbookPr@codeName` / `sheetPr@codeName`) are captured on open and written back in their schema-required positions on save, so module bindings in the VBA project (`ThisWorkbook`, sheet modules, event macros) are not broken by hosts being renamed.
>
> **Degradation rule**: if the workbook structure changed after open (sheets added/removed/renamed/moved), sheet-level unmapped relationships (drawings, hyperlinks) are not re-attached to the new file, though the raw part bytes are still kept as harmless unreferenced entries. Workbook-level parts (macros, theme) are unaffected by structure changes.

> **1904 date system**: when `Excel.Open` reads a 1904-date-system workbook (`workbookPr@date1904` / `BrtWbProp` flags / `DATE1904` record), date cells are converted on the 1904 base (1904-01-01 = serial 0). `SaveAs` to xlsx/xlsb/xls writes back the 1904 flag and keeps serials consistent, so 1904 workbooks do not shift 4 years across format conversion.

> **Encrypted file detection**: password-protected xlsx/xlsm/xlsb are actually OLE CFB containers (containing `EncryptionInfo`/`EncryptedPackage` streams). Since 2.4.0, open-password reading and password save are supported (see [§22 File-Level Security](#22-file-level-security-open-password--modify-password)); encrypted `.xls` (BIFF8 `FILEPASS` record) is detected with a clear error.

> **Macros & degradation**: if a workbook contains VBA macros (captured from `vbaProject.bin` on open), `SaveAs` to `.xlsx` or `.xls` (which do not support macros) throws `LiteExcelException` to prevent silent loss, before the file is created. Save macro workbooks as `.xlsm` or `.xlsb`. Cell data values are never silently lost across any format conversion; metadata unsupported by xls (comments, data validation, auto filter) follows the documented degradation (ignored, no corrupt file produced).

> **Stream API**: `Excel.Open(Stream, format)` supports reading all five formats via the object model (format must be specified explicitly — a stream has no extension); `Workbook.Save(Stream, format)` supports saving all five formats. The input stream is not closed (caller manages its lifetime); non-seekable streams are supported (copied to memory internally). `XlsxReader.StreamRows(Stream, ...)` remains an xlsx/xlsm-only low-level streaming row reader. After opening from a stream, `Workbook.CurrentPath` is null — use `SaveAs` to specify a save path.

### 2.13 Old vs New API

| Scenario | Object-Model API | XlsxWriter / XlsxReader |
|---|---|---|
| Open a file | `Excel.Open(path)` | `XlsxReader.Read(path, 0)` |
| Create / write | `Excel.Create()` + `SaveAs` | `XlsxWriter.Write(path, sheet)` |
| Read one sheet | `Workbook.Worksheets[i]` | `XlsxReader.Read(path, i)` |
| Read as List\<T\> | `Excel.Read<T>(path)` | `XlsxReader.Read<T>(path)` |
| Read as DataTable | `Excel.ReadAsDataTable(path)` | `XlsxReader.ReadAsDataTable(path)` |
| Streaming read | `Excel.StreamRows(path, name, cb)` | `XlsxReader.StreamRows(path, name, cb)` |
| Streaming write | `Excel.CreateWriter(path)` | `XlsxStreamWriter` |

---

## 3. Quick Start

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

## 4. Cells & Data Types

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

## 5. Reading

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

## 6. Writing

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

### Freeze Rows and Columns

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C" },
    Rows = ...,
    FreezeRows = 2,       // freeze first 2 rows
    FreezeColumns = 1,    // freeze first column
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

### CSV Write Separator (2.4.3+)

CSV output uses comma by default. To switch to semicolon (common in Chinese-locale Excel) or Tab:

```csharp
Excel.Write("out.csv", wb, new ExcelWriteOptions { Separator = ';' });
```

The setting takes effect for that write only.

> CSV supports only configurable separators. Other formats ignore this.
> Other Excel-only capabilities (comments, merges, filters, images, styles, etc.) are reported via the degradation callback. See [§25 Capability Degradation Callback](#25-capability-degradation-callback-ondegradation-242).

---

## 7. Styling

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

### Restyle a specific cell / range (object-model API)

After opening an existing file, assign directly to `Cell.Style` / `ExcelRange.Style` to restyle a **specific cell** or **range**. `Style` is a wholesale replacement; fields you leave unset keep Excel defaults:

```csharp
var wb = Excel.Open("styled.xlsx");
var ws = wb.Worksheets["Sheet1"];

// Single cell: background + font + font color
ws.Cell("A2").Style = new CellStyle
{
    FillColor = "#FFFF00",          // background (#RRGGBB)
    FontName  = "Microsoft YaHei",  // font name
    FontSize  = 14,                 // size in points
    Bold      = true,
    FontColor = "#FF0000",          // font color
};

// Range: every cell inside A2:C3 gets the same style
ws.Range("A2:C3").Style = new CellStyle
{
    FillColor = "#D9E1F2",
    Italic    = true,
    HorizontalAlignment = HorizontalAlignment.Center,
};

// Change value and style at once
ws.Cell("B2").SetValue("new value");
ws.Cell("B2").Style = new CellStyle { FillColor = "#92D050", Bold = true };

wb.Save();   // or wb.SaveAs("styled2.xlsx")
```

> `ExcelRange.Style` walks every cell in the range. `Style` is a **replacement**, not incremental merge — if you need to keep existing style and change one property, read the current style first and copy it (`CellStyle` provides `Clone()`).

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

## 8. Merged Cells

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

## 9. Auto Filter

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
};
```

---

## 10. Row Height & Column Width

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

## 11. Cell Comments

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

### Object-model API: add / read comments on a specific cell

After opening an existing file, add, update, read or remove comments per cell:

```csharp
var wb = Excel.Open("comments.xlsx");
var ws = wb.Worksheets["Sheet1"];

// Add a comment (initialize Comments if it is null)
ws.Comments ??= new();
ws.Comments["A2"] = "This is a comment on A2";

// Update an existing comment
ws.Comments["A2"] = "Comment updated";

// Read a comment
if (ws.Comments is not null && ws.Comments.TryGetValue("A2", out var text))
    Console.WriteLine($"A2 comment: {text}");

// Remove a comment
ws.Comments.Remove("A2");

wb.Save();
```

---

## 12. Hyperlinks

### Writing a Hyperlink

Hyperlinks are supported in xlsx/xlsm/xlsb/xls. Set them via the `Cell.Hyperlink` property:

```csharp
var sheet = new SheetData
{
    Headers = new() { "Name", "Homepage" },
    Rows = new()
    {
        new Cell[]
        {
            Cell.FromText("Zhang San"),
            new Cell { Type = CellType.Text, Text = "Click to visit", Hyperlink = new Hyperlink
            {
                Target = "https://example.com",
                Tooltip = "Zhang San's homepage",
            }},
        },
    },
};

XlsxWriter.Write("links.xlsx", sheet);
```

### Hyperlink Properties

| Property | Type | Description |
|---|---|---|
| `Target` | `string` | link target URL (required) |
| `Tooltip` | `string?` | hover tooltip text (optional) |
| `IsInternal` | `bool` | whether it is an internal link (e.g. `Sheet1!A1`) |

### Object-Model API: Set a Hyperlink

```csharp
var wb = Excel.Open("links.xlsx");
var ws = wb.Worksheets["Sheet1"];

// Set a hyperlink
ws.Cell("B2").Hyperlink = new Hyperlink
{
    Target = "https://example.com",
    Tooltip = "Click to visit",
};

wb.Save();
```

### Reading Hyperlinks

```csharp
var sheet = XlsxReader.Read("links.xlsx", 0);
var cell = sheet.Rows[0][1];  // B2
if (cell.Hyperlink is not null)
{
    Console.WriteLine($"Target: {cell.Hyperlink.Target}");
    Console.WriteLine($"Tooltip: {cell.Hyperlink.Tooltip}");
}
```

> Hyperlinks are read/write supported in xlsx/xlsm/xlsb/xls (external URL/file/mailto/UNC and internal `#Sheet!A1` jumps); csv does not support hyperlinks. When `IsInternal=true`, `Target` looks like `#Sheet1!A1`.

---

## 13. Freeze Panes

### Freeze Rows / Columns

Control via the `SheetData.FreezeRows` and `FreezeColumns` properties:

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C", "D" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("data1"), Cell.FromText("x"), Cell.FromText("y"), Cell.FromText("z") },
        new Cell[] { Cell.FromText("data2"), Cell.FromText("a"), Cell.FromText("b"), Cell.FromText("c") },
    },
    FreezeRows = 2,       // freeze first 2 rows
    FreezeColumns = 1,    // freeze first column
};
```

### FreezeHeader Compatibility

`FreezeHeader = true` is equivalent to `FreezeRows = 1`:

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    FreezeHeader = true,   // equivalent to FreezeRows = 1
};
```

### Object-Model API

```csharp
var wb = Excel.Open("report.xlsx");
var ws = wb.Worksheets["Sheet1"];

// Freeze the first 2 rows
ws.FreezeRows = 2;

// Freeze the first column
ws.FreezeColumns = 1;

// Or use the FreezeHeader compatibility syntax
ws.FreezeHeader = true;   // equivalent to FreezeRows = 1

wb.Save();
```

### Reading the Freeze State

```csharp
var sheet = XlsxReader.Read("frozen.xlsx", 0);
Console.WriteLine($"Frozen rows: {sheet.FreezeRows}");           // 0 = not frozen
Console.WriteLine($"Frozen columns: {sheet.FreezeColumns}");     // 0 = not frozen
Console.WriteLine($"Freeze header: {sheet.FreezeHeader}");       // true when FreezeRows > 0
```

---

## 14. Images

### Floating Image

Images are only supported in xlsx/xlsm format. Add a floating image via `Worksheet.AddImage`:

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];

// Read the image data
byte[] imageData = File.ReadAllBytes("logo.png");

// Add a floating image anchored at cell A1
ws.AddImage(imageData, row: 1, column: 1, widthPx: 200, heightPx: 100);

wb.SaveAs("image.xlsx");
```

### InCell Image

```csharp
// InCell mode (the image resizes with the cell)
ws.AddImage(imageData, row: 1, column: 1,
    placement: ImagePlacement.InCell);
```

### ImagePlacement Enum

| Value | Description |
|---|---|
| `Floating` | floating image; position and size can be specified (default) |
| `InCell` | InCell image; resizes with the cell |

### Coordinates and Size

- Coordinates (row, column) are **1-based**
- Width/height are in pixels; when omitted, the image's original size is used
- Image extension (png/jpg/gif/bmp) and pixel size are auto-detected

```csharp
// Omit width/height → use the original image size
ws.AddImage(imageData, row: 2, column: 3, placement: ImagePlacement.Floating);

// Specify extension and name
ws.AddImage(imageData, row: 1, column: 1,
    widthPx: 300, heightPx: 200,
    extension: "png", name: "Product image");
```

### Multiple Sheets

```csharp
var ws1 = wb.Worksheets["Sheet1"];
var ws2 = wb.Worksheets.Add("Sheet2");

ws1.AddImage(logoData, row: 1, column: 1, widthPx: 100, heightPx: 50);
ws2.AddImage(photoData, row: 3, column: 2, placement: ImagePlacement.InCell);
```

> Images are only supported in xlsx/xlsm format. xls/xlsb/csv do not support images. InCell images display as `#VALUE!` in Excel (consistent with natively produced Excel samples).
> Reading: **Floating images support read-back** (2.4.3+; opening xlsx/xlsm now fills `Images`). InCell image read-back will follow in a later version.

### Image Read-Back (2.4.3+)

After opening an xlsx/xlsm that contains floating images, `Worksheet.Images` is populated:

```csharp
var wb = Excel.Open("hasImage.xlsx");
foreach (var img in wb.Worksheets[0].Images)
{
    Console.WriteLine($"pos: {img.Row},{img.Column}  placement: {img.Placement}  size: {img.Anchor?.WidthPixels}x{img.Anchor?.HeightPixels}");
    Console.WriteLine($"altText: {img.AltText}  name: {img.Name}  bytes: {img.Data.Length}");
}
```

Supported fields: Row / Column (1-based), Placement (Floating only for now), Name, AltText, Anchor (TopLeftCell / offsets / size / MoveMode), Extension (png/jpg/gif/bmp auto), Data (bytes).

InCell image read-back is scheduled for a later version. Existing images are not removed on write-back.

### Image Anchor and Move Mode (2.4.1+)

Floating images support high-precision anchors: top-left cell + EMU offsets + display size + move/resize behavior with cells, plus accessibility alt text (AltText):

```csharp
ws.AddImage(logoData, new ImageAnchor
{
    TopLeftCell = "B2",           // top-left cell A1 reference
    TopLeftOffsetX = 9525,       // horizontal offset (EMU, 1px≈9525)
    TopLeftOffsetY = 0,           // vertical offset
    WidthPixels = 200,
    HeightPixels = 120,
    MoveMode = ImageMoveMode.MoveAndSizeWithCells, // move + size with cells
}, extension: "png", name: "logo", altText: "Company Logo");
```

`ImageMoveMode` three modes:

| Mode | OOXML | Behavior |
|---|---|---|
| `MoveButDontSizeWithCells` (default) | oneCellAnchor | Moves with cells, does not resize |
| `MoveAndSizeWithCells` | twoCellAnchor | Moves and resizes with cells (image stretches with grid) |
| `FixedPosition` | oneCellAnchor editAs="absolute" | Fixed position, does not move/resize with cells |

> `twoCellAnchor` end position is estimated using default column width≈64px / row height≈20px; with non-default grid the image scales with cells. `ImageAnchor` applies to Floating only; InCell ignores anchors.

---

## 15. Data Validation (Dropdown List)

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

## 16. Appending Data

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

## 17. List&lt;T&gt; Mapping (reflection, not AOT compatible)

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

## 18. DataTable Convenience API (AOT safe)

```csharp
var dt = new DataTable();
dt.Columns.Add("Name", typeof(string));
dt.Columns.Add("Age", typeof(int));
dt.Rows.Add("Zhang San", 25);
XlsxWriter.Write("data.xlsx", dt);

var read = XlsxReader.ReadAsDataTable("data.xlsx");
```

---

## 19. Stream Read/Write

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

## 20. Streaming Read & Progress Callback

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

## 21. Document Properties (Author/Time/Title)

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

## 22. File-Level Security (Open Password / Modify Password)

### Opening an Encrypted File

When reading an xlsx/xlsm/xlsb file that has an open password, provide the password via `ExcelReadOptions`:

```csharp
var wb = Excel.Open("encrypted.xlsx", new ExcelReadOptions
{
    OpenPassword = "your-password",            // open password
    ModifyPassword = "your-modify-password",   // modify password (optional)
});
```

> The encrypted samples in the repo (`files/` directory) follow the password convention: open password = `1`, modify password = `12`. For example `打开修改都需要密码.xlsx` (open=1, modify=12), `12.*` (modify only=12), `*.` (open only=1). Use your own passwords for production data.

### Reading the Security State

After opening, query the file's security state via `Workbook.Security`:

```csharp
var security = wb.Security;

bool hasOpenPwd = security.HasOpenPassword;     // has an open password
bool hasModPwd  = security.HasModifyPassword;   // has a modify password (write protection)
bool hasModAcc  = security.HasModifyAccess;     // modification authorized (optimistic authorization)
bool isReadOnly = security.IsReadOnly;          // read-only
bool canSave    = security.CanSave;             // can currently save
```

### Setting a Password

```csharp
// Set an open password (the file is encrypted on save)
wb.Security.SetOpenPassword("your-password");

// Set a modify password (write protection, fileSharing, not zip encryption)
wb.Security.SetModifyPassword("your-modify-password");

wb.SaveAs("protected.xlsx");
```

### Removing a Password

```csharp
// Remove the open password
wb.Security.RemoveOpenPassword();

// Remove the modify password (requires modify authorization, i.e. ModifyPassword was provided on open)
wb.Security.RemoveModifyPassword();

wb.SaveAs("plain.xlsx");
```

### Password Inheritance and Authorization Rules

- After opening an encrypted file, `SaveAs` **inherits the open password by default**; call `Security.RemoveOpenPassword()` before saving to remove it.
- A modify password is write protection (`<fileSharing>`), not zip encryption; providing `ModifyPassword` on read authorizes the modification (optimistic authorization, the sample value is not verified).
- Passwords **never** appear in exception messages, logs, or test output.
- Supported for xlsx/xlsm/xlsb only; csv/xls do not support passwords.

---

## 23. Error Handling

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

## 24. AOT Compatibility

### AOT-safe APIs (no reflection)

| API | Description |
|---|---|
| `Excel.Open` / `Workbook` / `Worksheet` / `Cell` / `ExcelRange` / `Cells` | object model |
| `Excel.CreateWriter` / `Excel.StreamRows` | streaming read/write |
| `Read(path/stream, ...)` | returns `SheetData` |
| `Write(path/stream, SheetData)` | accepts `SheetData` |
| `ReadAsDataTable(...)` | DataTable has its own schema |
| `Write(path, DataTable)` | DataTable write |
| `GetSheetNames(...)` | list sheet names |
| `Append(...)` | append |
| `AutoColumnWidths(...)` | auto column widths |

### AOT-unsafe APIs (reflection, marked `[RequiresUnreferencedCode]`)

| API | Description |
|---|---|
| `Excel.Read<T>(...)` / `Read<T>(...)` | List<T> read |
| `Excel.Write<T>(...)` / `Write<T>(...)` | List<T> write |

> AOT projects will get `IL3050`/`IL2026` warnings when calling these. Non-AOT projects (net48, net8 normal publish) are unaffected.

---

## 25. Conditional Formatting (2.4.3+)

Supports four rule types: cell comparison, expression, color scale, and data bar.

### Cell Comparison

```csharp
var ws = wb.Worksheets[0];
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.CellIs,
    Sqref = "B2:B100",
    Operator = ConditionalOperator.GreaterThan,
    Formula = "60",
    Style = new CellStyle { FontColor = "#FF0000" }, // red warning
});
```

### Expression (Formula)

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.Expression,
    Sqref = "A2:E100",
    Formula = "MOD(ROW(),2)=0",
    Style = new CellStyle { FillColor = "#F2F2F2" },
});
```

### Color Scale (2–3 colors)

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.ColorScale,
    Sqref = "D2:D100",
    ColorScale = new ColorScaleInfo
    {
        LowColor = "#FF0000",
        MidColor = "#FFFF00",
        HighColor = "#00FF00",
    },
});
```

### Data Bar

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.DataBar,
    Sqref = "E2:E100",
    DataBar = new DataBarInfo
    {
        Color = "#63C384",
        ShowValue = true,
        MinLengthPercent = 0,
        MaxLengthPercent = 100,
    },
});
```

### Reading

Opening an xlsx/xlsm file with conditional formats now populates `Worksheet.ConditionalFormats`. **xls/xlsb/csv do not support conditional formatting** — writes will report degradations via [§26 Capability Degradation Callback](#26-capability-degradation-callback-ondegradation-242). Additional rule types (`containsText`, `top10`, etc.) will follow in later versions.

---

## 26. Capability Degradation Callback (OnDegradation) (2.4.2+)

When saving to a format that does not support a capability (xls/xlsb/csv), those capabilities are silently dropped by default. Register the callback to receive a per-item list, so nothing is lost in silence.

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets[0];
ws.SetValue("A1", "x");
ws.Filter = new AutoFilter { Range = "A1:A10" };
ws.Comments = new Dictionary<string, string> { ["A1"] = "note" };
ws.FreezeRows = 1;
ws.Cell("B1").Style = new CellStyle { Bold = true };

var options = new ExcelWriteOptions
{
    OnDegradation = d =>
        Console.WriteLine($"[{d.Capability}] {d.SheetName} => {d.Message}")
};
Excel.Write("out.csv", wb, options);
```

Default `null`: behaves identically to previous versions (zero breakage).

### DegradationCapability enum

Comments, DataValidation, AutoFilter, Images, DocumentProperties, NamedRanges, Styles, MergedCells, FreezePanes, Hyperlinks, RowHeights, ColumnWidths, Formulas, Charts, PivotTables, RichData.

### Per-format degradation matrix (2.4.2)

| Capability | xlsx/xlsm | xlsb | xls | csv |
|---|---|---|---|---|
| Comments | supported | reported | reported | reported |
| Data validation | supported | reported | reported | reported |
| Auto filter | supported | reported | reported | reported |
| Images | write-back only | reported | reported | reported |
| Cell styles | supported | **NumberFormat only**, others reported | **NumberFormat only**, others reported | reported |
| Hyperlinks | supported | supported | supported | reported |
| Formulas | supported | cached value text | cached value text | reported (value text) |
| Row heights / column widths / merge / freeze | supported | supported | supported | reported |
| Charts / pivot tables / themes | round-trip preserved | round-trip preserved | not supported (never) | not supported |

`TargetFormat` is `ExcelFormat`; the library never mutates `DegradationInfo`.

---

## 26. Full API Reference

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
| `ReadProperties(string path)` / `ReadProperties(Stream stream)` | `WorkbookProperties` | read document properties (author/time/title) |

> ⚠️ Marked `[RequiresUnreferencedCode]`, not AOT compatible.

### XlsxWriter

| Method | Description |
|---|---|
| `Write(string path, SheetData data)` | write single sheet |
| `Write(string path, IReadOnlyList<SheetData> sheets)` | write multiple sheets |
| `Write(Stream stream, SheetData data)` | write single sheet to stream |
| `Write(Stream stream, IReadOnlyList<SheetData> sheets)` | write multiple sheets to stream |
| `Write(path/stream, sheets, WorkbookProperties? properties)` | write and carry document properties |
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
| `Cell` | cell (including the `Hyperlink` property, see §12) |
| `CellStyle` / `BorderStyle` / `BorderEdge` | styles |
| `CellRange` | range (merged cells) |
| `AutoFilter` / `FilterColumn` | auto filter |
| `DataValidation` | data validation |
| `Hyperlink` | cell hyperlink (`Target` / `Tooltip` / `IsInternal`) |
| `WorksheetImage` | worksheet image (added via `ws.AddImage`, see §14) |
| `WorkbookSecurity` | file security state (open password / modify password / read-only / can save) |
| `ExcelReadOptions` | read options (`OpenPassword` / `ModifyPassword`) |
| `WorkbookProperties` | document properties |
| `LiteExcelException` / `InvalidSheetNameException` | exceptions |
| `LiteColumnAttribute` | List<T> column attribute |
| `WriteOptions<T>` / `ReadOptions<T>` | List<T> configuration |

### Worksheet Members (hyperlinks / freeze panes / images)

| Member | Description |
|---|---|
| `Cell.Hyperlink` | get/set a cell hyperlink (see §12) |
| `Worksheet.FreezeRows` / `Worksheet.FreezeColumns` | freeze panes by rows/columns (see §13) |
| `ws.AddImage(byte[] data, int row, int col, int widthPx, int heightPx, ImagePlacement placement, string? extension, string? name)` | add an image to the sheet (see §14) |

### Enums

| Enum | Values |
|---|---|
| `CellType` | Text, Number, Date, Boolean, Empty |
| `HorizontalAlignment` | General, Left, Center, Right |
| `VerticalAlignment` | Top, Center, Bottom |
| `FilterType` | Equals, Compare, Contains, BeginsWith, EndsWith, Blank |
| `FilterOperator` | GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between |
| `DataValidationType` | List, WholeNumber, Decimal, Date |
| `ImagePlacement` | Floating, InCell |
