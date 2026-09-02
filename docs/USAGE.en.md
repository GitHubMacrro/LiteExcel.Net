# LiteExcel User Guide

> This manual reflects all public capabilities of the current mainline version of LiteExcel (object-model API + low-level API)
> Terminology conventions: **object-model API** refers to the everyday usage centered on the `Excel` → `Workbook` → `Worksheet` → `Cell`/`Cells`/`Range` chain; **low-level API** refers to the raw data entry points such as `SheetData` / `XlsxReader` / `XlsxWriter` / `CsvBackend` / `XlsxStreamWriter`, see Appendix B.

---

## 📚 Full Contents

| # | Chapter |
| :-: | :--- |
| **Getting Started** | |
| 1 | [Installation and References](#1-installation-and-references) |
| 2 | [Quick Start](#2-quick-start) |
| **Object Model** | |
| 3 | [File Navigation](#3-file-navigation) |
| 4 | [Cells and Values](#4-cells-and-values) |
| 5 | [Data Types and Conversion](#5-data-types-and-conversion) |
| 6 | [Styles](#6-styles) |
| 7 | [Merged Cells](#7-merged-cells) |
| 8 | [AutoFilter](#8-autofilter) |
| 9 | [Row Height and Column Width](#9-row-height-and-column-width) |
| 10 | [Comments](#10-comments) |
| 11 | [Hyperlinks](#11-hyperlinks) |
| 12 | [Freeze Panes](#12-freeze-panes) |
| 13 | [Images](#13-images) |
| 14 | [Data Validation](#14-data-validation) |
| 15 | [Conditional Formatting](#15-conditional-formatting) |
| 16 | [Excel Tables](#16-excel-tables) |
| 17 | [Named Ranges](#17-named-ranges) |
| 18 | [File-Level Passwords](#18-file-level-passwords) |
| 19 | [Worksheet and Workbook Protection](#19-worksheet-and-workbook-protection) |
| **Multi-Format and Platform** | |
| 20 | [Multi-Format Behavior](#20-multi-format-behavior) |
| 21 | [Streaming Read / Progress Callback / Append](#21-streaming-read--progress-callback--append) |
| 22 | [Degradation Callback OnDegradation](#22-degradation-callback-ondegradation) |
| 23 | [AOT Compatibility](#23-aot-compatibility) |
| **Operational Notes** | |
| 24 | [Exception Handling](#24-exception-handling) |
| **Appendices** | |
| A | [Object Model Quick Reference](#appendix-a-object-model-quick-reference) |
| B | [Low-Level API Reference](#appendix-b-low-level-api-reference) |

---


# 1. Installation and References

## 📑 Contents

| # | Section |
| :-: | :--- |
| 1.1 | [Acquiring the library](#11-acquiring-the-library) |
| 1.2 | [Preparation](#12-preparation) |

---

## 1.1 Acquiring the library

**NuGet Installation (recommended)**:

Install from the NuGet published package, suitable for most production projects:

```powershell
dotnet add package LiteExcel
```

You can also search for `LiteExcel` in the "Manage NuGet Packages" dialog of Visual Studio to install it.

**Local Reference from Source**:

When the package has not been published or you need to debug against the library source, reference it through a csproj project reference:

```xml
<ItemGroup>
  <ProjectReference Include="..\src\LiteExcel\LiteExcel.csproj" />
</ItemGroup>
```

> **Note**: Prefer the NuGet package for production projects; a source reference is only for local debugging against the library source or before a version has been released.

## 1.2 Preparation

**Namespace**:

All types reside in the `LiteExcel` namespace:

```csharp
using LiteExcel;
```

**Target Frameworks**:

The library targets **net48** and **net8.0** simultaneously. The net8.0 target additionally declares `IsAotCompatible=true`, and all public APIs are compatible with Native AOT / trimming (see Chapter 23).

---

# 2. Quick Start

The following uses the object-model API to complete the closed loop of "create → write values → save → open → read back":

```csharp
using LiteExcel;

var wb = Excel.Create();                       // Create a new workbook (contains Sheet1 by default)
var ws = wb.Worksheets["Sheet1"];

ws.SetValue("A1", "Name");                     // Header
ws.SetValue("B1", "Age");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B2", 25);
ws.SetValue("A3", "Li Si");
ws.SetValue("B3", 30);

wb.SaveAs("people.xlsx");                      // Save to disk

// Read back
var opened = Excel.Open("people.xlsx");
var sheet = opened.Worksheets[0];
Console.WriteLine(sheet.Cell("A2").GetString());   // Output: Zhang San
Console.WriteLine(sheet.Cell("B2").GetDouble());   // Output: 25
```

Output:

```
Zhang San
25
```

---


Chapters 3-19: the object-model API main line of capabilities; all day-to-day read/write lives here.

# 3. File Navigation

## 📑 Contents

| # | Section |
| :-: | :--- |
| 3.1 | [Open an Existing File](#31-open-an-existing-file) |
| 3.2 | [Open from a Stream](#32-open-from-a-stream) |
| 3.3 | [Read Options `ExcelReadOptions`](#33-read-options-excelreadoptions) |
| 3.4 | [Write Options `ExcelWriteOptions`](#34-write-options-excelwriteoptions) |
| 3.5 | [Create a New Workbook](#35-create-a-new-workbook) |
| 3.6 | [Save and Save As](#36-save-and-save-as) |
| 3.7 | [Format Enum `ExcelFormat`](#37-format-enum-excelformat) |
| 3.8 | [List Worksheet Names](#38-list-worksheet-names) |
| 3.9 | [Worksheet Management `Worksheets`](#39-worksheet-management-worksheets) |
| 3.10 | [Document Properties `WorkbookProperties`](#310-document-properties-workbookproperties) |

---

## 3.1 Open an Existing File

`Excel.Open` auto-detects the format by file extension and supports xlsx / xlsm / xls / xlsb / csv:

```csharp
var wb = Excel.Open("report.xlsx");            // auto-detected
var wb2 = Excel.Open("data.csv");              // auto-detected as Csv
var wb3 = Excel.Open("legacy.xls");            // auto-detected as Xls
```

You can also specify the format explicitly (useful when the extension does not match the content):

```csharp
var wb = Excel.Open("data.bin", ExcelFormat.Xlsx);
```

Output: (this example has no console output)

## 3.2 Open from a Stream

A stream has no extension, so the format **must be specified explicitly**. The input stream is not closed (managed by the caller); a non-seekable stream (e.g., a network stream) is internally copied to memory:

```csharp
using var fs = File.OpenRead("report.xlsx");
var wb = Excel.Open(fs, ExcelFormat.Xlsx);
// after opening, CurrentPath is null, so use SaveAs to specify the save path
wb.SaveAs("copy.xlsx");
```

Output:

```
written to copy.xlsx
```

Read all sheets from a stream: `Excel.Open(stream, ExcelFormat.Xlsx).Worksheets` directly iterates the multi-sheet worksheets.

## 3.3 Read Options `ExcelReadOptions`

The second parameter of `Open` can pass read options, including password, merged-cell fill, CSV separator, etc.:

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",        // open password (file encryption)
    ModifyPassword = "write",       // modify password (write protection, used to obtain edit permission)
    FillMergedCells = true,         // expand the top-left value of a merged range to the whole range
    Separator = ';',                // only applies to CSV; auto-detected when null
});
```

| Parameter | Type | Description |
|---|---|---|
| `OpenPassword` | `string?` | Open password (file encryption); decrypts password-protected xlsx/xlsm/xlsb |
| `ModifyPassword` | `string?` | Modify password (write protection); provides edit/save permission once supplied |
| `FillMergedCells` | `bool` | Expand the top-left value of a merged range to the whole merged range; default `false` |
| `Separator` | `char?` | Only applies to CSV; auto-detected when `null` |
| `ReadStyles` | `bool` | Whether to read styles; default `true` |
| `LeaveOpen` | `bool` | Whether to keep the input stream open after reading completes on the Stream overload; default `false` |

Output: (this example has no console output)

## 3.4 Write Options `ExcelWriteOptions`

The second parameter of `Excel.Write` can pass write options:

```csharp
Excel.Write("out.xlsx", wb, new ExcelWriteOptions
{
    Overwrite = true,           // whether to overwrite when the target file already exists, default true
    AutoFitColumns = true,      // auto-estimate column widths before writing, default false
    FreezeHeader = true,        // freeze the header row when writing, default false
    Properties = new WorkbookProperties { Creator = "Me" },  // override document properties
    OnDegradation = info => Console.WriteLine(info.Capability),  // degradation callback
    Separator = ';',            // only applies to CSV; defaults to comma when null
});
```

| Parameter | Type | Description |
|---|---|---|
| `Overwrite` | `bool` | Whether to overwrite when the target file already exists; default `true` |
| `AutoFitColumns` | `bool` | Auto-estimate column widths before writing; default `false` |
| `FreezeHeader` | `bool` | Freeze the header row when writing; default `false` |
| `Properties` | `WorkbookProperties?` | Override the workbook document properties |
| `OnDegradation` | `Action<DegradationInfo>?` | Capability-degradation callback (reported item by item when writing to a format that does not support a capability; default `null`) |
| `Separator` | `char?` | Only applies to CSV; defaults to comma when `null` |
| `LeaveOpen` | `bool` | Whether to keep the output stream open after writing completes on the Stream overload; default `false` |

Output:

```
written to out.xlsx
```

## 3.5 Create a New Workbook

`Excel.Create` has several overloads:

```csharp
var wb1 = Excel.Create();                    // empty workbook, default Sheet1
Console.WriteLine(wb1.Worksheets[0].Name);   // print to verify: default sheet name
var wb2 = Excel.Create("Data");              // specify the first worksheet name
var wb3 = Excel.Create(new[] { "Q1", "Q2", "Q3" });   // batch-create sheets
var wb4 = Excel.Create(ExcelFormat.Xlsm);    // specify the format
```

Create a workbook and write data in one step (List\<T\> DataTable):

```csharp
var people = new List<Person> { new() { Name = "A", Age = 1 } };
var wb5 = Excel.Create(people, "People");    // first row is the header

var dt = new System.Data.DataTable("T");
dt.Columns.Add("X");
dt.Rows.Add("v");
var wb6 = Excel.Create(dt);                  // when sheetName is empty, uses TableName
```

Output:

```
Sheet1
```

## 3.6 Save and Save As

`Workbook.Save` saves to the current path (throws `LiteExcelException` when a newly created workbook has no path); `SaveAs` specifies the path:

```csharp
wb.Save();                    // save to CurrentPath
wb.SaveAs("out.xlsx");        // save as, keeping the current format
wb.SaveAs("out.xlsm", ExcelFormat.Xlsm);   // save as and convert format
wb.Save(new FileStream("s.xlsx", FileMode.Create), ExcelFormat.Xlsx);  // save to a stream
```

Output:

```
written to out.xlsx / out.xlsm / s.xlsx
```

> ⚠️ **Important limits**
> This library **does not create or edit** charts (Chart) or pivot tables (PivotTable).
> Opening xlsx / xlsm / xlsb and saving again preserves these elements as-is (passthrough); xls / csv have no preservation mechanism, so they are lost on open-then-save. Make a backup copy before overwriting the source file.

- `SaveAs(path, format)` requires the path extension to match the format, otherwise it throws `LiteExcelException` (to avoid writing content that does not match the extension, producing a file Excel cannot open).
- A workbook containing VBA macros cannot be saved to a format that does not support macros (xlsx / xls); it errors out early.
- File-level passwords (open / modify) are only supported for xlsx / xlsm / xlsb; saving to csv / xls with a password set will error.

## 3.7 Format Enum `ExcelFormat`

```csharp
public enum ExcelFormat { Xlsx, Xlsm, Xlsb, Xls, Csv }
```

- `Excel.DetectFormat(path)` returns the format based on the extension.
- `Workbook.Format` returns the current workbook format.

Output: (this example has no console output)

## 3.8 List Worksheet Names

```csharp
var names = Excel.GetSheetNames("report.xlsx");   // List<string>
using var stream = File.OpenRead("report.xlsx");
var names2 = Excel.GetSheetNames(stream);         // only xlsx/xlsm (XML metadata of the zip container)
```

> `Excel.GetSheetNames(path)` reads the workbook metadata directly for xlsx / xlsm (lightweight); for **xlsb / xls / csv** it goes through `Excel.Open` to parse by format and also returns the sheet names correctly (xlsb returns its internal sheet names; csv returns the single sheet "Sheet1"). `GetSheetNames(Stream)` only supports xlsx / xlsm; to get xlsb sheet names from a stream, use `Excel.Open(stream, ExcelFormat.Xlsb).Worksheets.Names`.

Output: (returns a List<string>)

## 3.9 Worksheet Management `Worksheets`

`Workbook.Worksheets` is a `WorksheetCollection` supporting add/remove/move, index and name access:

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];         // access by name (throws LiteExcelException if missing)
var first = wb.Worksheets[0];             // access by index (0-based)
Console.WriteLine(wb.Worksheets.Count);   // count
Console.WriteLine(string.Join(", ", wb.Worksheets.Names));   // all sheet names
Console.WriteLine(wb.Worksheets.Contains("Sheet1"));         // true

wb.Worksheets.Add("Data");                // add an empty sheet
wb.Worksheets.Add("People", peopleList);  // add a sheet and write List<T> (first row is the header)
wb.Worksheets.Add("T", dataTable);        // add a sheet and write DataTable (first row is column names)
wb.Worksheets.Move(0, 2);                 // reorder (0-based)
wb.Worksheets.Remove("Data");             // remove by name (returns true if present)
wb.Worksheets.RemoveAt(0);                // remove by index
```

Iterate over worksheets:

```csharp
foreach (var sheet in wb.Worksheets)
    Console.WriteLine(sheet.Name);
```

Output:

```
1
Sheet1
True
Sheet1
T
```

> ⚠️ Sheet-name validation rules are described in Chapter 24: illegal sheet names (containing `\ / ? * [ ] :`, longer than 31 characters, etc.) throw `InvalidSheetNameException` when saving.

## 3.10 Document Properties `WorkbookProperties`

> ⚠️ Document properties are supported only for **xlsx / xlsm / xlsb** (OLE property sets are not implemented for xls). When writing to xls they are **silently dropped**, reported via `OnDegradation`.

`Workbook.Properties` corresponds to `docProps/core.xml` and `docProps/app.xml` inside the xlsx package:

```csharp
var wb = Excel.Create();
wb.Properties.Creator = "JackZ";
wb.Properties.LastModifiedBy = "Me";
wb.Properties.Created = DateTime.Now;
wb.Properties.Modified = DateTime.Now;
wb.Properties.Title = "季度报告";
wb.Properties.Subject = "财务";
wb.Properties.Application = "LiteExcel";
wb.SaveAs("props.xlsx");
```

| Parameter | Type | Description |
|---|---|---|
| `Creator` | `string?` | Author (dc:creator) |
| `LastModifiedBy` | `string?` | Last saved by (cp:lastModifiedBy) |
| `Created` | `DateTime?` | Creation time |
| `Modified` | `DateTime?` | Last modification time |
| `Title` | `string?` | Title |
| `Subject` | `string?` | Subject |
| `Application` | `string?` | Application name; when `null`, the host assembly name is used when writing |

Override properties when writing (`ExcelWriteOptions.Properties`):

```csharp
Excel.Write("out.xlsx", wb, new ExcelWriteOptions
{
    Properties = new WorkbookProperties { Creator = "Bot", Title = "Auto" },
});
```

Read back document properties:

```csharp
var opened = Excel.Open("props.xlsx");
Console.WriteLine(opened.Properties.Creator);    // JackZ
Console.WriteLine(opened.Properties.Title);      // 季度报告
```

Output:

```
JackZ
季度报告
```

---

# 4. Cells and Values

## 📑 Contents

| # | Section |
| :-: | :--- |
| 4.1 | [Accessing Cells by Coordinate / Address](#41-accessing-cells-by-coordinate--address) |
| 4.2 | [Setting Values with `SetValue`](#42-setting-values-with-setvalue) |
| 4.3 | [Collection-Style Access with `Cells`](#43-collection-style-access-with-cells) |
| 4.4 | [Range Operations with `ExcelRange`](#44-range-operations-with-excelrange) |
| 4.5 | [Cell Read Methods](#45-cell-read-methods) |
| 4.6 | [The `Value` Property](#46-the-value-property) |

---

## 4.1 Accessing Cells by Coordinate / Address

Coordinates are uniformly **1-based**. `Worksheet.Cell` provides access by row/column or by A1 address:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

var c1 = ws.Cell(1, 1);        // A1
var c2 = ws.Cell("B3");        // B3
Console.WriteLine(c1.IsEmpty);   // True (not yet assigned)
```

`Cell` is a reference type; accessing a nonexistent cell returns an `Empty` placeholder; once assigned it falls into the grid.

Output:

```
True
```

## 4.2 Setting Values with `SetValue`

`SetValue` automatically expands the grid when out of bounds; `null` / `DBNull` writes an empty cell:

```csharp
ws.SetValue(1, 1, "Header");
ws.SetValue("A2", 42);
ws.SetValue("B2", null);       // clears B2
Console.WriteLine(ws.Cell("A2").GetString());   // 42
```

Output:

```
42
```

## 4.3 Collection-Style Access with `Cells`

`Worksheet.Cells` provides a whole-sheet entry point, supporting indexers, range extraction, enumeration, and bulk clearing:

```csharp
ws.Cells[1, 1] = "A1 via cells";      // indexer (row, column)
ws.Cells["B1"] = "B1 via cells";      // indexer (A1 address)
ws.Cells.SetValue("C1", 3.14);        // convenient value write

var range = ws.Cells.Range("A1:D10"); // extract range
foreach (var cell in ws.Cells)        // enumerate existing cells in the grid
    Console.WriteLine(cell.GetString());
ws.Cells.Clear();                     // clear all values in the sheet (keeps rows/columns)
```

Output:

```
A1 via cells
B1 via cells
3.14
```

## 4.4 Range Operations with `ExcelRange`

`Worksheet.Range` returns a contiguous rectangular range (1-based, inclusive), supporting batch read/write, styles, merge, clear, and enumeration:

```csharp
var r = ws.Range("A1:D10");          // or ws.Range(1, 1, 10, 4)
Console.WriteLine(r.Address);        // "A1:D10"
Console.WriteLine($"{r.RowCount} x {r.ColumnCount}");  // 10 x 4

r.Fill("x");                          // fill the whole range with the same value
r.Fill(new object?[,] { { "a", "b" }, { "c", "d" } }); // write a 2D array
var vals = r.ToValues();              // read back as object?[,]
var cells = r.ToCells();              // read back as Cell[,]
r.Style = new CellStyle { Bold = true };  // uniform style for the whole range
r.Merge();                            // merge the range
r.Unmerge();                          // unmerge
r.Clear();                            // clear values within the range

var single = r.Cell(0, 0);            // relative offset within the range (0-based)
```

Output:

```
A1:D10
10 x 4
```

## 4.5 Cell Read Methods

`Cell` provides strongly typed and Try-style reads:

```csharp
var cell = ws.Cell("A1");
string? s = cell.GetString();        // text / number / date / bool formatted by convention
double d = cell.GetDouble();         // throws InvalidCastException on type mismatch
DateTime dt = cell.GetDateTime();
bool b = cell.GetBoolean();
object? raw = cell.GetValue();       // raw object; Empty returns null

bool ok = cell.TryGetString(out var s2);   // empty cell returns false
bool ok2 = cell.TryGetDouble(out double d2);
bool ok3 = cell.TryGetDateTime(out DateTime dt2);
bool ok4 = cell.TryGetBoolean(out bool b2);
```

Output: (cell's raw value)

## 4.6 The `Value` Property

`Cell.Value` and `ExcelRange.Value` are convenient property wrappers around `SetValue` / `GetValue`, matching the Excel interop idiom. Reading returns `object?`; for type safety use `GetString()` / `GetDouble()` etc. from 4.5.

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// Single cell: read/write a scalar (via Cell / Cells indexer / single-cell Range)
ws.Cell("A1").Value = "123";
Console.WriteLine(ws.Cell("A1").Value);        // 123
ws.Cells["B2"].Value = 42;
ws.Cells[3, 3].Value = true;
ws.Range("C1").Value = "written via Range";

// Multi-cell Range: read returns object?[,] (sized to the range); write a scalar fills, a 2D array fills cell by cell
ws.Range("A10:B11").Value = "x";                          // all 4 cells are x
var grid = (object?[,])ws.Range("A10:B11").Value;         // read as 2D
ws.Range("D1:E2").Value = new object?[,] { { 1, 2 }, { 3, 4 } };
```

`ExcelRange.Value` behavior: a single-cell range (1×1) reads/writes a scalar; a multi-cell range reads as `object?[,]`, writing a scalar is equivalent to `Fill`, and writing a 2D array is equivalent to `Fill(object?[,])`.

---

# 5. Data Types and Conversion

## 📑 Contents

| # | Section |
| :-: | :--- |
| 5.1 | [Cell Type `CellType`](#51-cell-type-celltype) |
| 5.2 | [Factory Methods](#52-factory-methods) |
| 5.3 | [Automatic Type Conversion](#53-automatic-type-conversion) |
| 5.4 | [Nullable Types](#54-nullable-types) |
| 5.5 | [Number Format Quick Reference](#55-number-format-quick-reference) |
| 5.6 | [Automatic Date Detection on Read](#56-automatic-date-detection-on-read) |
| 5.7 | [Formulas](#57-formulas) |
| 5.8 | [Byte[]](#58-byte) |
| 5.9 | [List\<T\> Mapping and `[LiteColumn]`](#59-listt-mapping-and-litecolumn) |
| 5.10 | [List\<T\> Fluent Configuration (WriteOptions\<T\> / ReadOptions\<T\>)](#510-listt-fluent-configuration-writeoptionst--readoptionst) |
| 5.11 | [DataTable Convenience API](#511-datatable-convenience-api) |

---

## 5.1 Cell Type `CellType`

```csharp
public enum CellType { Text, Number, Date, Boolean, Empty }
```

`Cell.Type` determines which value field is valid: `Text` / `Number` / `Date` / `Boolean`. `IsEmpty` indicates an empty cell.

Output: (this example has no console output)

## 5.2 Factory Methods

```csharp
var t = Cell.FromText("hello");
var n = Cell.FromNumber(42, "#,##0.00");      // can carry a number format
var d = Cell.FromDate(new DateTime(2024, 1, 1));  // default format yyyy-MM-dd
var b = Cell.FromBoolean(true);
var f = Cell.FromFormula("SUM(A1:A3)");        // formula cell
var e = Cell.Empty;
Console.WriteLine(d.GetString());   // default date format yyyy-MM-dd
```

Output:

```
2024-01-01
```

## 5.3 Automatic Type Conversion

`SetValue(object?)` auto-maps based on the CLR type:

| CLR type | Cell type |
|---|---|
| `bool` | `Boolean` |
| `DateTime` | `Date` (default `yyyy-MM-dd`) |
| `sbyte/byte/short/ushort/int/uint/long/ulong/float/double/decimal` | `Number` |
| `null` / `DBNull` | `Empty` |
| others (including `string`) | `Text` |

```csharp
ws.SetValue("A1", 3.14);        // Number
ws.SetValue("A2", true);        // Boolean
ws.SetValue("A3", DateTime.Now); // Date
ws.SetValue("A4", "text");      // Text
ws.SetValue("A5", null);        // Empty

Console.WriteLine(ws.Cell("A1").Type);   // Number
Console.WriteLine(ws.Cell("A2").Type);   // Boolean
Console.WriteLine(ws.Cell("A3").Type);   // Date
Console.WriteLine(ws.Cell("A4").Type);   // Text
Console.WriteLine(ws.Cell("A5").Type);   // Empty
```

Output:

```
Number
Boolean
Date
Text
Empty
```

## 5.4 Nullable Types

`SetValue` accepts `object?`; nullable value types (`int?` / `DateTime?`, etc.) are handled by their underlying value after boxing; `null` writes an empty cell:

```csharp
int? maybe = null;
ws.SetValue("A1", maybe);        // writes empty
maybe = 7;
ws.SetValue("A2", maybe);        // Number 7
Console.WriteLine(ws.Cell("A1").IsEmpty);     // True (null writes empty)
Console.WriteLine(ws.Cell("A2").GetDouble()); // 7
```

Output:

```
True
7
```

## 5.5 Number Format Quick Reference

`NumberFormat` uses Excel format code strings; common examples:

| Format code | Effect |
|---|---|
| `"0"` | integer |
| `"0.00"` | two decimal places |
| `"#,##0.00"` | thousands separator + two decimal places |
| `"0%"` | percentage |
| `"yyyy/m/d"` / `"yyyy-MM-dd"` | date |
| `"hh:mm"` | time |
| `"@"` | text |

```csharp
ws.Cell("A1").SetValue(12345.678);
ws.Cell("A1").NumberFormat = "#,##0.00";   // displays 12,345.68
Console.WriteLine(ws.Cell("A1").GetString());  // read back per the format
```

Output:

```
12,345.68
```

## 5.6 Automatic Date Detection on Read

When reading xlsx / xlsm / xlsb, if a cell's number format is a built-in Excel date format (IDs 14-22, 27-36, 45-47, 50-58, etc.), it is automatically read as `CellType.Date`:

```csharp
var opened = Excel.Open("dates.xlsx");
var cell = opened.Worksheets[0].Cell("A1");
if (cell.Type == CellType.Date)
    Console.WriteLine(cell.GetDateTime().ToString("yyyy-MM-dd"));
```

The 1904 date system flag (`Date1904`) captured on open is written back to the corresponding format flag on save, keeping the date serial-value base consistent.

Output:

```
2024-01-01
```

## 5.7 Formulas

`Cell.Formula` is a field separate from the cached value (the formula string no longer occupies `Text`, avoiding overwriting the cached result value of text formulas). The old style of writing the formula into `Text` and setting `IsFormula=true` remains compatible:

```csharp
var cell = ws.Cell("C1");
cell.Formula = "SUM(A1:B1)";     // writes only the formula string, does not compute the result
cell.IsFormula = true;           // treated as a formula on write (compatibility shim)

// or directly assign a formula Cell to a cell (SetValue accepts a Cell and copies its content)
ws.Cell("C2").SetValue(Cell.FromFormula("A1*2"));
Console.WriteLine(cell.Formula);             // SUM(A1:B1)
Console.WriteLine(ws.Cell("C2").Formula);    // A1*2
```

Output:

```
SUM(A1:B1)
A1*2
```

In List\<T\> mapping, `[LiteColumn(IsFormula = true)]` treats a string property as a formula column (see section 5.9).

## 5.8 Byte[]

`SetValue` treats any non-numeric type as `Text` (`value.ToString()`). **For binary data, use the image API** (`Worksheet.AddImage`, see chapter 13) or encode it to text yourself. The library itself does not map `byte[]` to a binary cell type.

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", Convert.ToBase64String(new byte[] { 1, 2, 3 }));  // encode to text yourself
Console.WriteLine(ws.Cell("A1").GetString());                       // AQID
```

Output:

```
AQID
```

## 5.9 List\<T\> Mapping and `[LiteColumn]`

The `[LiteColumn]` attribute controls column name / order / format / ignore / formula during List\<T\> mapping:

```csharp
public class Person
{
    [LiteColumn(Name = "姓名", Order = 0)]
    public string Name { get; set; } = "";

    [LiteColumn(Name = "年龄", Order = 1, Format = "0")]
    public int Age { get; set; }

    [LiteColumn(Name = "总额", Order = 2, Format = "#,##0.00", IsFormula = true)]
    public string Total { get; set; } = "";   // value may or may not have a leading "="

    [LiteColumn(Ignore = true)]
    public string Secret { get; set; } = "";  // not output
}
```

```csharp
var people = new List<Person> { new() { Name = "张三", Age = 30, Total = "=100*2", Secret = "隐藏" } };
var wb = Excel.Create(people, "People");   // headers: 姓名 | 年龄 | 总额; Secret column ignored
wb.SaveAs("people.xlsx");
```

Output:

```
written to people.xlsx
```

CLR types auto-converted by List\<T\> mapping:

| CLR type | Cell type | Description |
|---|---|---|
| `int` / `long` / `short` / `byte` | `Number` | integer |
| `double` / `float` / `decimal` | `Number` | decimal |
| `DateTime` | `Date` | date-time |
| `bool` | `Boolean` | boolean |
| `string` | `Text` | text |

All the above types support nullable versions (`int?` / `DateTime?`, etc.); null writes an empty cell.

## 5.10 List\<T\> Fluent Configuration (WriteOptions\<T\> / ReadOptions\<T\>)

Besides the `[LiteColumn]` attribute, Fluent API and dictionary mapping are also supported, suitable for ad-hoc column name / format / ignore / formula adjustments:

```csharp
// Fluent configuration when writing
Excel.Write("people.xlsx", people, "Employees", opt => opt
    .Column(p => p.Name, "姓名")                    // specify column name
    .Column(p => p.Age, "年龄", format: "0")        // specify column name + number format
    .Column(p => p.Total, "总额", isFormula: true)  // formula column (value may or may not have a leading "=")
    .Ignore(p => p.Secret)                          // ignore property
);

// Fluent configuration when reading
var list = Excel.Read<Person>("people.xlsx", "Employees", opt => opt
    .Column(p => p.Name, "姓名")                    // specify header name -> property mapping
    .Column(p => p.Age, "年龄")
);

// Dictionary mapping (common in legacy projects); configure uses named parameters, sheetName defaults to "Sheet1"
Excel.Write("people.xlsx", people, configure: opt => opt
    .Map(new Dictionary<string, string> { { "Name", "姓名" }, { "Age", "年龄" } })
);
```

Output: (directly writes out a people.xlsx file)

## 5.11 DataTable Convenience API

DataTable carries its own column structure, **no reflection required** (does not trigger reflection-based mapping), AOT safe. The first row is automatically written as column names:

```csharp
var dt = new DataTable("订单");
dt.Columns.Add("OrderID", typeof(int));
dt.Columns.Add("Customer", typeof(string));
dt.Columns.Add("Amount", typeof(decimal));
dt.Columns.Add("Date", typeof(DateTime));
dt.Rows.Add(1001, "Alice", 599.99m, new DateTime(2024, 6, 1));
dt.Rows.Add(1002, "Bob", 1299.50m, new DateTime(2024, 6, 15));

Excel.Write("orders.xlsx", dt, "Orders");   // write in one step

var back = Excel.ReadAsDataTable("orders.xlsx", "Orders");   // read back (first row is header)
foreach (DataRow row in back.Rows)
    Console.WriteLine($"#{row["OrderID"]} | {row["Customer"]} | {row["Amount"]:0.00}");

var wb = Excel.Create(dt);        // create a workbook in one step; sheetName defaults to DataTable.TableName (then Sheet1 if empty)
Console.WriteLine("sheet: " + wb.Worksheets[0].Name);
wb.SaveAs("orders2.xlsx");

var opened = Excel.Open("orders.xlsx");
// import into an existing sheet: clear existing content, then rebuild from A1; includeHeader=false does not write the column-name row
opened.Worksheets[0].ImportData(dt, includeHeader: false);
opened.SaveAs("orders3.xlsx");
Console.WriteLine("imported rows: " + Excel.ReadAsDataTable("orders3.xlsx", "Orders", firstRowIsHeader: false).Rows.Count);
```

Important parameters:

| Parameter | Type | Description |
|---|---|---|
| `Excel.Write` | `sheetName` | `string` | target sheet name, defaults to `"Sheet1"` |
| `Excel.Write` | `options` | `ExcelWriteOptions?` | write options (see 3.4) |
| `Excel.Create` | `sheetName` | `string?` | uses `DataTable.TableName` if empty, then `"Sheet1"` if also empty |
| `Excel.ReadAsDataTable` | `sheetName` | `string?` | target sheet; null reads the first sheet |
| `Excel.ReadAsDataTable` | `firstRowIsHeader` | `bool` | whether the first row is used as column names, default `true` |
| `ImportData` | `includeHeader` | `bool` | whether to write the column-name row, default `true`; import clears the whole sheet then rebuilds from A1 |

Output:

```
#1001 | Alice | 599.99
#1002 | Bob | 1299.50
sheet: 订单
imported rows: 2
```

> ⚠️ The DataTable path does not go through reflection-based mapping (`[LiteColumn]` / Fluent configuration does not apply); the first row is always the column names (except with `includeHeader=false`).

---

# 6. Styles

## 📑 Contents

| # | Section |
| :-: | :--- |
| 6.1 | [Cell Style `CellStyle`](#61-cell-style-cellstyle) |
| 6.2 | [Borders `BorderStyle` / `BorderEdge`](#62-borders-borderstyle--borderedge) |
| 6.3 | [Alignment and Wrapping](#63-alignment-and-wrapping) |
| 6.4 | [Setting Cell / Range Styles (object-model API)](#64-setting-cell--range-styles-object-model-api) |
| 6.5 | [Header Style `HeaderStyle`](#65-header-style-headerstyle) |
| 6.6 | [Whole-Sheet Default Style `DefaultStyle`](#66-whole-sheet-default-style-defaultstyle) |
| 6.7 | [Row-Level Styles `RowStyles`](#67-row-level-styles-rowstyles) |
| 6.8 | [Column-Level Styles `ColumnStyles`](#68-column-level-styles-columnstyles) |
| 6.9 | [Style Priority (Overriding)](#69-style-priority-overriding) |

---

## 6.1 Cell Style `CellStyle`

Colors uniformly use the `"#RRGGBB"` format:

```csharp
var style = new CellStyle
{
    FontName = "Arial",
    FontSize = 12,
    Bold = true,
    Italic = true,
    Underline = false,
    Strikeout = false,
    FontColor = "#FF0000",
    FillColor = "#FFFF00",
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    WrapText = true,
    NumberFormat = "#,##0.00",   // used for dxf read-back (table column format / conditional formatting)
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
    },
};
```

| Parameter | Type | Description |
|---|---|---|
| `FontName` | `string?` | font name, e.g. `"Arial"` |
| `FontSize` | `double` | font size, default 11 |
| `Bold` | `bool` | bold |
| `Italic` | `bool` | italic |
| `Underline` | `bool` | underline |
| `Strikeout` | `bool` | strikethrough |
| `FontColor` | `string?` | font color, `"#RRGGBB"` format |
| `FillColor` | `string?` | fill color, `"#RRGGBB"` format |
| `HorizontalAlignment` | `HorizontalAlignment` | horizontal alignment, default `General` |
| `VerticalAlignment` | `VerticalAlignment` | vertical alignment, default `Bottom` |
| `WrapText` | `bool` | automatic wrapping |
| `NumberFormat` | `string?` | number format code (used for dxf read-back) |
| `Border` | `BorderStyle?` | border |

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("A1").Style = style;
Console.WriteLine(ws.Cell("A1").Style.Bold);   // True
```

Output:

```
True
```

## 6.2 Borders `BorderStyle` / `BorderEdge`

```csharp
var style = new CellStyle
{
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
        Bottom = new BorderEdge { Style = "double" },
        Left = new BorderEdge { Style = "medium", Color = "#333333" },
        Right = new BorderEdge { Style = "dashed" },
    },
};
```

`BorderEdge.Style` is a string; common values: `thin` / `medium` / `thick` / `double` / `dashed` / `dotted`, etc.

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("A1").Style = style;
Console.WriteLine(ws.Cell("A1").Style.Border.Top.Style);   // thin
```

Output:

```
thin
```

## 6.3 Alignment and Wrapping

```csharp
public enum HorizontalAlignment { General, Left, Center, Right }
public enum VerticalAlignment { Top, Center, Bottom }
```

`WrapText = true` enables automatic wrapping within the cell.

Output: (this example has no console output)

## 6.4 Setting Cell / Range Styles (object-model API)

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// single cell
ws.Cell("A1").Style = new CellStyle { Bold = true, FillColor = "#D9E1F2" };
ws.Cell("A1").NumberFormat = "yyyy/m/d";

// uniform style for a range
ws.Range("A1:C10").Style = new CellStyle { HorizontalAlignment = HorizontalAlignment.Center };
Console.WriteLine(ws.Cell("A1").NumberFormat);   // yyyy/m/d
```

Output:

```
yyyy/m/d
```

## 6.5 Header Style `HeaderStyle`

Applies to the `SheetData.Headers` header row.

> ⚠️ **Prerequisite**: `HeaderStyle` only takes effect when there is a separate header row (List\<T\> / DataTable writes, or low-level `SheetData.Headers` (see Appendix B.1)). Since the object model is a "whole-sheet grid" model (all rows written via `ws.SetValue` count as data rows, with no separate header row), setting `ws.HeaderStyle = ...` directly has **no** effect. To style the first row in the object model, use `RowStyles` to target row 0:

```csharp
// Option 1: low-level / List<T> / DataTable paths (have Headers)
ws.HeaderStyle = new CellStyle { Bold = true, FillColor = "#4472C4", FontColor = "#FFFFFF" };
```

```csharp
// Option 2: object-model grid (treat the first row as header -> use RowStyles for row 0)
ws.SetValue("A1", "Name"); ws.SetValue("B1", "Age");
ws.RowStyles = new Dictionary<int, CellStyle>
{
    { 0, new CellStyle { Bold = true, FillColor = "#4472C4", FontColor = "#FFFFFF" } },
};
Console.WriteLine(ws.RowStyles[0].Bold);   // True
```

Output:

```
True
```

## 6.6 Whole-Sheet Default Style `DefaultStyle`

Has the lowest priority:

```csharp
ws.DefaultStyle = new CellStyle { FontName = "Consolas", FontSize = 10 };
Console.WriteLine(ws.DefaultStyle.FontName);   // Consolas
```

Output:

```
Consolas
```

## 6.7 Row-Level Styles `RowStyles`

The key is a **0-based row index**:

```csharp
ws.RowStyles = new Dictionary<int, CellStyle>
{
    { 1, new CellStyle { FillColor = "#FCE4D6" } },   // row 2 (0-based 1)
};
Console.WriteLine(ws.RowStyles[1].FillColor);   // #FCE4D6
```

Output:

```
#FCE4D6
```

## 6.8 Column-Level Styles `ColumnStyles`

The key is a **0-based column index**:

```csharp
ws.ColumnStyles = new Dictionary<int, CellStyle>
{
    { 2, new CellStyle { HorizontalAlignment = HorizontalAlignment.Right } },  // column 3
};
Console.WriteLine(ws.ColumnStyles[2].HorizontalAlignment);   // Right
```

Output:

```
Right
```

## 6.9 Style Priority (Overriding)

When writing, resolution follows this priority (**row/column-level style precedence is more explicit**):

- **Data rows**: `Cell.Style` > `RowStyle` > `ColumnStyle` > `DefaultStyle`
- **Header row**: `HeaderStyle` > `ColumnStyle` > `DefaultStyle`

```csharp
// Example: cell style overrides row style, row style overrides column style, column style overrides default style
ws.DefaultStyle = new CellStyle { FontSize = 10 };
ws.ColumnStyles = new Dictionary<int, CellStyle> { { 0, new CellStyle { Bold = true } } };
ws.RowStyles = new Dictionary<int, CellStyle> { { 0, new CellStyle { Italic = true } } };
ws.Cell("A1").Style = new CellStyle { Underline = true };
// A1 final: Underline (cell) + Italic (row) + Bold (column) + FontSize 10 (default)
```

Output: (this example has no console output)

![Styles and number formats](screenshots/style_number.png)

*Produced by the example code in this chapter, opened in Excel: font, color, border, alignment; currency / percent / date formats.*
---

# 7. Merged Cells

## 📑 Contents

| # | Section |
| :-: | :--- |
| 7.1 | [Writing Merged Cells](#71-writing-merged-cells) |
| 7.2 | [Unmerging](#72-unmerging) |
| 7.3 | [Reading Merged Ranges](#73-reading-merged-ranges) |
| 7.4 | [Filling Merged Ranges](#74-filling-merged-ranges) |

---

## 7.1 Writing Merged Cells

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.SetValue("A1", "Merged Title");
ws.Merge("A1:D1");                    // A1 address
ws.Merge(2, 1, 2, 3);                 // 1-based row/column (merges A2:C2)
ws.Merge("A3:B4");

// Range-based approach
ws.Range("C5:E5").Merge();
Console.WriteLine(ws.MergedRanges.Count);   // 4
```

Output:

```
4
```

## 7.2 Unmerging

```csharp
ws.Unmerge("A1:D1");
ws.Range("C5:E5").Unmerge();
```

Output: (this example has no console output)

## 7.3 Reading Merged Ranges

`Worksheet.MergedRanges` returns `IReadOnlyList<CellRange>` (**0-based**, consistent with the low-level model (see Appendix B.1)):

```csharp
var opened = Excel.Open("merged.xlsx");
var ws = opened.Worksheets[0];
foreach (var m in ws.MergedRanges)
    Console.WriteLine($"{m.FirstRow},{m.FirstCol} - {m.LastRow},{m.LastCol}");
```

Output:

```
0,0 - 0,3
1,0 - 1,2
2,0 - 3,1
4,2 - 4,4
```

## 7.4 Filling Merged Ranges

Setting `FillMergedCells = true` when reading expands the top-left value across the entire merged range:

```csharp
var wb = Excel.Open("merged.xlsx", new ExcelReadOptions { FillMergedCells = true });
// cells in the merged range other than the top-left now also have values
```

Output: (this example has no console output)

![Merged cells and hyperlinks](screenshots/merge_link.png)

*Produced by the example code in this chapter, opened in Excel: merged cells, external and internal links, a formula column.*
---

# 8. AutoFilter

## 📑 Contents

| # | Section |
| :-: | :--- |
| 8.1 | [Writing a Filter](#81-writing-a-filter) |
| 8.2 | [Filter Condition Type `FilterType`](#82-filter-condition-type-filtertype) |
| 8.3 | [Compare Operator `FilterOperator`](#83-compare-operator-filteroperator) |
| 8.4 | [Between Example](#84-between-example) |
| 8.5 | [Multiple Conditions (AND Logic)](#85-multiple-conditions-and-logic) |
| 8.6 | [Manually Specifying Hidden Rows](#86-manually-specifying-hidden-rows) |
| 8.7 | [Reading a Filter](#87-reading-a-filter) |

---

## 8.1 Writing a Filter

`Worksheet.Filter` is an `AutoFilter` object containing `Range`, per-column conditions `Columns`, and `HiddenRows`. The first row of an Excel filter range is **always the header** — so write the header row first, and data starts from row 2:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// Header row (row 1 of the filter range)
ws.SetValue("A1", "Name");
ws.SetValue("B1", "Type");
ws.SetValue("C1", "Score");

// Data starts from row 2, 541 rows total → range A1:C542
for (int r = 2; r <= 542; r++)
{
    ws.SetValue(r, 1, $"Name {r - 1}");
    ws.SetValue(r, 2, r % 2 == 0 ? "Active" : "Inactive");
    ws.SetValue(r, 3, (r - 1) * 10);
}

ws.Filter = new AutoFilter
{
    Range = "A1:C542",
    Columns = new List<FilterColumn>
    {
        new FilterColumn
        {
            ColumnIndex = 1,                 // 0-based column index (2nd column)
            Type = FilterType.Equals,
            Values = new List<string> { "Active" },
        },
    },
};
```

Key members of `AutoFilter`:

| Parameter | Type | Description |
|---|---|---|
| `Range` | `string` | Filter range in A1-style reference, first row is the header |
| `Columns` | `List<FilterColumn>` | Per-column filter conditions, 0-based column index |
| `HiddenRows` | `HashSet<int>` | Manually hidden 0-based row index set (optional, see 8.6) |

Key members of `FilterColumn`:

| Parameter | Type | Description |
|---|---|---|
| `ColumnIndex` | `int` | 0-based column index (1st column = 0) |
| `Type` | `FilterType` | Condition type, see 8.2 |
| `Values` | `List<string>` | Set of matching values |
| `Operator` | `FilterOperator` | Comparison operator when `Type = Compare`, see 8.3 |
| `MinValue` / `MaxValue` | `string?` | Lower / upper bound of `Between` |

Output: (this example has no console output)

> ⚠️ Do not write data into the first row of the filter range — Excel treats that row as the header (showing the filter arrow); data only participates in filtering correctly when it starts from row 2.

## 8.2 Filter Condition Type `FilterType`

`FilterColumn.Type` specifies the filter condition type for a column; the enum values are as follows:

```csharp
public enum FilterType { Equals, Compare, Contains, BeginsWith, EndsWith, Blank }
```

Output: (this example has no console output)

## 8.3 Compare Operator `FilterOperator`

When `FilterColumn.Type = FilterType.Compare`, use `Operator` to specify the comparison relationship:

```csharp
public enum FilterOperator { GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between }
```

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.GreaterThan,
    Values = new List<string> { "500" },
};
```

Output: (this example has no console output)

## 8.4 Between Example

For a range filter, use `Operator = FilterOperator.Between` together with `MinValue` / `MaxValue`:

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.Between,
    MinValue = "100",
    MaxValue = "500",
};
```

Output: (this example has no console output)

## 8.5 Multiple Conditions (AND Logic)

Multiple `FilterColumn` entries take effect simultaneously (a row must satisfy all column conditions):

```csharp
ws.Filter = new AutoFilter
{
    Range = "A1:D542",
    Columns = new List<FilterColumn>
    {
        new FilterColumn { ColumnIndex = 1, Type = FilterType.Equals, Values = new() { "Active" } },
        new FilterColumn { ColumnIndex = 2, Type = FilterType.Compare, Operator = FilterOperator.GreaterThan, Values = new() { "500" } },
    },
};
```

Output: (this example has no console output)

## 8.6 Manually Specifying Hidden Rows

`HiddenRows` is a 0-based row index set (relative to `Rows`):

```csharp
ws.Filter = new AutoFilter
{
    Range = "A1:D542",
    HiddenRows = new HashSet<int> { 1, 3, 5 },   // hides rows 2, 4, 6
};
```

Output: (this example has no console output)

## 8.7 Reading a Filter

After opening a file, read the filter range and per-column conditions via `Worksheet.Filter`:

```csharp
var opened = Excel.Open("filtered.xlsx");
var filter = opened.Worksheets[0].Filter;
if (filter is not null)
{
    Console.WriteLine(filter.Range);
    foreach (var col in filter.Columns)
        Console.WriteLine($"{col.ColumnIndex}: {col.Type} {string.Join(",", col.Values)}");
}
```

Output:

```
A1:C542
1: Equals Active
```
---

# 9. Row Height and Column Width

## 📑 Contents

| # | Section |
| :-: | :--- |
| 9.1 | [Setting Row Height](#91-setting-row-height) |
| 9.2 | [Setting Column Width](#92-setting-column-width) |
| 9.3 | [Auto-Fitting Column Width `AutoColumnWidths`](#93-auto-fitting-column-width-autocolumnwidths) |
| 9.4 | [Auto-Fitting on Write](#94-auto-fitting-on-write) |

---

## 9.1 Setting Row Height

`Worksheet.RowHeights` is a `Dictionary<int, double>` whose key is the **0-based row index**, in **points**:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Tall row");
ws.RowHeights = new Dictionary<int, double> { { 0, 30.0 } };   // row 1 height 30 points
```

Output: (this example has no console output)

## 9.2 Setting Column Width

`Worksheet.ColumnWidths` is a `Dictionary<int, double>` whose key is the **0-based column index**:

```csharp
ws.ColumnWidths = new Dictionary<int, double>
{
    { 0, 20.0 },
    { 1, 15.0 },
};
```

Output: (this example has no console output)

## 9.3 Auto-Fitting Column Width `AutoColumnWidths`

`Worksheet.AutoColumnWidths()` estimates the width of each column based on the existing content in the sheet (Chinese characters count as 2, English / digits count as 1, clamped to `[8, 50]`), and writes the results into `ColumnWidths`:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Name");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B1", "Description");
ws.SetValue("B2", "A very long description that should widen the column");
ws.AutoColumnWidths();
// ws.ColumnWidths is now estimated from the content
```

Read-back verification:

```csharp
var opened = Excel.Open("autowidth.xlsx");
var widths = opened.Worksheets[0].ColumnWidths;
if (widths is not null)
    foreach (var kv in widths)
        Console.WriteLine($"Col {kv.Key}: {kv.Value:F1}");
```

Output:

```
Col 0: 9.0
Col 1: 50.0
```

> ⚠️ `AutoColumnWidths` produces estimates (Chinese characters count as 2, English/digits count as 1, clamped to `[8, 50]`), which may differ slightly from Excel's actual rendered width.

## 9.4 Auto-Fitting on Write

Setting `ExcelWriteOptions.AutoFitColumns = true` on `Excel.Write` auto-estimates column widths for every sheet before writing:

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", "自动适配列宽");
Excel.Write("out.xlsx", wb, new ExcelWriteOptions { AutoFitColumns = true });
```

Output: written to out.xlsx

---
# 10. Comments

## 📑 Contents

| # | Section |
| :-: | :--- |
| 10.1 | [Writing Comments](#101-writing-comments) |
| 10.2 | [Reading Back Comments](#102-reading-back-comments) |
| 10.3 | [Object-Model API: Adding / Reading a Comment on a Specific Cell](#103-object-model-api-adding--reading-a-comment-on-a-specific-cell) |

---

## 10.1 Writing Comments

`Worksheet.Comments` is a `Dictionary<string, string>` whose key is an **A1-style cell reference** and whose value is the comment text:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string>
{
    { "A1", "This is a comment on A1" },
    { "B1", "Note for B1 <with special chars>" },
};
```

Output: (this example has no console output)

> ⚠️ Comments are supported only for xlsx / xlsm; when writing to xls / xlsb / csv they are dropped via the degradation mechanism (see Chapter 22). Comment write-back relies on the OOXML VML legacyDrawing, so verify with a real Excel open.

## 10.2 Reading Back Comments

```csharp
var opened = Excel.Open("comments.xlsx");
var comments = opened.Worksheets[0].Comments;
if (comments is not null && comments.TryGetValue("A1", out var text))
    Console.WriteLine(text);   // Output: This is a comment on A1
```

Output:

```
This is a comment on A1
```

## 10.3 Object-Model API: Adding / Reading a Comment on a Specific Cell

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("C5").SetValue("data");

// Add a comment
ws.Comments ??= new Dictionary<string, string>();
ws.Comments["C5"] = "审核通过";

// Read a comment
var opened = Excel.Open("comments2.xlsx");
string? note = null;
opened.Worksheets[0].Comments?.TryGetValue("C5", out note);
Console.WriteLine(note);   // Output: 审核通过
```

Output:

```
审核通过
```

![Comments and data validation](screenshots/comment_validation.png)

*Produced by the example code in this chapter, opened in Excel: a comment bubble and a data validation dropdown.*
---

# 11. Hyperlinks

## 📑 Contents

| # | Section |
| :-: | :--- |
| 11.1 | [Writing hyperlinks](#111-writing-hyperlinks) |
| 11.2 | [Hyperlink properties](#112-hyperlink-properties) |
| 11.3 | [Reading back hyperlinks](#113-reading-back-hyperlinks) |

---

## 11.1 Writing hyperlinks

`Cell.Hyperlink` is a `Hyperlink` object that supports external links (URL / file path) and in-workbook internal jumps:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// external link
ws.Cell("A1").SetValue("Example");
ws.Cell("A1").Hyperlink = new Hyperlink
{
    Target = "https://example.com",
    Tooltip = "Visit Example",
    IsInternal = false,
};

// internal jump (Target starts with '#')
ws.Cell("B1").SetValue("Go to Sheet2");
ws.Cell("B1").Hyperlink = new Hyperlink
{
    Target = "#Sheet2!A1",
    IsInternal = true,
};
```

Output: (this example has no console output)

## 11.2 Hyperlink properties

`Cell.Hyperlink` is a `Hyperlink` object with the following members:

- `Target`: link target. Internal links use the format `#SheetName!A1`; external ones are a full URL or file path.
- `Tooltip`: mouse-hover tooltip text (optional).
- `IsInternal`: whether it is an in-workbook internal jump.

Output: (this example has no console output)

## 11.3 Reading back hyperlinks

After opening a file, read hyperlink information via `Cell.Hyperlink`:

```csharp
var opened = Excel.Open("links.xlsx");
var cell = opened.Worksheets[0].Cell("A1");
if (cell.Hyperlink is { } h)
    Console.WriteLine($"{h.Target} internal={h.IsInternal} tooltip={h.Tooltip}");
```

Output:

```
https://example.com internal=False tooltip=Visit Example
```

Hyperlinks are supported for reading and writing in all four formats: xlsx / xlsm / xlsb / xls.

See the screenshot in [Chapter 7](#7-merged-cells).
---

# 12. Freeze Panes

## 📑 Contents

| # | Section |
| :-: | :--- |
| 12.1 | [Setting frozen rows / columns](#121-setting-frozen-rows--columns) |
| 12.2 | [FreezeHeader compatibility](#122-freezeheader-compatibility) |
| 12.3 | [Object-model API](#123-object-model-api) |
| 12.4 | [Reading back freeze](#124-reading-back-freeze) |

---

## 12.1 Setting frozen rows / columns

`Worksheet.FreezeRows` / `FreezeColumns` are the 1-based number of frozen rows / columns (0 = not frozen):

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.FreezeRows = 2;       // freeze the first 2 rows
ws.FreezeColumns = 3;    // freeze the first 3 columns
```

Output: (this example has no console output)

## 12.2 FreezeHeader compatibility

`FreezeHeader = true` is equivalent to `FreezeRows = 1`:

```csharp
ws.FreezeHeader = true;   // freeze the first row
```

Output: (this example has no console output)

## 12.3 Object-model API

Set frozen rows / columns directly through the properties (equivalent to 12.1):

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.FreezeRows = 1;
ws.FreezeColumns = 1;
```

Output: (this example has no console output)

## 12.4 Reading back freeze

After opening a file, read `FreezeRows` / `FreezeColumns`:

```csharp
var opened = Excel.Open("frozen.xlsx");
var ws = opened.Worksheets[0];
Console.WriteLine($"{ws.FreezeRows} rows, {ws.FreezeColumns} cols");
```

Output:

```
2 rows, 3 cols
```

Freeze panes are supported for freezing any rows / columns in all three formats: xlsx / xlsb / xls.

![Freeze panes](screenshots/image_freeze.png)

*Produced by the example code in this chapter, opened in Excel: the first two rows and the first column stay visible while scrolling.*
---

# 13. Images

## 📑 Contents

| # | Section |
| :-: | :--- |
| 13.1 | [Floating images](#131-floating-images) |
| 13.2 | [InCell images](#132-incell-images) |
| 13.3 | [Image placement enum `ImagePlacement`](#133-image-placement-enum-imageplacement) |
| 13.4 | [High-precision anchor `ImageAnchor` and move mode `ImageMoveMode`](#134-high-precision-anchor-imageanchor-and-move-mode-imagemovemode) |
| 13.5 | [Reading back images](#135-reading-back-images) |
| 13.6 | [Mixed use across multiple sheets](#136-mixed-use-across-multiple-sheets) |

---

Images are supported only in xlsx / xlsm. Use `Worksheet.AddImage` to add an image and `Worksheet.Images` to read them back.

## 13.1 Floating images

Anchored at the top-left corner of `row/column`, displayed at the image's original size by default:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
byte[] png = File.ReadAllBytes("logo.png");

ws.AddImage(png, 1, 1);                            // anchor A1, original size
ws.AddImage(png, 1, 3, 120, 60);                   // specify display size (pixels)
```

Important parameters of `AddImage` (row/column overload):

| Parameter | Type | Description |
|---|---|---|
| `data` | `byte[]` | Image binary (PNG/JPEG/GIF/BMP) |
| `row` / `column` | `int` | 1-based top-left anchor row / column |
| `widthPx` / `heightPx` | `double?` | Display size (pixels), null = original image size |
| `placement` | `ImagePlacement` | `Floating` / `InCell`, default `Floating` |

Output: (this example has no console output)

> ⚠️ Images are supported only in xlsx / xlsm (see the format support matrix in chapter 20).

## 13.2 InCell images

Excel 365 InCell images (richData system):

```csharp
ws.AddImage(png, 2, 1, placement: ImagePlacement.InCell);
```

Output: (this example has no console output)

> ⚠️ InCell images are based on the Excel 365 richData system (written back as a richData part); older versions of Excel may not recognize them.

## 13.3 Image placement enum `ImagePlacement`

`ImagePlacement` determines whether an image is embedded in a cell or floating:

```csharp
public enum ImagePlacement { InCell, Floating }
```

Output: (this example has no console output)

## 13.4 High-precision anchor `ImageAnchor` and move mode `ImageMoveMode`

`ImageAnchor` provides the top-left cell + EMU offsets + display size + move mode, and takes precedence over `Row`/`Column` when writing back:

```csharp
var anchor = new ImageAnchor
{
    TopLeftCell = "B2",
    TopLeftOffsetX = 100,        // EMU offset (1px ≈ 9525)
    TopLeftOffsetY = 50,
    WidthPixels = 200,
    HeightPixels = 100,
    MoveMode = ImageMoveMode.MoveAndSizeWithCells,
};
ws.AddImage(png, anchor, name: "Chart", altText: "季度图表");
```

```csharp
public enum ImageMoveMode
{
    MoveAndSizeWithCells,        // move and size with the cells
    MoveButDontSizeWithCells,    // move with the cells but do not size (default)
    FixedPosition,               // fixed position
}
```

Important parameters of `ImageAnchor`:

| Parameter | Type | Description |
|---|---|---|
| `TopLeftCell` | `string` | Top-left cell A1 reference |
| `TopLeftOffsetX` / `TopLeftOffsetY` | `int` | Top-left offset (EMU, 1px≈9525) |
| `WidthPixels` / `HeightPixels` | `double` | Display size (pixels) |
| `MoveMode` | `ImageMoveMode` | Move / size mode, default `MoveButDontSizeWithCells` |

Important parameters of `AddImage` (anchor overload):

| Parameter | Type | Description |
|---|---|---|
| `data` | `byte[]` | Image binary |
| `anchor` | `ImageAnchor` | High-precision anchor |
| `name` | `string?` | Image name (optional) |
| `altText` | `string?` | Accessibility alternative text (optional) |

Output: (this example has no console output)

> ⚠️ `ImageAnchor` applies only to Floating images; for InCell use the row/column overload (`Anchor` is ignored).

## 13.5 Reading back images

After opening a file that contains images, `Worksheet.Images` is populated automatically:

```csharp
var opened = Excel.Open("with_images.xlsx");
foreach (var img in opened.Worksheets[0].Images)
{
    Console.WriteLine($"{img.CellAddress} {img.Placement} {img.Data.Length} bytes");
    // img.Data is the original image bytes, img.Row/Column is the anchor, img.Extension is the extension
}
```

Output:

```
A1 Floating 70 bytes
C1 Floating 70 bytes
A2 InCell 70 bytes
```

Key members of `WorksheetImage`: `Data` (bytes), `Extension`, `Row`/`Column` (1-based anchor), `Placement`, `WidthPx`/`HeightPx`, `Name`, `Anchor`, `AltText`, `CellAddress` (read-only A1 reference).

## 13.6 Mixed use across multiple sheets

Different worksheets can use Floating and InCell placement independently without interfering with each other:

```csharp
byte[] img = File.ReadAllBytes("logo.png");
var wb = Excel.Create();

var wsBanner = wb.Worksheets[0];
wsBanner.Name = "Banner";
wsBanner.AddImage(img, 1, 1);                                 // floating image

var wsEmbed = wb.Worksheets.Add("Embed");
wsEmbed.AddImage(img, 1, 1, placement: ImagePlacement.InCell); // in-cell embedded image

wb.SaveAs("multi_images.xlsx");

// read-back verification
var opened = Excel.Open("multi_images.xlsx");
foreach (var s in opened.Worksheets)
    foreach (var im in s.Images)
        Console.WriteLine($"{s.Name}: {im.CellAddress} {im.Placement} {im.Data.Length} bytes");
```

Output:

```
Banner: A1 Floating 70 bytes
Embed: A1 InCell 70 bytes
```

> ⚠️ Images are supported only in xlsx / xlsm (see the format support matrix in chapter 20). InCell images may not be recognized by older versions of Excel.
See the screenshot in [Chapter 12](#12-freeze-panes).
---

# 14. Data Validation

## 📑 Contents

| # | Section |
| :-: | :--- |
| 14.1 | [Writing Data Validation](#141-writing-data-validation) |
| 14.2 | [Data Validation Type `DataValidationType`](#142-data-validation-type-datavalidationtype) |
| 14.3 | [Reading Back Data Validation](#143-reading-back-data-validation) |

---

## 14.1 Writing Data Validation

`Worksheet.Validations` is a `List<DataValidation>`, configured rule by rule via an object initializer; data validation is written only for xlsx / xlsm (other formats are discarded via degradation reporting):

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.Validations = new List<DataValidation>
{
    // Drop-down list (comma-separated, wrapped in quotes)
    new DataValidation
    {
        Type = DataValidationType.List,
        Sqref = "A1:A10",
        Formula1 = "\"Active,Inactive,Pending\"",
        AllowBlank = true,
        PromptTitle = "请选择",
        Prompt = "从下拉列表选择状态",
    },
    // Whole-number range validation
    new DataValidation
    {
        Type = DataValidationType.WholeNumber,
        Sqref = "B1:B10",
        Formula1 = "1",
        Formula2 = "100",
    },
};
```

| Parameter | Type | Description |
|---|---|---|
| `Type` | `DataValidationType` | Validation type (see 14.2) |
| `Sqref` | `string` | Applied range (A1 style, e.g. `A1:A10`) |
| `Formula1` | `string` | For list validation, a comma-separated list wrapped in quotes; for range validation, the lower bound |
| `Formula2` | `string?` | Upper bound for range validation (omittable for non-range types) |
| `AllowBlank` | `bool` | Whether empty values are allowed (default false) |
| `PromptTitle` / `Prompt` | `string?` | Input prompt title / body shown when the cell is selected |

Output: (this example has no console output)

## 14.2 Data Validation Type `DataValidationType`

`DataValidationType` determines the validation rule category, used together with `Formula1` / `Formula2`:

```csharp
public enum DataValidationType { List, WholeNumber, Decimal, Date }
```

- `List`: drop-down list validation; `Formula1` is a comma-separated list wrapped in quotes.
- `WholeNumber` / `Decimal` / `Date`: numeric / date validation; `Formula1` is the lower bound, `Formula2` the upper bound (range validation).

Output: (this example has no console output)

## 14.3 Reading Back Data Validation

After opening a file that contains data validation, `Worksheet.Validations` is populated automatically; iterate to print each rule:

```csharp
var opened = Excel.Open("validations.xlsx");
var validations = opened.Worksheets[0].Validations;
if (validations is not null)
    foreach (var v in validations)
        Console.WriteLine($"{v.Type} {v.Sqref} {v.Formula1} {v.Formula2}");
```

Output:

```
List A1:A10 Active,Inactive,Pending 
WholeNumber B1:B10 1 100
```

See the screenshot in [Chapter 10](#10-comments).
---

# 15. Conditional Formatting

## 📑 Contents

| # | Section |
| :-: | :--- |
| 15.1 | [Cell Value Comparison (cellIs)](#151-cell-value-comparison-cellis) |
| 15.2 | [Formula Condition (expression)](#152-formula-condition-expression) |
| 15.3 | [Color Scale (colorScale)](#153-color-scale-colorscale) |
| 15.4 | [Data Bar (dataBar)](#154-data-bar-databar) |
| 15.5 | [Long-tail Text / Blanks / Errors / Duplicates / Top N / Average Line](#155-long-tail-text--blanks--errors--duplicates--top-n--average-line) |
| 15.6 | [Icon Set (iconSet)](#156-icon-set-iconset) |
| 15.7 | [Reading Back Conditional Formatting](#157-reading-back-conditional-formatting) |

---

Conditional formatting is read/written in xlsx / xlsm. `Worksheet.ConditionalFormats` is a `List<ConditionalFormat>`.

## 15.1 Cell Value Comparison (cellIs)

`ConditionalFormatType.CellIs` compares against a fixed value using `ConditionalOperator`; when matched, `Style` is applied:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "B2:B10",
    Type = ConditionalFormatType.CellIs,
    Operator = ConditionalOperator.GreaterThan,
    Formula = "100",
    Style = new CellStyle { Bold = true, FillColor = "#FFC7CE" },
});
```

`ConditionalOperator`: `LessThan` / `LessThanOrEqual` / `Equal` / `NotEqual` / `GreaterThan` / `GreaterThanOrEqual` / `Between` / `NotBetween`. Between uses `Formula` (lower bound) + `Formula2` (upper bound).

`ConditionalFormat` common members:

| Parameter | Type | Description |
|---|---|---|
| `Sqref` | `string` | Applied range (A1 style, may contain multiple areas, e.g. `A1:A100 D2:D9`) |
| `Type` | `ConditionalFormatType` | Rule type (see the subsections of this chapter) |
| `Operator` | `ConditionalOperator` | Only valid for `CellIs` (default `GreaterThan`) |
| `Formula` / `Formula2` | `string?` | Comparison target / Between upper bound |
| `Style` | `CellStyle?` | Style applied when matched (font / fill / border; excludes alignment and number format) |

Output: (this example has no console output)

## 15.2 Formula Condition (expression)

`ConditionalFormatType.Expression` uses a formula returning TRUE / FALSE to decide; the formula uses relative references (relative to the current cell):

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "A1:A100",
    Type = ConditionalFormatType.Expression,
    Formula = "MOD(ROW(),2)=0",     // highlight even rows
    Style = new CellStyle { FillColor = "#D9E1F2" },
});
```

Output: (this example has no console output)

## 15.3 Color Scale (colorScale)

`ConditionalFormatType.ColorScale` blends between low / high colors based on the numeric value; when `MidColor` is non-null it becomes a 3-color scale:

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "C2:C10",
    Type = ConditionalFormatType.ColorScale,
    ColorScale = new ColorScaleInfo
    {
        LowColor = "F8696B",
        HighColor = "63BE7B",
        MidColor = "FFEB84",        // set nonNull to enable 3-color scale
    },
});
```

| Parameter | Type | Description |
|---|---|---|
| `LowColor` | `string` | Low-value color (`#RRGGBB` or `RRGGBB`) |
| `HighColor` | `string` | High-value color |
| `MidColor` | `string?` | Middle color; when non-null it is a 3-color scale, otherwise 2-color |

Output: (this example has no console output)

## 15.4 Data Bar (dataBar)

`ConditionalFormatType.DataBar` draws a bar proportional to the value inside the cell:

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "D2:D10",
    Type = ConditionalFormatType.DataBar,
    DataBar = new DataBarInfo
    {
        Color = "638EC6",
        ShowValue = true,
        MinLengthPercent = 0,
        MaxLengthPercent = 100,
    },
});
```

| Parameter | Type | Description |
|---|---|---|
| `Color` | `string` | Bar color (default Excel blue `638EC6`) |
| `ShowValue` | `bool` | Whether to also show the value (default true; false shows only the bar) |
| `MinLengthPercent` / `MaxLengthPercent` | `int` | Shortest / longest bar length percentage (0–100) |

Output: (this example has no console output)

## 15.5 Long-tail Text / Blanks / Errors / Duplicates / Top N / Average Line

The following types are the long-tail capabilities of conditional formatting; the matching rules and dedicated properties are described in the comments:

```csharp
// Contains the specified text
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "A2:A100",
    Type = ConditionalFormatType.ContainsText,
    Text = "urgent",
    Style = new CellStyle { Bold = true },
});

// Begins with / ends with / does not contain the specified text
// Type = BeginsWith / EndsWith / NotContainsText, also uses Text

// Text length comparison
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "B2:B100",
    Type = ConditionalFormatType.TextLength,
    Operator = ConditionalOperator.GreaterThan,
    Formula = "10",
});

// Time period
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "C2:C100",
    Type = ConditionalFormatType.TimePeriod,
    TimePeriod = "today",   // yesterday/today/tomorrow/lastWeek/thisWeek/nextWeek/lastMonth/thisMonth/nextMonth
});

// Blanks / non-blanks / errors / non-errors
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "D2:D100", Type = ConditionalFormatType.Blanks });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "D2:D100", Type = ConditionalFormatType.NoBlanks });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "E2:E100", Type = ConditionalFormatType.Errors });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "E2:E100", Type = ConditionalFormatType.NoErrors });

// Unique / duplicate values
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "F2:F100", Type = ConditionalFormatType.Unique });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "F2:F100", Type = ConditionalFormatType.Duplicate });

// Top N items / top N%
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "G2:G100",
    Type = ConditionalFormatType.Top10,
    Rank = 10,
    Percent = false,
});

// Above / below average
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "H2:H100", Type = ConditionalFormatType.AboveAverage });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "H2:H100", Type = ConditionalFormatType.BelowAverage });
```

Output: (this example has no console output)

## 15.6 Icon Set (iconSet)

`IconSetInfo` provides 17 built-in set enums + any custom set name + thresholds:

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "I2:I100",
    Type = ConditionalFormatType.IconSet,
    IconSet = new IconSetInfo
    {
        Style = IconSetStyle.ThreeArrows,       // default three colored arrows
        Percent = true,
        ShowValue = true,
        // When Thresholds is empty, percentages are divided evenly by the icon count
        // Thresholds = new double[] { 33, 66 },
    },
});
```

`IconSetStyle` enum (17 kinds): `ThreeArrows` / `ThreeArrowsGray` / `ThreeFlags` / `ThreeTrafficLights` / `ThreeTrafficLights2` / `ThreeSigns` / `ThreeSymbols` / `ThreeSymbols2` / `FourArrows` / `FourArrowsGray` / `FourRedToBlack` / `FourRating` / `FourTrafficLights` / `FiveArrows` / `FiveArrowsGray` / `FiveRating` / `FiveQuarters`. You can also use `CustomStyleName` to specify any set name string (takes precedence when non-empty).

| Parameter | Type | Description |
|---|---|---|
| `Style` | `IconSetStyle` | Built-in set (default `ThreeArrows`) |
| `CustomStyleName` | `string?` | Any set name string; takes precedence when non-empty |
| `Percent` | `bool` | Whether thresholds are percentages (true) or absolute values (false), default true |
| `ShowValue` | `bool` | Whether to also show the value in the cell, default true |
| `Thresholds` | `double[]?` | Custom thresholds (icon count - 1 values, ascending); when empty, divided evenly by icon count |

Output: (this example has no console output)

## 15.7 Reading Back Conditional Formatting

After opening a file that contains conditional formatting, `Worksheet.ConditionalFormats` is populated automatically; iterate to print each rule:

```csharp
var opened = Excel.Open("cf.xlsx");
var cfs = opened.Worksheets[0].ConditionalFormats;
foreach (var cf in cfs)
    Console.WriteLine($"{cf.Type} {cf.Sqref} {cf.Formula}");
```

Output:

```
CellIs B2:B10 100
Expression A1:A100 MOD(ROW(),2)=0
ColorScale C2:C10 
DataBar D2:D10 
IconSet I2:I100 
```

Other `ConditionalFormat` members: `Priority` (priority, auto-numbered by registration order by default), `Style` (style applied when the condition is met; excludes alignment and number format).

![Conditional formatting](screenshots/conditional.png)

*Produced by the example code in this chapter, opened in Excel: data bars, a color scale, icon sets, and top-N highlighting.*
---

# 16. Excel Tables

## 📑 Contents

| # | Section |
| :-: | :--- |
| 16.1 | [Creating an Excel Table](#161-creating-an-excel-table) |
| 16.2 | [Style Enum `TableStyleStyle`](#162-style-enum-tablestylestyle) |
| 16.3 | [Custom Style Name `CustomStyleName`](#163-custom-style-name-customstylename) |
| 16.4 | [Table Properties](#164-table-properties) |
| 16.5 | [Column Format (`XlTableColumn`)](#165-column-format-xltablecolumn) |
| 16.6 | [Removing an Excel Table](#166-removing-an-excel-table) |
| 16.7 | [Reading Back Excel Tables](#167-reading-back-excel-tables) |

---

Excel tables are read/written in xlsx / xlsm. `Worksheet.AddTable` creates them; `Worksheet.Tables` reads them back.

## 16.1 Creating an Excel Table

The first row of the covered range is used as the header column names; at least a header + 1 data row is required:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// First write the header + data
ws.SetValue("A1", "Product");
ws.SetValue("B1", "Price");
ws.SetValue("A2", "Apple");
ws.SetValue("B2", 3.5);
ws.SetValue("A3", "Banana");
ws.SetValue("B3", 2.0);

var table = ws.AddTable("A1:B3", "Products");     // default Medium9 style
```

| Parameter | Type | Description |
|---|---|---|
| `refAddress` | `string` | Table covered range (A1 style; the first row is always the header) |
| `name` | `string` | Table name (unique across the workbook; Chinese allowed, cannot start with a digit, cannot contain spaces, cannot collide with a cell address) |
| `style` | `TableStyleStyle?` | Built-in style enum; defaults to `Medium9` when omitted (there is also a `string styleName` overload, see 16.3) |

Returns an `XlTable`, on which you can directly set properties such as column format.

Output: (this example has no console output)

## 16.2 Style Enum `TableStyleStyle`

The `TableStyleStyle` enum includes 60 built-in stripe names (Light 1-21 / Medium 1-28 / Dark 1-11) + `None`; the style appearance is rendered by Excel itself, and only the style name is saved in the file:

```csharp
var table = ws.AddTable("A1:B3", "Products", TableStyleStyle.Medium2);
```

Output: (this example has no console output)

## 16.3 Custom Style Name `CustomStyleName`

The `string` overload `AddTable(ref, name, styleName)` accepts any style name string (including style names Excel may add in the future):

```csharp
var table = ws.AddTable("A1:B3", "Products", "TableStyleMedium9");
// When not among the 60 built-in names, Excel opens it degraded to no style (reported via OnDegradation)
```

Output: (this example has no console output)

> ⚠️ When the style name is not among the 60 built-in names, Excel silently degrades it to no style on open, reported via the `OnDegradation` callback (see Chapter 22).

## 16.4 Table Properties

`XlTable` members: `Name` (unique across the workbook, Chinese allowed, cannot start with a digit, cannot contain spaces, cannot collide with a cell address), `Ref`, `Style`, `CustomStyleName`, `ShowRowStripes` (default true), `ShowFirstColumn`, `ShowLastColumn`, `ShowColumnStripes`, `AutoFilter` (default true), `TotalsRowShown` (preserved on read-back), `HeaderStyle`, `Columns`. The returned `XlTable` lets you read these properties directly:

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Product");
ws.SetValue("B1", "Price");
ws.SetValue("A2", "Apple");
ws.SetValue("B2", 3.5);

var table = ws.AddTable("A1:B2", "Products", TableStyleStyle.Medium2);
Console.WriteLine($"{table.Name} {table.Ref} {table.Style} 行条纹={table.ShowRowStripes} 筛选={table.AutoFilter}");
```

Output:

```text
Products A1:B2 Medium2 行条纹=True 筛选=True
```

## 16.5 Column Format (`XlTableColumn`)

`table.Column(name)` gets a column by its name (case-insensitive); set `Style` (font/fill/border → dxf) and `NumberFormat`:

```csharp
var table = ws.AddTable("A1:B3", "Products");
table.Column("Price").NumberFormat = "#,##0.00";
table.Column("Price").Style = new CellStyle { Bold = true };
```

| Parameter | Type | Description |
|---|---|---|
| `name` of `Column(name)` | `string` | Column name (= header cell text, case-insensitive; throws `LiteExcelException` if it does not exist) |
| `NumberFormat` | `string?` | Number format for this column (e.g. `"#,##0.00"`) |
| `Style` | `CellStyle?` | Style for this column (font / fill / border; written out mapped to dxf) |

Output: (this example has no console output)

## 16.6 Removing an Excel Table

`RemoveTable(name)` removes a table by name (case-insensitive); returns `true` if it exists, otherwise `false`:

```csharp
bool removed = ws.RemoveTable("Products");   // removes it and returns true if it exists
```

Output: (this example has no console output)

## 16.7 Reading Back Excel Tables

After opening a file that contains Excel tables, `Worksheet.Tables` is populated automatically (including styles and column formats); iterate to print:

```csharp
var opened = Excel.Open("tables.xlsx");
foreach (var t in opened.Worksheets[0].Tables)
{
    Console.WriteLine($"{t.Name} {t.Ref} 样式={t.CustomStyleName ?? t.Style.ToString()}");
    foreach (var col in t.Columns)
        Console.WriteLine($"  {col.Name} fmt={col.NumberFormat}");
}
```

Output:

```
Products A1:B3 样式=TableStyleMedium2
  Product fmt=
  Price fmt=#,##0.00
```

![Excel table and filters](screenshots/table_filter.png)

*Produced by the example code in this chapter, opened in Excel: banded rows, header filter dropdowns, and currency format.*
---

# 17. Named Ranges

## 📑 Contents

| # | Section |
| :-: | :--- |
| 17.1 | [Reading Named Ranges](#171-reading-named-ranges) |
| 17.2 | [Preservation on Write](#172-preservation-on-write) |

---

> ⚠️ Named-range support: **xlsx / xlsm** full read-back (from `definedNames` in `workbook.xml`); **xls** supports simple cell/range references (PtgRef3d / PtgArea3d), names with complex formulas are skipped; **xlsb is not yet supported**. When writing to a format that does not support this capability, named ranges are **silently dropped**, reported via `OnDegradation`.

## 17.1 Reading Named Ranges

After opening a file containing named ranges, `Workbook.Names` is populated automatically (global + sheet-local); simply iterate to print them:

```csharp
var opened = Excel.Open("names.xlsx");
foreach (var nr in opened.Names)
    Console.WriteLine($"{nr.Name} = {nr.Reference} local={nr.LocalSheetId}");
```

`NamedRange` members: `Name`, `Reference` (e.g. `Sheet1!$A$1:$C$9`), `LocalSheetId` (-1 means a global name), `IsLocalSheet`.

Output:

```
MyRange = Sheet1!$A$1:$C$9 local=-1
LocalRange = Sheet1!$B$2 local=0
```

> xls files are also supported (simple cell/range reference read-back), e.g. `Excel.Open("names.xls")`. The `Workbook.Names` population logic is identical for xlsx/xlsm/xls.

## 17.2 Preservation on Write

Named ranges are **preserved as-is** when the file is saved after being opened (xlsx/xlsm pass through the `definedNames` in `workbook.xml`), so they are not lost through editing:

```csharp
var opened = Excel.Open("names.xlsx");
opened.Worksheets[0].SetValue("A1", "edited");
opened.Save();   // named ranges are still preserved
```

> ⚠️ Named ranges are **not written back** when saving to xls (xls write-back is not implemented); xlsx/xlsm saving passes them through.

Output: written to names.xlsx

---

# 18. File-Level Passwords

## 📑 Contents

| # | Section |
| :-: | :--- |
| 18.1 | [Opening an Encrypted File](#181-opening-an-encrypted-file) |
| 18.2 | [Reading the Security State](#182-reading-the-security-state) |
| 18.3 | [Setting Passwords](#183-setting-passwords) |
| 18.4 | [Removing Passwords](#184-removing-passwords) |
| 18.5 | [Modify Password Access and Read-Only](#185-modify-password-access-and-read-only) |
| 18.6 | [Fidelity Write-Back](#186-fidelity-write-back) |

---

File-level security is managed through `Workbook.Security` (`WorkbookSecurity`) and supports xlsx / xlsm / xlsb. The password payload itself is stored only inside the object and is never exposed in plain text.

## 18.1 Opening an Encrypted File

The open password (Agile encryption) is provided via `ExcelReadOptions.OpenPassword`; if the file is encrypted and no password is provided, a clear exception is thrown:

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions { OpenPassword = "secret" });
```

Output: (this example has no console output)

> ⚠️ Open / modify passwords are supported only for xlsx / xlsm / xlsb; when opening an encrypted file, reading throws an exception if the password is not provided (or is wrong).

## 18.2 Reading the Security State

`Workbook.Security` (`WorkbookSecurity`) exposes read-only security state properties; read them together with the open options:

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",
    ModifyPassword = "write",
});

var sec = wb.Security;
Console.WriteLine(sec.HasOpenPassword);      // true
Console.WriteLine(sec.HasModifyPassword);    // true
Console.WriteLine(sec.HasModifyAccess);      // true (a modify password was provided)
Console.WriteLine(sec.IsReadOnly);           // false
Console.WriteLine(sec.CanSave);              // true
```

| Property | Type | Description |
|---|---|---|
| `HasOpenPassword` | `bool` | whether the file has an open password (file encryption) |
| `HasModifyPassword` | `bool` | whether the file has a modify password (write protection) |
| `HasModifyAccess` | `bool` | whether modify access has been granted (the correct modify password was provided) |
| `IsReadOnly` | `bool` | read-only when the file has a modify password but access has not been granted |
| `CanSave` | `bool` | whether saving is allowed (`!IsReadOnly`) |

Output:

```
True
True
True
False
True
```

## 18.3 Setting Passwords

`Workbook.Security` provides methods for setting the open / modify passwords, which take effect on the next save:

```csharp
var wb = Excel.Create();
wb.Security.SetOpenPassword("secret");       // open password (file encryption)
wb.Security.SetModifyPassword("write");      // modify password (write protection), read-only recommended by default
wb.Security.SetModifyPassword("write", readOnlyRecommended: false);  // no read-only prompt
wb.SaveAs("secured.xlsx");
```

| Method | Parameter | Description |
|---|---|---|
| `SetOpenPassword` | `password` | sets the open password (file encryption), overwriting the old value; null / blank is treated as removal |
| `SetModifyPassword` | `password` | sets the modify password (write protection), overwriting the old value; null / blank is treated as removal |
| `SetModifyPassword` | `readOnlyRecommended` | whether to recommend opening read-only (default true) |
| `RemoveOpenPassword` | — | removes the open password (the next save has no open password) |
| `RemoveModifyPassword` | — | removes the modify password (requires modify access) |
| `ClearAll` | — | clears all file-level passwords (requires modify access) |

Output: written to secured.xlsx

> ⚠️ The password payload is stored only inside the `WorkbookSecurity` object and is never exposed in plain text; error messages and logs do not contain passwords.

## 18.4 Removing Passwords

After opening a password-protected file (providing the correct password to gain authorization), call the removal methods and save to strip the passwords:

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions { OpenPassword = "secret", ModifyPassword = "write" });
wb.Security.RemoveOpenPassword();            // the next save has no open password
wb.Security.RemoveModifyPassword();          // the next save has no modify protection
wb.Security.ClearAll();                      // clears all file-level passwords
wb.SaveAs("plain.xlsx");
```

Output: written to plain.xlsx

> ⚠️ `RemoveModifyPassword` / `ClearAll` require modify access to have been granted (otherwise a `LiteExcelException` is thrown), preventing unauthorized stripping / replacement of write protection.

## 18.5 Modify Password Access and Read-Only

A file that has a modify password but is opened without providing it opens in read-only mode; confirm via the security state properties:

```csharp
var wb = Excel.Open("readonly.xlsx");        // this file has a modify password but none was provided
Console.WriteLine(wb.Security.IsReadOnly);   // true
Console.WriteLine(wb.Security.CanSave);      // false
```

Output:

```
True
False
```

- When the file has a modify password but it is not provided (or is wrong), the workbook opens in **read-only** mode with `IsReadOnly = true`, `CanSave = false`; saving throws a `LiteExcelException`.
- Providing the correct `ModifyPassword` grants editing authorization (`HasModifyAccess = true`).
- `SetModifyPassword` / `RemoveModifyPassword` / `ClearAll` require modify access to have been granted, otherwise an exception is thrown (prevents unauthorized stripping / replacement of write protection).
- The original `fileSharing` captured at open time is passed through on save; it is regenerated when the user explicitly sets a new modify password.

## 18.6 Fidelity Write-Back

After opening an encrypted file, `SaveAs` inherits the password by default, with no need to set it again:

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",
    ModifyPassword = "write",
});
wb.SaveAs("secured_copy.xlsx");   // inherits the open password by default
```

Output: written to secured_copy.xlsx

> ⚠️ When `ModifyPasswordTouched` (the user explicitly changed the modify password) is set, the original fileSharing is not passed through; it is regenerated according to the newly set modify password. Saving a workbook that contains VBA macros as xlsx / xls throws an error (the format does not support macros).

---

# 19. Worksheet and Workbook Protection

## 📑 Contents

| # | Section |
| :-: | :--- |
| 19.1 | [Worksheet Protection `SheetProtection`](#191-worksheet-protection-sheetprotection) |
| 19.2 | [Workbook Protection `WorkbookProtection`](#192-workbook-protection-workbookprotection) |

---

## 19.1 Worksheet Protection `SheetProtection`

`Worksheet.Protection` controls which operations are allowed / disallowed on a protected worksheet, with an optional password (SHA-512 + salt hash):

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
var p = new SheetProtection
{
    Enabled = true,
    // Allowed operations (default false means disallowed):
    SelectLockedCells = true,
    SelectUnlockedCells = true,
    // FormatCells / FormatColumns / FormatRows / InsertColumns / InsertRows /
    // InsertHyperlinks / DeleteColumns / DeleteRows / Sort / AutoFilter / PivotTables
    Objects = true,                // default true: allows editing objects
    Scenarios = true,              // default true: allows editing scenarios
};
p.SetPassword("protect123");      // optional password (a method, cannot be placed in the initializer)
ws.Protection = p;
```

Read-back:

```csharp
var opened = Excel.Open("protected.xlsx");
var p = opened.Worksheets[0].Protection;
if (p is not null)
{
    Console.WriteLine(p.Enabled);
    Console.WriteLine(p.HasPassword);
    Console.WriteLine(p.VerifyPassword("protect123"));   // applies to the hash read from the file
    p.RemovePassword();              // removes the protection password (null/blank is also treated as removal)
}
```

`SheetProtection` parameters:

| Parameter | Type | Description |
|---|---|---|
| `Enabled` | `bool` | whether protection is enabled (prerequisite for writing `sheetProtection`) |
| `SelectLockedCells` / `SelectUnlockedCells` | `bool` | whether selecting locked / unlocked cells is allowed (default true) |
| `Objects` / `Scenarios` | `bool` | whether editing objects / scenarios is allowed (default true) |
| `FormatCells`…`PivotTables` | `bool` | allowed editing operations (default false, disallowed) |
| `SetPassword` / `RemovePassword` | method | sets / removes the protection password (null / blank treated as removal) |
| `VerifyPassword` | method | verifies a password (only applies to the hash read from the file) |

Output:

```
True
False
True
```

> ⚠️ On read-back `HasPassword` is always `False` — the password is stored on disk as a SHA-512 + salt hash, and the library never reads plain text back into memory. Whether a protection password is set should be determined via `VerifyPassword(...)`, not `HasPassword`.

## 19.2 Workbook Protection `WorkbookProtection`

`Workbook.Protection` locks the workbook structure / windows, with an optional password:

```csharp
var wb = Excel.Create();
var p2 = new WorkbookProtection
{
    Enabled = true,
    LockStructure = true,   // prevents inserting/deleting/moving/hiding/renaming worksheets
    LockWindows = false,
};
p2.SetPassword("wbpass");     // optional password (a method, cannot be placed in the initializer)
wb.Protection = p2;
// p2.RemovePassword() removes the workbook protection password
wb.SaveAs("wbprotected.xlsx");
```

Read-back:

```csharp
var opened = Excel.Open("wbprotected.xlsx");
var p = opened.Protection;
if (p is not null)
    Console.WriteLine($"{p.Enabled} structure={p.LockStructure} hasPwd={p.HasPassword}");
```

`WorkbookProtection` parameters:

| Parameter | Type | Description |
|---|---|---|
| `Enabled` | `bool` | whether protection is enabled (prerequisite for writing `workbookProtection`) |
| `LockStructure` | `bool` | prevents inserting / deleting / moving / hiding / renaming worksheets (default true) |
| `LockWindows` | `bool` | locks windows (default false) |
| `SetPassword` / `RemovePassword` | method | sets / removes the protection password (null / blank treated as removal) |
| `VerifyPassword` | method | verifies a password (only applies to the hash read from the file) |

Output:

```
True structure=True hasPwd=False
```

> ⚠️ Same as 19.1: the read-back `hasPwd=False` does not mean no password is set — the reason is above (plain text never enters memory; use `VerifyPassword` to determine).

---


Chapters 20–23: multi-format behavior and degradation, streaming and append, AOT compatibility — cross-format and cross-platform differences are centralized here.

# 20. Multi-Format Behavior

## 📑 Contents

| # | Section |
| :-: | :--- |
| 20.1 | [Format Capability Matrix](#201-format-capability-matrix) |
| 20.2 | [xls / xlsb read/write degradation](#202-xls--xlsb-readwrite-degradation) |
| 20.3 | [CSV Behavior](#203-csv-behavior) |
| 20.4 | [Encrypted File Format Restrictions](#204-encrypted-file-format-restrictions) |
| 20.5 | [Fidelity Round-Trip](#205-fidelity-round-trip) |

---

## 20.1 Format Capability Matrix

The table below lists the support status of each capability across formats. Capabilities not supported by xls / xlsb / csv are reported via `ExcelWriteOptions.OnDegradation` when writing (see Chapter 22).

| Capability | xlsx | xlsm | xlsb | xls | csv |
|---|---|---|---|---|---|
| Cell value / Header | ☑️ | ☑️ | ☑️ | ☑️ | text only |
| Style (font/color/border/alignment/wrap) | ☑️ | ☑️ | NumberFormat only | NumberFormat only | ❌ |
| Number format | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Merged cells | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Auto filter | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Row height / Column width | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Comments | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Data validation | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Hyperlinks | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Freeze panes | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Images (Floating / InCell) | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Conditional formatting | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Tables | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Named ranges | ☑️ | ☑️ | ❌ | read only | ❌ |
| Document properties | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| Open / Modify password | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| Formulas (write) | ☑️ | ☑️ | cached value | cached value | ❌ |
| Formulas (read) | ☑️ | ☑️ | restored when parseable | restored when parseable | ❌ |
| Charts / PivotTables | passthrough | passthrough | passthrough | ❌ | ❌ |
| Streaming read (StreamRows) | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Streaming write (XlsxStreamWriter) | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Append | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Progress callback (ReadWithProgress) | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| CSV separator (auto-detect read / explicit write) | n/a | n/a | n/a | n/a | ☑️ |
| Degradation reporting (OnDegradation) | n/a | n/a | ☑️ | ☑️ | ☑️ |
| Auto column width (AutoFitColumns / AutoColumnWidths) | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |

Write to csv with the degradation callback connected to observe discarded capabilities:

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };

Excel.Write("matrix.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
        Console.WriteLine($"[Degradation] {info.Capability} -> {info.TargetFormat} @ {info.SheetName}: {info.Message}"),
});
```

Output:

```
[Degradation] Comments -> Csv @ Sheet1: CSV does not support comments, comments on sheet 'Sheet1' have been discarded.
```

## 20.2 xls / xlsb read/write degradation

Writing to xls / xlsb: styles degrade to `NumberFormat` only (to avoid BIFF hand-writing risks); comments / data validation / conditional formatting / images / tables / named ranges are dropped; formula text is not kept and is written as the cached value. These degradations are reported via `OnDegradation` (see Chapter 22).

Reading xls / xlsb: only `NumberFormat` is retained from styles; advanced capabilities such as comments / data validation / conditional formatting / images / tables are not read back; parseable formulas are restored as A1 text into `Cell.Formula` (array formulas / 3D references / names fall back to the cached value only). **These degradations are explicitly reported via `OnDegradation` when writing** (see Chapter 22).

Reading an xls file (styles retain only the number format):

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", "hello");
Excel.Write("roundtrip.xls", wb);

var reopened = Excel.Open("roundtrip.xls");
var ws = reopened.Worksheets["Sheet1"];
Console.WriteLine($"{ws.Name}: {ws.Cell("A1").GetString()}");
```

Output:

```
Sheet1: hello
```

## 20.3 CSV Behavior

- CSV only supports single-sheet workbooks (writing multiple sheets throws `NotSupportedException`).
- When reading, the first row is not split into headers (`CsvBackend.Read` (see Appendix B.4) returns raw rows).
- Separator: auto-detected on read (comma > semicolon > tab, counting only outside quotes), `ExcelReadOptions.Separator` can be used to fix it; default separator for writing is comma, `ExcelWriteOptions.Separator` can be specified.
- CSV does not support styles / merged cells / comments / data validation / hyperlinks / images / conditional formatting / tables / named ranges / document properties / formulas / passwords.

Write then read back (first row is treated as data row, not split into headers):

```csharp
var sheet = new SheetData
{
    SheetName = "People",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
    },
};
Excel.Write("people.csv", sheet);

var wb = Excel.Open("people.csv");
var csv = wb.Worksheets[0].ToSheetData();
for (int r = 0; r < csv.Rows.Count; r++)
    Console.WriteLine($"{csv.Rows[r][0].GetString()}: {csv.Rows[r][1].GetString()}");
```

Output:

```
Name: Score
Alice: 95
Bob: 88
```

## 20.4 Encrypted File Format Restrictions

File-level passwords (open / modify) are only supported for xlsx / xlsm / xlsb. Saving to csv / xls with a password set throws `LiteExcelException`:

```csharp
var wb = Excel.Create();
wb.Security.SetOpenPassword("1");
try
{
    wb.SaveAs("enc.csv", ExcelFormat.Csv);
}
catch (LiteExcelException ex)
{
    Console.WriteLine(ex.Message);
}
```

Output:

```
Cannot write Csv: Csv format does not support file-level passwords (open password/modify password). Please use xlsx/xlsm/xlsb to save, or remove the password first.
```

## 20.5 Fidelity Round-Trip

When opening xlsx / xlsm / xlsb, unmapped OOXML parts (macros / themes / drawings / charts / pivot tables, etc.) are captured and transparently passed through when saving, avoiding silent deletion. Renaming a sheet no longer loses drawing associations; appending data no longer loses macros / charts.

```csharp
var wb = Excel.Open("macro.xlsm");   // open an xlsm containing macros
wb.Worksheets[0].Name = "Renamed Sheet";  // renaming sheet does not lose drawing association
wb.SaveAs("macro_copy.xlsm");
```

Output:

```
written to macro_copy.xlsm
```

---

# 21. Streaming Read / Progress Callback / Append

## 📑 Contents

| # | Section |
| :-: | :--- |
| 21.1 | [Streaming Read `StreamRows`](#211-streaming-read-streamrows) |
| 21.2 | [Reading with Progress `ReadWithProgress`](#212-reading-with-progress-readwithprogress) |
| 21.3 | [Append Data `Append`](#213-append-data-append) |
| 21.4 | [Streaming Writer `CreateWriter`](#214-streaming-writer-createwriter) |
| 21.5 | [Large Files and Memory Model](#215-large-files-and-memory-model) |
| 21.6 | [Pull-Based Enumeration `EnumerateRows`](#216-pull-based-enumeration-enumeraterows) | LINQ composable, early termination, no header skip |

---

## 21.1 Streaming Read `StreamRows`

Row-by-row callback without holding all data in memory, suitable for large files. **xlsx / xlsm only**:

```csharp
var sheet = new SheetData
{
    SheetName = "Scores",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
    },
};
Excel.Write("scores.xlsx", sheet);

Excel.StreamRows("scores.xlsx", "Scores", row =>
{
    foreach (var cell in row)
        Console.Write($"{cell.GetString()} | ");
    Console.WriteLine();
});
```

Output:

```
Alice | 95 |
Bob | 88 |
```

## 21.2 Reading with Progress `ReadWithProgress`

First quickly scans the total number of data rows, then streams row-by-row. `current` increments from 1 to `total` (number of data rows, excluding headers):

```csharp
var sheet = new SheetData
{
    SheetName = "Scores",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
        new Cell[] { Cell.FromText("Cara"), Cell.FromNumber(72) },
    },
};
Excel.Write("progress.xlsx", sheet);

Excel.ReadWithProgress("progress.xlsx", 0, (current, total) =>
    Console.WriteLine($"{current}/{total}"));
```

Output:

```
1/3
2/3
3/3
```

## 21.3 Append Data `Append`

`Excel.Append(path, SheetData, WorkbookProperties?)` (`SheetData` see Appendix B.1): for an existing sheet with the same name, merges headers then appends rows; for a different name, adds as a new sheet; creates the file if it does not exist. **xlsx / xlsm only**:

```csharp
// write 3 rows first
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
Excel.Write("append.xlsx", sheet1);

// append 2 rows
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
Excel.Append("append.xlsx", appendData);

// read back to verify
var read = Excel.ReadAsDataTable("append.xlsx");
Console.WriteLine(read.Rows.Count);
```

Output:

```
5
```

Appending does not change the existing sheet order; worksheet-level preserved rels remain reusable (xlsm append does not lose macros).

**Header alignment on append**: when appending to a sheet with the same name, headers from the new data that do not exist in the original headers are **appended to the end of the original headers**; data rows are aligned by column name to the merged column positions; missing columns are filled with `Empty`:

```csharp
// original file headers: ID | Name
Excel.Write("align.xlsx", new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID", "Name" },
    Rows = new() { new Cell[] { Cell.FromNumber(1), Cell.FromText("Alice") } },
});

// appended headers include new column Price → merged to ID | Name | Price; data rows aligned by column name
Excel.Append("align.xlsx", new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID", "Price" },
    Rows = new() { new Cell[] { Cell.FromNumber(2), Cell.FromNumber(9.5) } },
});

var sheet = XlsxReader.Read("align.xlsx", 0);   // low-level read back (see Appendix B.2)
Console.WriteLine("headers: " + string.Join(" | ", sheet.Headers));
var ws2 = Excel.Open("align.xlsx").Worksheets["Data"];
Console.WriteLine($"C3 = {ws2.Cell("C3").Number}");   // Price aligned to column 3
Console.WriteLine($"B3 type = {ws2.Cell("B3").Type}"); // missing column filled with Empty
```

Output:

```
headers: ID | Name | Price
C3 = 9.5
B3 type = Empty
```

## 21.4 Streaming Writer `CreateWriter`

`Excel.CreateWriter` returns an `XlsxStreamWriter` (see Appendix B.5), writing large files row by row without holding all data in memory. **Only supports .xlsx / .xlsm extensions**; must `Dispose` / `Close` to finalize the file:

```csharp
using var writer = Excel.CreateWriter("big_out.xlsx");
for (int i = 0; i < 1_000_000; i++)
    writer.WriteRow(new object?[] { i, $"row {i}", i * 1.5, i % 2 == 0 });
// using block ends, auto Close
```

Output:

```
written to big_out.xlsx
```

Can also write to a stream (`LeaveOpen` is managed by the caller):

```csharp
using var ms = new MemoryStream();
using (var writer = Excel.CreateWriter(ms))
    writer.WriteRow(new object?[] { 1, "a" });
ms.Position = 0;
var read = XlsxReader.Read(ms, 0);
```

`XlsxStreamWriter` supports styles / formulas / hyperlinks per row (styles.xml and sheet rels are written uniformly on Close); advanced capabilities like merge / filter / images are not supported. When hyperlinks are extremely numerous, memory is no longer constant (all hyperlink references are buffered internally).

**Per-sheet row limit**: each worksheet holds at most 1,048,576 rows. The behavior at the limit is controlled by `RowLimitExceededMode`:

- `Throw` (default): throws `RowLimitExceededException` (inherits `LiteExcelException`); the file stays valid
- `SpillToNewSheet`: automatically starts a new worksheet (`Sheet1` / `Sheet2` / ...) to continue
- `Truncate`: stops writing further rows; check `writer.Truncated` to see if truncation occurred

```csharp
// Default: throw
using var w1 = Excel.CreateWriter("a.xlsx");

// Auto-spill
using var w2 = Excel.CreateWriter("b.xlsx", RowLimitExceededMode.SpillToNewSheet);

// Truncate
using var w3 = Excel.CreateWriter("c.xlsx", RowLimitExceededMode.Truncate);
for (int i = 0; i < 3_000_000; i++) w3.WriteRow(new object?[] { i });
// w3.Truncated == true (rows beyond 1,048,576 are dropped)
```

When spilling, you can provide a header that is written as the first row of every sheet:

```csharp
using var writer = Excel.CreateWriter("big_out.xlsx", RowLimitExceededMode.SpillToNewSheet,
    spillHeader: new object?[] { "ID", "Name", "Value" });
for (int i = 0; i < 3_000_000; i++)
    writer.WriteRow(new object?[] { i, "row" + i, i * 1.5 });
// 3 sheets, each starting with ID/Name/Value; caller writes data rows only
```

`spillHeader` only takes effect in `SpillToNewSheet` mode; ignored otherwise.

---

## 21.5 Large Files and Memory Model

For large files, prefer the streaming entry points over loading everything into memory:

```csharp
using LiteExcel;

using (var writer = Excel.CreateWriter("big.xlsx"))
{
    writer.WriteRow(new[] { Cell.FromText("No."), Cell.FromText("Value") });
    for (int i = 1; i <= 100000; i++)
        writer.WriteRow(new[] { Cell.FromNumber(i), Cell.FromNumber(i * 1.5) });
}

long rows = 0;
Excel.StreamRows("big.xlsx", "Sheet1", row => rows++);
Console.WriteLine($"rows: {rows}");

Excel.ReadWithProgress("big.xlsx", 0, (current, total) =>
    Console.WriteLine($"progress {current}/{total}"));
```

- **In-memory model**: the `Workbook` returned by `Excel.Open` / `Excel.Create` is an in-memory model; the entire workbook is loaded into memory. For very large files use the streaming APIs instead of `Excel.Open`.
- **Streaming scope**: `Excel.CreateWriter` / `Excel.StreamRows` / `Excel.Append` support xlsx / xlsm only (see 21.1).
- **Hyperlink count**: when the number of hyperlinks is extremely large, the streaming writer's memory is no longer constant (all hyperlink references are buffered internally).
- **Append**: `Excel.Append` reads the entire existing file before writing; suited to incremental appends of small/medium files.

## 21.6 Pull-Based Enumeration EnumerateRows

`Excel.EnumerateRows` returns `IEnumerable<IReadOnlyList<Cell>>`, yielding one row at a time with deferred execution, LINQ support, and early termination, without holding the whole sheet in memory.

Comparison with `StreamRows` (21.1):

| | `StreamRows` | `EnumerateRows` |
|---|---|---|
| Model | Push (`Action` callback) | Pull (`IEnumerable`) |
| First row | Skipped | Not skipped; all rows returned |
| Early termination | No | Yes (`break` / `First()` / `Take(n)`) |
| LINQ | No | Yes |

```csharp
// First row only (stops after the first row, does not scan the whole sheet)
var first = Excel.EnumerateRows("big.xlsx", "Sheet1").First();

// First 100 rows
foreach (var row in Excel.EnumerateRows("big.xlsx", "Sheet1").Take(100))
    Console.WriteLine(row[0].GetString());

// Skip the header row
foreach (var row in Excel.EnumerateRows("big.xlsx", "Sheet1").Skip(1))
    Process(row);

// No sheet name -> first sheet
foreach (var row in Excel.EnumerateRows("big.xlsx"))
    Process(row);
```

`sheetName` null defaults to the first sheet. Only xlsx / xlsm are supported. The iterator releases the file handle on dispose; `break` also releases it.

```csharp
// WinForms async call to avoid UI thread blocking
var names = await Task.Run(() => Excel.GetSheetNames(sPath));
await Task.Run(() =>
{
    foreach (var row in Excel.EnumerateRows(sPath, names[0]))
        Handle(row);
});
```

# 22. Degradation Callback OnDegradation

## 📑 Contents

| # | Section |
| :-: | :--- |
| 22.1 | [Capability Enum `DegradationCapability`](#221-capability-enum-degradationcapability) |
| 22.2 | [Degradation Info `DegradationInfo`](#222-degradation-info-degradationinfo) |
| 22.3 | [Style Degradation Details](#223-style-degradation-details) |

---

`ExcelWriteOptions.OnDegradation` is an optional callback (default null; when not registered, behavior is identical to previous versions — no breaking change). When writing to a format that does not support a certain capability (xls / xlsb / csv), each capability that is silently discarded is reported via callback:

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };   // csv does not support comments

Excel.Write("out.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
    {
        Console.WriteLine($"[Degradation] {info.Capability} -> {info.TargetFormat} @ {info.SheetName}: {info.Message}");
    },
});
```

Output:

```
[Degradation] Comments -> Csv @ Sheet1: CSV does not support comments, comments on sheet 'Sheet1' have been discarded.
```

## 22.1 Capability Enum `DegradationCapability`

The `DegradationCapability` enum lists all capabilities that can be reported for degradation:

```csharp
foreach (var cap in Enum.GetNames(typeof(DegradationCapability)))
    Console.WriteLine(cap);
```

Output:

```
Comments
DataValidation
AutoFilter
Images
DocumentProperties
NamedRanges
Styles
MergedCells
FreezePanes
Hyperlinks
RowHeights
ColumnWidths
Formulas
Charts
PivotTables
RichData
ConditionalFormatting
Tables
```

## 22.2 Degradation Info `DegradationInfo`

`DegradationInfo` carries the full description of a single degradation event: `Capability` (discarded capability), `SheetName` (null for workbook-level capabilities), `TargetFormat` (target format), `Message` (human-readable explanation).

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };

Excel.Write("deg.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
    {
        Console.WriteLine($"Capability={info.Capability}");
        Console.WriteLine($"SheetName={info.SheetName}");
        Console.WriteLine($"TargetFormat={info.TargetFormat}");
        Console.WriteLine($"Message={info.Message}");
    },
});
```

Output:

```
Capability=Comments
SheetName=Sheet1
TargetFormat=Csv
Message=CSV does not support comments, comments on sheet 'Sheet1' have been discarded.
```

## 22.3 Style Degradation Details

When writing to xls / xlsb, full styles (font / color / border / alignment / wrap) are degraded to retain only NumberFormat, reported via `DegradationCapability.Styles`:

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", 1);
ws.Cell("A1").Style = new CellStyle { Bold = true, FillColor = "FFFF00" };

Excel.Write("style.xls", wb, new ExcelWriteOptions
{
    OnDegradation = info => Console.WriteLine($"{info.Capability}: {info.Message}"),
});
```

Output:

```
Styles: xls only supports number format, full styles (font/color/border/alignment/wrap) on sheet 'Sheet1' have been degraded.
```

---

# 23. AOT Compatibility

## 📑 Contents

| # | Section |
| :-: | :--- |
| 23.1 | [DAM Annotations](#231-dam-annotations) |
| 23.2 | [IsAotCompatible](#232-isaotcompatible) |
| 23.3 | [Verification and Results Summary](#233-verification-and-results-summary) |
| 23.4 | [InvariantGlobalization](#234-invariantglobalization) |

---

## 23.1 DAM Annotations

List\<T\> reflection-mapping APIs are annotated with `[DynamicallyAccessedMembers]`, safe for AOT / trimming:

- `Excel.Create<T>` / `Excel.Write<T>` / `Excel.Read<T>` annotated with `PublicProperties` (reading also includes `PublicParameterlessConstructor`).
- `Worksheet.ImportData<T>` / `WorksheetCollection.Add<T>` annotated with `PublicProperties`.
- `XlsxReader.Read<T>` / `XlsxWriter.Write<T>` also annotated (see Appendix B.2 / B.3).

Using the `Person` class from Chapter 5 as an example (mapping requires no extra configuration, public properties preserved by library annotations):

```csharp
var people = new List<Person> { new() { Name = "Zhang San", Age = 30 } };
Excel.Write("people.xlsx", people);

var read = Excel.Read<Person>("people.xlsx");
Console.WriteLine(read.Count);
```

Output:

```
1
```

## 23.2 IsAotCompatible

The net8.0 target declares `IsAotCompatible=true` in the csproj, all public APIs are compatible with Native AOT / trimming:

```csharp
var wb = Excel.Create("Sheet1");
wb.Worksheets["Sheet1"].SetValue("A1", "x");
Excel.Write("aot.xlsx", wb);

var reopened = Excel.Open("aot.xlsx");
Console.WriteLine(reopened.Worksheets.Count);
```

Output:

```
1
```

## 23.3 Verification and Results Summary

- Verified via native AOT executable, all public APIs pass.
- AOT zero IL warnings + runtime assertions pass.
- Note: `Excel.Read<T>` / `XlsxReader.Read` (see Appendix B.2) only support xlsx / xlsm; **xls / xlsb / csv must use `Excel.Open(path)`** (routes by extension to the backend). Use `Excel.Open` for non-zip formats when reading sheet names on demand.

Reading entry for non-zip formats:

```csharp
var wb = Excel.Open("data.xls");
Console.WriteLine(string.Join(",", wb.Worksheets.Names));
```

Output:

```
Sheet1
```

## 23.4 InvariantGlobalization

Global invariants (common in AOT / containers):

- Setting `<InvariantGlobalization>true</InvariantGlobalization>`` at publish time does not affect any functionality of this library; both `Encoding.GetEncoding` on the read side and `CultureInfo.InvariantCulture` on the write side have been verified.
- **Prerequisite**: the base date boundary (1900/1904 date system) and xls ANSI strings require `Latin1` — characters not in the current system code page may be distorted, which is an inherent limitation of BIFF8 and unrelated to AOT.

Date write / read-back under invariant mode (fixed `yyyy-MM-dd` number format):

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", new DateTime(2024, 1, 2));
Excel.Write("inv.csv", wb);

var reopened = Excel.Open("inv.csv");
var data = reopened.Worksheets[0].ToSheetData();
Console.WriteLine(data.Rows[0][0].GetString());
```

Output:

```
2024-01-02
```

---


Chapters 24–25 cover exception handling and large-file considerations; it is recommended to read them through before going live.

# 24. Exception Handling

## 📑 Contents

| # | Section |
| :-: | :--- |
| 24.1 | [Exception Hierarchy](#241-exception-hierarchy) |
| 24.2 | [Common Exception Scenarios](#242-common-exception-scenarios) |
| 24.3 | [Recommendations](#243-recommendations) |

---

## 24.1 Exception Hierarchy

- `LiteExcelException`: the base class of all library exceptions.
- `LiteXlsxException`: a compatibility alias for the old exception name (`[Obsolete]`, use `LiteExcelException` instead).
- `InvalidSheetNameException`: thrown when a sheet name is invalid (empty, longer than 31 characters, or containing illegal characters); it carries a `SheetName` property.

An invalid sheet name is thrown at write time (`SheetName` carries the original name):

```csharp
var sheet = new SheetData { SheetName = "非法?名称" };
try
{
    Excel.Write("bad.xlsx", sheet);
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine($"非法 Sheet 名：{ex.SheetName}");
}
```

Output:

```
非法 Sheet 名：非法?名称
```

## 24.2 Common Exception Scenarios

| Scenario | Exception |
|---|---|
| File not found | `FileNotFoundException` |
| Empty path | `ArgumentException` |
| Duplicate / missing worksheet name | `LiteExcelException` |
| Invalid sheet name | `InvalidSheetNameException` |
| Save path extension does not match format | `LiteExcelException` |
| Saving a new workbook without a target path | `LiteExcelException` |
| Saving a read-only workbook (modify password not authorized) | `LiteExcelException` |
| Saving with a password as csv / xls | `LiteExcelException` |
| Saving a workbook containing macros as xlsx / xls | `LiteExcelException` |
| Streaming read / append of a non-xlsx/xlsm file | `LiteExcelException` |
| CSV multi-sheet write | `NotSupportedException` |
| Cell type mismatch on strongly-typed read | `InvalidCastException` |

Typical catch order (specific exceptions before the base class):

```csharp
var wb = Excel.Create();
try
{
    wb.Save();   // saving a new workbook without a target path
}
catch (LiteExcelException ex)
{
    Console.WriteLine(ex.Message);
}
```

Output:

```
当前工作簿没有目标路径，请使用 SaveAs 指定保存位置
```

## 24.3 Recommendations

```csharp
try
{
    var wb = Excel.Open("report.xlsx", new ExcelReadOptions { OpenPassword = "secret" });
    wb.SaveAs("out.xlsx");
}
catch (LiteExcelException ex)
{
    Console.WriteLine($"LiteExcel: {ex.Message}");
}
catch (Exception ex)
{
    Console.WriteLine($"Other: {ex.Message}");
}
```

Output:

```
written to out.xlsx
```


# Appendix A Object Model Quick Reference

## 📑 Contents

| # | Section |
| :-: | :--- |
| A.1 | [`Excel` static class](#a1-excel-static-class) |
| A.2 | [`Workbook`](#a2-workbook) |
| A.3 | [`Worksheet`](#a3-worksheet) |
| A.4 | [`Cells`](#a4-cells) |
| A.5 | [`ExcelRange`](#a5-excelrange) |
| A.6 | [`Cell`](#a6-cell) |
| A.7 | [Model Classes](#a7-model-classes) |

---

## A.1 `Excel` static class

| Member | Description |
|---|---|
| `Open(path[, options])` | Opens a file, auto-detecting the format by extension |
| `Open(path, format[, options])` | Opens a file with the specified format |
| `Open(stream, format[, options])` | Opens from a stream (the format must be specified) |
| `Create()` / `Create(sheetName)` / `Create(string[])` / `Create(format)` | Creates a new empty workbook |
| `Create<T>(data[, sheetName, format, configure])` | Creates a workbook and writes a List\<T\> |
| `Create(DataTable[, sheetName, format])` | Creates a workbook and writes a DataTable |
| `Write(path, Workbook[, options])` | Writes out a workbook |
| `Write(path, SheetData[, options])` | Writes a single sheet (low-level) |
| `Write(path, DataTable[, sheetName, options])` | Writes out a DataTable |
| `Write<T>(path, data[, sheetName, configure])` | Writes out a List\<T\> |
| `Read<T>(path[, sheetName, configure])` | Reads into a List\<T\> |
| `ReadAsDataTable(path[, sheetName, firstRowIsHeader])` | Reads into a DataTable |
| `GetSheetNames(path)` / `GetSheetNames(stream)` | Lists worksheet names |
| `StreamRows(path, sheetName, onRow)` | Streams rows one by one |
| `CreateWriter(path)` / `CreateWriter(stream)` | Creates a streaming writer |
| `Append(path, SheetData[, properties])` | Appends data |
| `ReadWithProgress(path, sheetIndex, onProgress)` | Reads with progress reporting |
| `DetectFormat(path)` | Detects the format by extension |

## A.2 `Workbook`

| Member | Description |
|---|---|
| `Worksheets` | Worksheet collection (`WorksheetCollection`) |
| `Properties` | Document properties (`WorkbookProperties`) |
| `Format` | Current format (`ExcelFormat`) |
| `Security` | File-level security (`WorkbookSecurity`) |
| `Protection` | Workbook protection (`WorkbookProtection`) |
| `Names` | Named ranges (`List<NamedRange>`) |
| `CurrentPath` | Current target path |
| `Save()` / `SaveAs(path[, format])` / `Save(stream, format)` | Saves |

## A.3 `Worksheet`

| Member | Description |
|---|---|
| `Name` | Worksheet name |
| `Cell(row, col)` / `Cell(address)` | Accesses a cell |
| `Range(address)` / `Range(r1, c1, r2, c2)` | Accesses a range |
| `Cells` | Cell collection of the whole sheet |
| `SetValue(row, col, value)` / `SetValue(address, value)` | Sets a value |
| `Merge` / `Unmerge` / `MergedRanges` | Merge |
| `RowHeights` / `ColumnWidths` | Row heights / column widths |
| `AutoColumnWidths()` | Auto-fits column widths |
| `HeaderStyle` / `DefaultStyle` / `RowStyles` / `ColumnStyles` | Styles |
| `Comments` | Comments |
| `Validations` | Data validation |
| `Filter` | Auto-filter |
| `ConditionalFormats` | Conditional formatting |
| `Images` / `AddImage(...)` | Images |
| `Protection` | Worksheet protection |
| `Tables` / `AddTable` / `RemoveTable` | Tables |
| `FreezeRows` / `FreezeColumns` / `FreezeHeader` | Frozen panes |
| `ImportData<T>(data[, configure])` / `ImportData(DataTable[, includeHeader])` | Clears and rebuilds via import |
| `ToSheetData()` | Exports to the low-level SheetData model |
| `RowCount` / `MaxColumn` | Size information |

## A.4 `Cells`

| Member | Description |
|---|---|
| `this[int row, int column]` | Indexer by row and column |
| `this[string address]` | Indexer by A1 address |
| `Range(address)` / `Range(r1, c1, r2, c2)` | Extracts a range |
| `SetValue(...)` | Convenient value writing |
| `Clear()` | Clears all sheet values |
| `GetEnumerator()` | Enumerates existing cells |

## A.5 `ExcelRange`

| Member | Description |
|---|---|
| `FirstRow` / `FirstCol` / `LastRow` / `LastCol` | Range boundaries (1-based) |
| `Address` | A1 address |
| `RowCount` / `ColumnCount` | Size |
| `Cell(rowOffset, colOffset)` | Relative offset within the range |
| `Fill(value)` / `Fill(object?[,])` | Bulk write |
| `ToValues()` / `ToCells()` | Read back |
| `Style` | Uniform style for the whole range |
| `Merge()` / `Unmerge()` | Merge |
| `Clear()` | Clears |
| `GetEnumerator()` | Enumerates (row-major) |

## A.6 `Cell`

| Member | Description |
|---|---|
| `Type` / `Text` / `Number` / `Date` / `Boolean` | Value fields |
| `Style` / `NumberFormat` | Style |
| `Formula` / `IsFormula` | Formula |
| `Hyperlink` | Hyperlink |
| `IsEmpty` | Whether empty |
| `FromText` / `FromNumber` / `FromDate` / `FromBoolean` / `FromFormula` / `Empty` | Factory methods |
| `SetValue(object?)` | Sets a value |
| `GetString` / `GetDouble` / `GetDateTime` / `GetBoolean` / `GetValue` | Strongly-typed reads |
| `TryGetString` / `TryGetDouble` / `TryGetDateTime` / `TryGetBoolean` | Try reads |

## A.7 Model Classes

- `CellStyle` / `BorderStyle` / `BorderEdge` / `HorizontalAlignment` / `VerticalAlignment`
- `CellRange` (0-based)
- `AutoFilter` / `FilterColumn` / `FilterType` / `FilterOperator`
- `DataValidation` / `DataValidationType`
- `ConditionalFormat` / `ConditionalFormatType` / `ConditionalOperator` / `ColorScaleInfo` / `DataBarInfo` / `IconSetInfo` / `IconSetStyle`
- `Hyperlink`
- `NamedRange`
- `XlTable` / `XlTableColumn` / `TableStyleStyle`
- `WorksheetImage` / `ImageAnchor` / `ImagePlacement` / `ImageMoveMode`
- `SheetProtection` / `WorkbookProtection`
- `WorkbookProperties` / `WorkbookSecurity`
- `DegradationInfo` / `DegradationCapability`
- `ExcelFormat` / `ExcelReadOptions` / `ExcelWriteOptions` / `WriteOptions<T>` / `ReadOptions<T>` / `LiteColumnAttribute`
- `CellRef` (A1 reference utility, static class)
- `XlsxStreamWriter`
- `LiteExcelException` / `LiteXlsxException` / `InvalidSheetNameException`

---

# Appendix B Low-Level API Reference

## 📑 Contents

| # | Section |
| :-: | :--- |
| B.1 | [`SheetData` (complete data of one worksheet)](#b1-sheetdata-complete-data-of-one-worksheet) |
| B.2 | [`XlsxReader` (static class, zero reflection, AOT-safe)](#b2-xlsxreader-static-class-zero-reflection-aot-safe) |
| B.3 | [`XlsxWriter` (static class, zero reflection, AOT-safe)](#b3-xlsxwriter-static-class-zero-reflection-aot-safe) |
| B.4 | [`CsvBackend` (internal class, CSV format backend)](#b4-csvbackend-internal-class-csv-format-backend) |
| B.5 | [`XlsxStreamWriter` (streaming writer)](#b5-xlsxstreamwriter-streaming-writer) |
| B.6 | [`CellRef` (A1 reference utility, static class)](#b6-cellref-a1-reference-utility-static-class) |
| B.7 | [Old vs. New API Mapping](#b7-old-vs-new-api-mapping) |

---

> **Use cases**: suited to custom / bare-row-data / large-file scenarios. For everyday use, prefer the object model API (Chapters 2–25). Low-level API coordinate convention: the `Cell` in `SheetData.Rows` is a 0-based grid, and `Headers` holds the first-row header text.

## B.1 `SheetData` (complete data of one worksheet)

↳ Main guide: Chapter 5 Data Types and Conversion (the underlying mapping of List<T> / DataTable is SheetData), Chapter 21 Streaming Read / Progress Callback / Append (the data carrier for streaming / append)

```csharp
public sealed class SheetData
{
    public string SheetName { get; set; } = "Sheet1";
    public List<string> Headers { get; set; } = new();
    public List<IReadOnlyList<Cell>> Rows { get; set; } = new();
    public List<CellRange> MergedRanges { get; set; } = new();
    public AutoFilter? Filter { get; set; }
    public bool FreezeHeader { get; set; }
    public int FreezeRows { get; set; }
    public int FreezeColumns { get; set; }
    public List<double>? ColumnWidths { get; set; }
    public CellStyle? HeaderStyle { get; set; }
    public CellStyle? DefaultStyle { get; set; }
    public Dictionary<int, CellStyle>? RowStyles { get; set; }
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }
    public Dictionary<int, double>? RowHeights { get; set; }
    public Dictionary<string, string>? Comments { get; set; }
    public List<DataValidation>? Validations { get; set; }
    public List<ConditionalFormat> ConditionalFormats { get; set; } = new();
    public string? CodeName { get; set; }
    public List<WorksheetImage>? Images { get; set; }
    public SheetProtection? Protection { get; set; }
    public List<XlTable> Tables { get; set; } = new();
}
```

## B.2 `XlsxReader` (static class, zero reflection, AOT-safe)

↳ Main guide: Chapter 3 File Navigation: Open / Create / Save / Format (the low-level entry point for open / stream read), Chapter 21 Streaming Read / Progress Callback / Append (21.4 streaming read-back)

| Member | Description |
|---|---|
| `Read(path, sheetIndex[, firstRowIsHeader])` | Reads a single sheet by index |
| `Read(path, sheetName[, firstRowIsHeader])` | Reads a single sheet by name |
| `Read(stream, sheetIndex/name[, firstRowIsHeader])` | Reads a single sheet from a stream |
| `ReadAll(path)` / `ReadAll(stream)` | Reads all sheets |
| `Read<T>(path, sheetIndex/name[, configure])` | Reads into a List\<T\> |
| `ReadAsDataTable(path, sheetIndex/name[, firstRowIsHeader])` | Reads into a DataTable |
| `GetSheetNames(path)` / `GetSheetNames(stream)` | Lists worksheet names |
| `StreamRows(path/stream, sheetName, onRow)` | Streams rows one by one |
| `ReadWithProgress(path, sheetIndex, onProgress)` | Reads with progress reporting |
| `ReadProperties(path)` / `ReadProperties(stream)` | Reads document properties |

## B.3 `XlsxWriter` (static class, zero reflection, AOT-safe)

↳ Main guide: Chapter 3 File Navigation: Open / Create / Save / Format (the low-level write behind Excel.Write), Chapter 21 Streaming Read / Progress Callback / Append (21.3 the low-level write of Append)

| Member | Description |
|---|---|
| `Write(path, SheetData[, properties])` | Writes a single sheet |
| `Write(path, IReadOnlyList<SheetData>[, properties])` | Writes multiple sheets |
| `Write(stream, SheetData[, properties])` | Writes a single sheet to a stream |
| `Write(stream, IReadOnlyList<SheetData>[, properties])` | Writes multiple sheets to a stream |
| `Write<T>(path, data[, configure])` | Writes a List\<T\> |
| `Write(path, DataTable[, sheetName])` | Writes a DataTable |
| `Append(path, SheetData[, properties])` | Appends data |
| `AutoColumnWidths(SheetData)` | Auto-fits column widths |

Note: `XlsxWriter.Write` automatically writes out the macroEnabled main document type for the `.xlsm` extension; when writing a `SheetData`, the sheet name (`InvalidSheetNameException`) and duplicate sheet names (`LiteExcelException`) are validated.

## B.4 `CsvBackend` (internal class, CSV format backend)

↳ Main guide: Chapter 20 Multi-Format Behavior (20.3 CSV behavior)

> The low-level CSV backend is `internal`; for everyday CSV reads/writes, use `Excel.Open` / `Excel.Write`. The behavioral notes are listed here for reference.

- Implements a basic subset of RFC 4180: fields containing separators / line breaks / quotes are wrapped in double quotes.
- On read, the separator is auto-detected (comma > semicolon > Tab, counting only outside quotes); `ExcelReadOptions.Separator` can force a fixed one.
- On write, comma is the default; `ExcelWriteOptions.Separator` can be specified.
- Tabular data only; Excel-specific capabilities such as styles / merges / comments are not supported.

## B.5 `XlsxStreamWriter` (streaming writer)

↳ Main guide: Chapter 21 Streaming Read / Progress Callback / Append (21.4 streaming write via CreateWriter)

> Suited to writing large files row by row. Obtain it via `Excel.CreateWriter(path|stream)`.

| Member | Description |
|---|---|
| `Create(path)` / `Create(stream)` | Creates a writer |
| `WriteRow(IEnumerable<object?>)` | Writes one row of values |
| `WriteRow(IEnumerable<Cell>)` | Writes one row of Cells |
| `Close()` / `Dispose()` | Finalizes the file (must be called) |

- Uses inline strings (inlineStr) to avoid a pre-scan of the shared string table.
- Supports a single worksheet; styles / formulas / hyperlinks are written with each row (styles.xml and sheet rels are written out together at Close).
- Advanced capabilities such as merges / filters / images are not supported.
- Memory usage is no longer constant when the number of hyperlinks is extremely large.

## B.6 `CellRef` (A1 reference utility, static class)

↳ Main guide: Chapter 4 Cells and Values (accessing cells by A1 address)

| Member | Description |
|---|---|
| `Parse(cellRef)` | `"A1"` -> `(row=0, col=0)` |
| `TryParse(cellRef, out pos)` | Tries to parse |
| `ParseRange(range)` | Parses a range reference (0-based, inclusive) |
| `ToString(row, col)` | `(0,0)` -> `"A1"` |
| `ColToLetter(col)` | `0` -> `"A"` |
| `LetterToCol(letters)` | `"A"` -> `0` |

## B.7 Old vs. New API Mapping

A quick reference of equivalences between the object model API and the low-level API (the object model routes formats automatically by extension; the low-level API only handles xlsx/xlsm):

| Scenario | Object model API | Low-level API |
|---|---|---|
| Open a file | `Excel.Open(path[, options])` | `XlsxReader.Read(path, 0)` (single sheet) / `XlsxReader.ReadAll(path)` (all sheets) |
| Create and write out | `Excel.Create(...)` + `wb.SaveAs(path)` | `XlsxWriter.Write(path, sheet)` |
| Read a single sheet | `wb.Worksheets[i]` / `wb.Worksheets["name"]` | `XlsxReader.Read(path, sheetIndex)` / `XlsxReader.Read(path, sheetName)` |
| Read a List\<T\> | `Excel.Read<T>(path[, sheetName, configure])` | `XlsxReader.Read<T>(path, 0[, configure])` |
| Read a DataTable | `Excel.ReadAsDataTable(path[, sheetName])` | `XlsxReader.ReadAsDataTable(path, 0)` |
| Write a List\<T\> | `Excel.Write(path, list[, sheetName, configure])` | `XlsxWriter.Write(path, list)` |
| Write a DataTable | `Excel.Write(path, table[, sheetName])` | `XlsxWriter.Write(path, table, sheetName)` |
| Streaming read | `Excel.StreamRows(path, sheetName, onRow)` | `XlsxReader.StreamRows(path, sheetName, onRow)` |
| Streaming write | `Excel.CreateWriter(path)` | `XlsxStreamWriter.Create(path)` |
| Append data | `Excel.Append(path, sheetData[, properties])` | `XlsxWriter.Append(path, sheetData[, properties])` |
| List sheet names | `Excel.GetSheetNames(path)` | `XlsxReader.GetSheetNames(path)` |

> ⚠️ xls / xlsb / csv have no low-level API; always use the object model `Excel.Open` / `Excel.Write` (routed by extension).

---

*This guide covers all public capabilities of the current LiteExcel mainline.*
