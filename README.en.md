# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

A lightweight, zero-dependency .NET library for reading and writing xlsx/xlsm/csv files, plus legacy xls (read/write) and xlsb (read/write). AOT-friendly. **Version 2.4.7**

> [中文 README](README.zh-CN.md)

## Features

- **Zero dependencies**: .NET BCL only (ZipArchive + XmlReader/XDocument), no third-party packages
- **AOT friendly**: all public APIs are compatible with Native AOT/trimming; List\<T\> reflection mapping is annotated with `[DynamicallyAccessedMembers]` and verified by a native AOT executable (incl. xls/xlsb/csv, conditional formatting, named ranges, images, streaming)
- **Dual target**: net48 + net8.0 (works with legacy WinForms projects and new projects)
- **Intuitive object-model API**: natural hierarchy `Excel -> Workbook -> Worksheet -> Cell/Range/Cells`, one-liner read/write
- **Extensible formats**: xlsx/xlsm/csv read+write; xls read/write (formulas as static values); xlsb read/write (formulas as static values)
- **Full featured**: read/write, styles, merged cells, auto filter, row height/column width, comments, data validation, append, Stream, List\<T\>/DataTable convenience APIs, streaming read/write for large files
- **File-level security**: open password (Agile Encryption) + modify password (write protection) on xlsx/xlsm/xlsb, via `Workbook.Security` (`SetOpenPassword` / `SetModifyPassword` / `RemoveOpenPassword` / `RemoveModifyPassword`)
- **Sheet/Workbook protection** (2.4.6+): `sheetProtection` / `workbookProtection` (lock editing / lock structure, optional password, `ws.Protection` / `wb.Protection`)
- **Hyperlinks**: read + write on all 4 formats (xlsx/xlsm/xlsb/xls; external URLs/files/mailto/UNC + internal `#Sheet1!A1` jumps via `Cell.Hyperlink`)
- **Freeze panes**: `FreezeRows` / `FreezeColumns` on xlsx/xlsb/xls (arbitrary rows/columns), `FreezeHeader` compatible
- **Images (write-only)**: xlsx/xlsm floating images + in-cell images, with anchor/move-mode/AltText (`ws.AddImage`, `ImageAnchor`, `ImageMoveMode`; image reading is out of scope for 2.4.0)
- **Create workbook with data in one step**: `Excel.Create<T>` / `Excel.Create(DataTable)`, `Worksheet.ImportData`, `WorksheetCollection.Add<T>`
- **Named ranges read-back** (2.4.5+): `Workbook.Names`
- **Formula columns** (2.4.5+): `[LiteColumn(IsFormula = true)]` / `WriteOptions.Column(..., isFormula:)`
- **Super table** (2.4.6+): `Worksheet.AddTable` / `Tables` read+write (60 built-in Excel banded styles as an enum, plus arbitrary style names and per-column formats)
- **Icon-set conditional formatting** (2.4.6+): `ConditionalFormatType.IconSet` (17 built-in sets as an enum — arrows/traffic lights/symbols/ratings — with custom thresholds)
- **Real file compatibility**: correctly reads xlsx created by Excel/WPS (including Table/theme extension parts)

## Installation

```powershell
dotnet add package LiteExcel
```

Or reference a local .nupkg:

```xml
<PackageReference Include="LiteExcel" Version="2.4.2" />
```

## Quick Start (recommended: object-model API)

```csharp
using LiteExcel;

// Create a workbook, write through a natural hierarchy
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "Name");
ws.SetValue("B1", "Age");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B2", 25);
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
wb.SaveAs("output.xlsx");

// Open and read
var opened = Excel.Open("output.xlsx");
var name = opened.Worksheets[0].Cell("A2").GetString();   // "Zhang San"
var age = opened.Worksheets[0].Cells[2, 2].GetDouble();   // 25

// Modify and save
opened.Worksheets[0].SetValue("B2", 26);
opened.Save();
```

## XlsxWriter / XlsxReader (classic API)

```csharp
// Write
var sheet = new SheetData
{
    SheetName = "Sheet1",
    Headers = new() { "Name", "Age", "Birthday" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Zhang San"), Cell.FromNumber(25), Cell.FromDate(new DateTime(2000, 1, 1)) },
        new Cell[] { Cell.FromText("Li Si"), Cell.FromNumber(30), Cell.FromDate(new DateTime(1995, 5, 10)) },
    },
    FreezeHeader = true,          // freeze header
    ColumnWidths = new() { 12, 8, 14 },  // column widths
};
XlsxWriter.Write("output.xlsx", sheet);

// Read
var read = XlsxReader.Read("output.xlsx", 0);
foreach (var row in read.Rows)
{
    Console.WriteLine($"{row[0].Text}, {row[1].Number}, {row[2].Date:yyyy-MM-dd}");
}
```

## Cell Types

| Type | Description |
|---|---|
| `Text` | Text (shared/inline strings) |
| `Number` | Number (exact up to 12 digits) |
| `Date` | Date (auto-detects Excel date formats, supports 1900/1904 systems) |
| `Boolean` | Boolean |
| `Empty` | Empty cell |
| Formula | Reads the `<f>` formula string + cached `<v>` value; writes formula strings without calculating |

## API Reference

### Object-Model API (recommended)

| Type / Method | Description |
|---|---|
| `Excel.Open(path)` | open a workbook, format auto-detected from extension |
| `Excel.Open(stream, format)` | open a workbook from stream (format must be specified) |
| `Excel.Create(format)` / `Excel.Create(sheetName, format)` / `Excel.Create(sheetNames[], format)` | create a workbook (xlsx/xlsm/csv/xls/xlsb), batch sheet creation supported |
| `Excel.Create<T>(data, sheetName, format, configure?)` / `Excel.Create(dataTable, ...)` | create a workbook and write List\<T\> / DataTable data directly (first row is header, AOT safe) |
| `Excel.Read<T>(path, sheetName?)` | read as List\<T\> (reflection, AOT safe) |
| `Excel.ReadAsDataTable(path, sheetName?)` | read as DataTable (AOT safe) |
| `Excel.Write(path, Workbook)` | write a workbook |
| `Excel.Write(path, DataTable)` | write from DataTable (AOT safe) |
| `Excel.Write<T>(path, IEnumerable<T>)` | write from List\<T\> (reflection, AOT safe) |
| `Excel.CreateWriter(path/stream)` | streaming write for large files (row by row) |
| `Excel.StreamRows(path, sheetName, onRow)` | streaming read for large files |
| `Excel.GetSheetNames(path)` | list all sheet names |
| `Workbook.Worksheets / Properties / Save() / SaveAs(path, format)` | file-level operations |
| `Worksheet.Cell("A1") / Cell(row, col) / Range("A1:D10") / Cells` | sheet-level access |
| `Worksheet.ImportData<T>(data, configure?)` / `ImportData(dataTable)` | clear the sheet and rebuild from A1 (List\<T\> / DataTable) |
| `Worksheet.SetValue / Merge / Unmerge` | sheet-level write and merge |
| `Cell.GetString/GetDouble/GetDateTime/GetBoolean/TryGet*` | typed cell accessors |
| `Cell.SetValue / IsFormula / FromFormula` | cell write and formula |
| `Cell.Style / Cell.NumberFormat` | style & number format of a specific cell (background/font/alignment etc.) |
| `Worksheet.Cell("A2").Style = ...` / `Worksheet.Range("A2:C3").Style = ...` | restyle a single cell / range |
| `Worksheet.Comments["A2"] = "note"` | add a comment to a specific cell |
| `Cells[row, col] / Cells["A1"] / Cells.Range(...)` | collection access |
| `ExcelRange.Fill / Clear / Style / Merge / ToValues` | range operations |

### XlsxWriter

| Method | Description |
|---|---|
| `Write(path, SheetData)` / `Write(stream, SheetData)` | write single sheet |
| `Write(path, IReadOnlyList<SheetData>)` / `Write(stream, ...)` | write multiple sheets |
| `Write(path/stream, sheets, WorkbookProperties)` | write with document properties |
| `Write(path, DataTable)` | write from DataTable (AOT safe) |
| `Write<T>(path, IEnumerable<T>)` | write from List\<T\> (reflection, AOT safe) |
| `Append(path, SheetData, WorkbookProperties?)` | append data and optionally update document properties |
| `AutoColumnWidths(sheet)` | auto estimate column widths |

### XlsxReader

| Method | Description |
|---|---|
| `GetSheetNames(path)` / `GetSheetNames(stream)` | list all sheet names |
| `Read(path, sheetIndex)` / `Read(stream, sheetIndex)` | read single sheet by index |
| `Read(path, sheetName)` / `Read(stream, sheetName)` | read single sheet by name |
| `ReadAll(path)` / `ReadAll(stream)` | read all worksheets |
| `StreamRows(path, sheetName, onRow)` / `StreamRows(stream, ...)` | stream large files |
| `ReadWithProgress(path, sheetIndex, onProgress)` | read with progress |
| `ReadAsDataTable(path)` / `ReadAsDataTable(stream)` | read as DataTable (AOT safe) |
| `Read<T>(path)` | read as List\<T\> (reflection, AOT safe) |
| `ReadProperties(path)` / `ReadProperties(stream)` | read document properties (author/time/title) |

### Model Classes

| Class | Description |
|---|---|
| `Workbook` / `Worksheet` / `Cells` / `ExcelRange` | object models |
| `ExcelFormat` / `ExcelReadOptions` / `ExcelWriteOptions` | formats and options |
| `SheetData` | worksheet data (headers/rows/styles/merge/filter/height/comments/validation) |
| `Cell` | cell |
| `CellStyle` / `BorderStyle` / `BorderEdge` | styles |
| `CellRange` | range (merged cells) |
| `AutoFilter` / `FilterColumn` | auto filter |
| `DataValidation` | data validation |
| `Hyperlink` / `Cell.Hyperlink` | cell hyperlink (external URL / internal jump) |
| `WorksheetImage` / `ImagePlacement` / `ImageAnchor` / `ImageMoveMode` | image (floating / in-cell, anchor / move-mode / alt-text) |
| `WorkbookSecurity` / `Workbook.Security` | file-level security (open password / modify password / read-only state) |
| `WorkbookProperties` | document properties (author/time/title/app name) |
| `LiteExcelException` / `InvalidSheetNameException` | exceptions |
| `LiteColumnAttribute` | List\<T\> column attribute |
| `WriteOptions<T>` / `ReadOptions<T>` | List\<T\> configuration |

## Target Frameworks

- net48
- net8.0

## AOT Compatibility

| API | AOT |
|---|---|
| `Excel.Open` / `Workbook` / `Worksheet` / `Cell` / `ExcelRange` / `Cells` | ✅ Safe (no reflection) |
| `Excel.ReadAsDataTable` / `DataTable` read/write | ✅ Safe (no reflection) |
| `SheetData` / `XlsxWriter` / `XlsxReader` read/write | ✅ Safe (no reflection) |
| `Excel.Read<T>` / `Excel.Write<T>` / `List<T>` read/write | ✅ Safe (reflection mapping annotated with `[DynamicallyAccessedMembers]`) |

## Detailed Docs

- 📖 [Usage Guide](docs/USAGE.en.md) — full API reference and feature examples (object-model API/styles/merge/filter/comments/data validation/Stream, etc.)
- 📝 [Changelog](docs/CHANGELOG.md) — version history
- 🌐 [中文 README](README.zh-CN.md)
