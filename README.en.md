# LiteExcel

A lightweight, zero-dependency .NET library for reading and writing Excel xlsx files. AOT-friendly. **Version 2.1.3**

> [中文 README](README.zh-CN.md)

## Features

- **Zero dependencies**: .NET BCL only (ZipArchive + XmlReader/XDocument), no third-party packages
- **AOT friendly**: low-level APIs have no reflection; high-level reflection APIs are marked `[RequiresUnreferencedCode]`
- **Dual target**: net48 + net8.0 (works with legacy WinForms projects and new projects)
- **Full featured**: read/write, styles (4-level priority), merged cells, auto filter, row height/column width, comments, data validation, append, Stream, List\<T\>/DataTable convenience APIs
- **Real file compatibility**: correctly reads xlsx created by Excel/WPS (including Table/theme extension parts)

## Installation

```powershell
dotnet add package LiteExcel
```

Or reference a local .nupkg:

```xml
<PackageReference Include="LiteExcel" Version="2.1.3" />
```

## Quick Start

```csharp
using LiteExcel;

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
| `Number` | Number (exact long up to 12 digits) |
| `Date` | Date (auto-detects Excel date formats, supports 1900/1904 systems) |
| `Boolean` | Boolean |
| `Empty` | Empty cell |
| Formula result | Reads cached `<v>` value, not the formula expression |

## API Reference

### XlsxWriter

| Method | Description |
|---|---|
| `Write(path, SheetData)` / `Write(stream, SheetData)` | write single sheet |
| `Write(path, IReadOnlyList<SheetData>)` / `Write(stream, ...)` | write multiple sheets |
| `Write(path/stream, sheets, WorkbookProperties)` | write with document properties |
| `Write(path, DataTable)` | write from DataTable (AOT safe) |
| `Write<T>(path, IEnumerable<T>)` | write from List\<T\> (reflection, not AOT compatible) |
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
| `Read<T>(path)` | read as List\<T\> (reflection, not AOT compatible) |
| `ReadProperties(path)` / `ReadProperties(stream)` | read document properties (author/time/title) |

### Model Classes

| Class | Description |
|---|---|
| `SheetData` | worksheet data (headers/rows/styles/merge/filter/height/comments/validation) |
| `Cell` | cell |
| `CellStyle` / `BorderStyle` / `BorderEdge` | styles |
| `CellRange` | range (merged cells) |
| `AutoFilter` / `FilterColumn` | auto filter |
| `DataValidation` | data validation |
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
| `SheetData` / `DataTable` read/write | ✅ Safe (no reflection) |
| `List<T>` read/write (`Read<T>` / `Write<T>`) | ⚠️ Marked `[RequiresUnreferencedCode]`, warns on AOT compile |

## Detailed Docs

- 📖 [Usage Guide](docs/USAGE.en.md) — full API reference and feature examples (styles/merge/filter/comments/data validation/Stream, etc.)
- 📝 [Changelog](docs/CHANGELOG.md) — version history
- 🌐 [中文 README](README.zh-CN.md)
