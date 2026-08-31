# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

A lightweight .NET library to read and write xlsx / xlsm / xlsb / xls / csv without installing Excel. Zero third-party dependencies, targets net48 and net8.0, AOT friendly.

> [中文 README](README.zh-CN.md)

## Docs

- [Usage Guide](docs/USAGE.en.md): full API reference and examples
- [Changelog](docs/CHANGELOG.md): version history
- [中文 README](README.zh-CN.md)
## Preview

Below are files written by LiteExcel, opened in Excel:

[![Conditional formatting](docs/screenshots/conditional.png)](docs/screenshots/conditional.png)

[![Excel tables and filters](docs/screenshots/table_filter.png)](docs/screenshots/table_filter.png)

[![Images and freeze panes](docs/screenshots/image_freeze.png)](docs/screenshots/image_freeze.png)

<details>
<summary>More: styles and number formats · comments and validation · merged cells and hyperlinks</summary>

[![Styles and number formats](docs/screenshots/style_number.png)](docs/screenshots/style_number.png)

[![Comments and data validation](docs/screenshots/comment_validation.png)](docs/screenshots/comment_validation.png)

[![Merged cells and hyperlinks](docs/screenshots/merge_link.png)](docs/screenshots/merge_link.png)

</details>

## Features

- Zero dependencies, built only on the .NET base class library, ready to use on reference with no extra native components in the deploy package.
- Targets net48 and net8.0, all public APIs are Native AOT / trim compatible, verified by a native executable.
- One object model across five formats; the same code with a different format argument writes xls or csv.
- Covers common office needs: styles, number formats, merge, filter, row/column sizing, comments, data validation, hyperlinks, freeze panes, images, conditional formatting, tables, named ranges, formulas, file passwords, large-file streaming.
- Open-then-save preserves untouched parts; macros, charts, and pivot tables pass through for xlsx / xlsm / xlsb.
- File-level security: open and modify passwords, sheet and workbook protection with optional password.
- Streaming read and write keep memory flat for large files.
- When writing to xls / xlsb / csv, capabilities the target format lacks are reported item by item, never silently dropped.

## Install

```powershell
dotnet add package LiteExcel
```

To use a locally packed nupkg, point the source at the package folder:

```powershell
dotnet add package LiteExcel --source .\packages
```

## Quick Start

**Object model**: create a workbook, write by natural hierarchy, then open and read.

```csharp
using LiteExcel;

var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "Name");
ws.SetValue("B1", "Age");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B2", 25);
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
wb.SaveAs("output.xlsx");

var opened = Excel.Open("output.xlsx");
var name = opened.Worksheets[0].Cell("A2").GetString();
var age = opened.Worksheets[0].Cells[2, 2].GetDouble();
```

For `List<T>` mapping, DataTable, and low-level `SheetData`, see the [usage guide](docs/USAGE.en.md).

## Capability Matrix

Legend: ☑️ supported · ❌ not supported · text in a cell means partial support

| Capability | xlsx | xlsm | xlsb | xls | csv |
|---|---|---|---|---|---|
| Cell read/write | ☑️ | ☑️ | ☑️ | ☑️ | text only |
| Styles & number formats | ☑️ | ☑️ | number format only | number format only | ❌ |
| Layout (merge / row height / column width) | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Auto filter | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Comments | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Data validation | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Hyperlinks | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Freeze panes | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Images | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Conditional formatting | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Tables | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Formulas | ☑️ | ☑️ | read only | read only | ❌ |
| File passwords | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| Charts / pivot tables | passthrough | passthrough | passthrough | ❌ | ❌ |
| Large-file streaming | ☑️ | ☑️ | ❌ | ❌ | ❌ |

> The full 26-item breakdown is in the [usage guide §20.1](docs/USAGE.en.md#201-format-capability-matrix).

## Compatibility

- Target frameworks: net48, net8.0
- AOT: all public APIs are Native AOT / trim compatible; `List<T>` reflection mapping is annotated

## Known Limits

1. **Read entry**: `Excel.Read<T>` supports xlsx / xlsm only; for xls / xlsb / csv use `Excel.Open` which routes by extension.
2. **CSV**: single sheet, plain text, no styles, all values read back as text.
3. **Passwords & macros**: xls has no password support; workbooks with macros can only be saved as xlsm or xlsb.
4. **Charts & pivot tables**: preserved but not edited; xlsx / xlsm / xlsb keep them on open-then-save, xls / csv drop them.
5. **Streaming & append**: xlsx / xlsm only.

## Run the Demo

The repo ships a console sample with 31 demos covering read/write, styles, filters, comments, encryption, images, conditional formatting, and more. From the repo root:

```powershell
dotnet run --project demo/LiteExcel.Demo
```

Output goes to an `Output` folder under the program directory; the console prints the full path.



## License

MIT, see [LICENSE](LICENSE).