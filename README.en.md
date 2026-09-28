# LiteExcel.Net

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![Native AOT](https://img.shields.io/badge/Native%20AOT-tested-success)
![License](https://img.shields.io/badge/license-MIT-green)

A zero-dependency .NET library to read and write xlsx / xlsm / xlsb / xls / csv with one object model. Targets net48 and net8.0, needs no Office installation, and has passed a Native AOT smoke test.

> [中文 README](README.md)

## Features

- **Zero third-party dependencies**: built only on the .NET base class library, ready to use on reference, with no extra native components in the deploy package.
- **Targets net48 and net8.0**: one API for both legacy and modern projects.
- **Native AOT smoke test passed**: net8.0 sets `IsAotCompatible`; reflection is used only for `List<T>` mapping and is annotated with `[DynamicallyAccessedMembers]`. The repo ships a Native AOT publish smoke (`tests/LiteExcel.AotSmoke`) covering 10 cases such as `List<T>` mapping, DataTable, and conditional formatting, verified locally.
- **One object model across five formats**: the same code with a different format argument writes xls or csv.
- **Common spreadsheet operations**: styles and number formats, merge, row height / column width, auto filter, comments, data validation, hyperlinks, freeze panes, images, conditional formatting, tables, named ranges, formulas, file-level passwords, `List<T>` / DataTable mapping, and streaming read/write.
- **Advanced content is primarily preserved**: charts, pivot tables, slicers, VBA macros, Power Query, the data model, external connections, and ActiveX shapes are preserved along the established save paths on open-then-save, rather than created or modified as part of the object model.

### Five Formats at a Glance

| Format | Read | Create / Edit | Advanced content preservation | Streaming read | Streaming write |
|---|---|---|---|---|---|
| xlsx / xlsm | yes | yes | yes | yes | yes |
| xlsb | yes | partial (number format styles only) | yes (including save-as xlsx / xlsm) | yes | no |
| xls | yes | partial (number format styles only) | no (reported as degraded) | yes | no |
| csv | text values only | text only | — | no | no |

> `Excel.Open` / `Read<T>` / `ReadSheet` / `ReadAsDataTable` / `GetSheetNames` route by extension, and all five formats can be read; per-row streaming read supports xlsx / xlsm / xlsb / xls (not csv). For the item-by-item breakdown, see the [usage guide §20.1](docs/USAGE.en.md#201-format-capability-matrix).

### Capability Boundaries: Editable / Preserved / Degraded

LiteExcel.Net capabilities fall into three categories. These three are enough to judge how a given feature behaves in a given format:

- **Editable**: it can be read, modified, and written back.
- **Preserved**: it cannot be edited directly as part of the object model, but on open-then-save it is preserved as-is when the relevant save path applies.
- **Degraded**: when the target format cannot express it, the loss is **reported item by item** on save (query it via `Workbook.SaveDegradations`, or receive it through the `ExcelWriteOptions.OnDegradation` callback). Nothing is silently dropped.

Note these three distinctions in particular:

- **Editable ≠ Preserved**: content you can modify is not necessarily handled via preservation, and vice versa.
- **Preserved ≠ Editable**: advanced content can be preserved as-is without offering create / edit APIs.
- **Unsupported ≠ silently dropped**: unsupported capabilities are explicitly reported on write, not quietly discarded.

### Advanced Content and Cross-Format

Charts, pivot tables, slicers, VBA macros, Power Query, the data model, external connections, and ActiveX shapes are handled mainly through **preservation** — they cross the read/write process along the established save paths, rather than being created or modified as part of the object model.

- **Open-then-save preservation**: on xlsx / xlsm / xlsb the advanced content above is preserved along the established save paths; xls does no fidelity passthrough (reported as degraded).
- **Cross-format preservation (xlsb → xlsx / xlsm)**: when you open an `.xlsb` and save as `.xlsx` / `.xlsm`, pivot tables, slicers, Power Query, the data model, shapes, and other advanced content can be preserved through transcoding (enabled by default).

> Advanced content preservation depends on the established transcode / fidelity paths. On an `.xlsb` containing Power Query / the data model, some operations that trigger a full rebuild have known limits; see the [usage guide](docs/USAGE.en.md).

## Why LiteExcel.Net

- **Lightweight**: zero third-party dependencies, a single package, and no Excel / Office installation required.
- **Focused API**: one unified object model for common spreadsheet operations, shared across all five formats.
- **Explicit behavior**: capability boundaries are stated clearly — what is editable, what is only preserved, and what is degraded; unsupported content is never silently dropped.

## Preview

<details>
<summary>Preview (files written by LiteExcel.Net, opened in Excel)</summary>

[![Conditional formatting](docs/screenshots/conditional.png)](docs/screenshots/conditional.png)

[![Excel tables and filters](docs/screenshots/table_filter.png)](docs/screenshots/table_filter.png)

[![Images and freeze panes](docs/screenshots/image_freeze.png)](docs/screenshots/image_freeze.png)

[![Styles and number formats](docs/screenshots/style_number.png)](docs/screenshots/style_number.png)

[![Comments and data validation](docs/screenshots/comment_validation.png)](docs/screenshots/comment_validation.png)

[![Merged cells and hyperlinks](docs/screenshots/merge_link.png)](docs/screenshots/merge_link.png)

</details>

## Installation

The project / repository is named **LiteExcel.Net**; the NuGet package is named **LiteExcel**:

```powershell
dotnet add package LiteExcel
```

To use a locally packed nupkg, point the source at the package folder:

```powershell
dotnet add package LiteExcel --source .\packages
```

## Quick Start

The examples below target net48 and net8.0.

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

For `List<T>` mapping, DataTable, and low-level `SheetData`, see [usage guide chapter 2](docs/USAGE.en.md) and [Appendix B](docs/USAGE.en.md).

## Documentation

- [Usage Guide](docs/USAGE.en.md): full guide and examples
- [使用手册（中文）](docs/USAGE.zh-CN.md): Chinese usage guide
- [Changelog](docs/CHANGELOG.md): version history
- [中文 README](README.md)

The repo ships a console sample with 33 demos covering read/write, styles, filters, comments, encryption, images, conditional formatting, row/column insert-delete, formula write-back, and more. From the repo root:

```powershell
dotnet run --project demo/LiteExcel.Demo
```

Output goes to an `Output` folder under the program directory; the console prints the full path.

## Project Status

LiteExcel.Net is an actively maintained open-source project. The core API is becoming stable, while some advanced capabilities may continue to evolve.

Current known limits:

1. **CSV**: single sheet, plain text, no styles, all values read back as text.
2. **Passwords & macros**: xls has no password support; workbooks with macros can only be saved as xlsm or xlsb.
3. **Advanced content**: charts, pivot tables, etc. are preserved but not edited; xls / csv do not support them. On an `.xlsb` containing Power Query / the data model, some operations that trigger a full rebuild have known limits (they are explicitly reported).
4. **Streaming write & append**: `Excel.CreateWriter` / `Excel.Append` support xlsx / xlsm only (for the streaming read range, see "Five Formats at a Glance").
5. **Styles**: xls / xlsb support number formats only; other styles are reported as degraded on write.

## Roadmap

The following are directions, not commitments to a timeline or version:

- **Next**: improve reliability when deleting worksheets from `.xlsb` workbooks that contain Power Query / the data model (under research). A conservative safety net is already in place for saves that cannot preserve byte-for-byte and fall back to a full rebuild (lenient mode produces the file and reports; strict mode blocks); the underlying rebuild compatibility is still to be improved.
- **Candidate**: XLSB named-range write; XLSB InCell images; enhanced formula write-back (array formulas / 3D references / named ranges); streaming write for xlsb / xls.

## Contributing

Issues, discussions, and pull requests are welcome.

## License

MIT, see [LICENSE](LICENSE).
