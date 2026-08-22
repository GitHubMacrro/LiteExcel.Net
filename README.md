# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

A lightweight, zero-dependency .NET library for reading and writing Excel xlsx/xlsm/csv files, plus legacy xls/xlsb files (xls read/write, xlsb read/write).
轻量级 xlsx/xlsm/csv 读写库（另支持 xls 读写、xlsb 读写），零第三方依赖，AOT 友好。

**Version / 版本**: 2.4.4

## Language / 语言

- [English Documentation / English README](README.en.md)
- [中文文档 / 中文 README](README.zh-CN.md)

## Install / 安装

```powershell
dotnet add package LiteExcel
```

## Quick Start / 快速上手

```csharp
using LiteExcel;

// Object-model API / 对象模型 API
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", "Name");
wb.Worksheets["Sheet1"].SetValue("A2", "Zhang San");
wb.SaveAs("output.xlsx");

var opened = Excel.Open("output.xlsx");
var name = opened.Worksheets[0].Cell("A2").GetString();

// Classic API / 经典 API（XlsxWriter / XlsxReader）
var sheet = new SheetData
{
    SheetName = "Sheet1",
    Headers = new() { "Name", "Age" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Zhang San"), Cell.FromNumber(25) },
    },
};
XlsxWriter.Write("output2.xlsx", sheet);
var read = XlsxReader.Read("output2.xlsx", 0);
```

## Features / 特性

- Zero dependencies / 零依赖
- AOT friendly / AOT 友好
- net48 + net8.0
- Object-model API `Excel -> Workbook -> Worksheet -> Cell/Range/Cells` / 直觉化对象模型 API
- xlsx / xlsm / csv read+write, xls read+write, xlsb read/write / 多格式支持（xlsx/xlsm/csv 读写、xls 读写、xlsb 读写）
- Streaming read/write for large files / 大文件流式读写
- Read/Write, styles, merged cells, auto filter, row height, comments, data validation, append, Stream, List\<T\>/DataTable
- 读/写、样式、合并、自动筛选、行高、批注、数据验证、追加、Stream、List\<T\>/DataTable
- File-level security: open password (Agile Encryption) + modify password (write protection) on xlsx/xlsm/xlsb, via `Workbook.Security` (`SetOpenPassword` / `SetModifyPassword` / `RemoveOpenPassword` / `RemoveModifyPassword`) / 文件级安全：打开密码（Agile 加密）+ 修改密码（写保护），支持 xlsx/xlsm/xlsb
- Hyperlinks: read + write on all 4 formats (xlsx/xlsm/xlsb/xls; external URLs + internal `#Sheet1!A1` jumps via `Cell.Hyperlink`) / 超链接：xlsx/xlsm/xlsb/xls 四格式读写（外部链接 + 内部跳转）
- Freeze panes: `FreezeRows` / `FreezeColumns` on xlsx/xlsb/xls (arbitrary rows/columns), `FreezeHeader` compatible / 冻结窗格：xlsx/xlsb/xls 任意行列冻结，`FreezeHeader` 兼容
- Images: floating image read-back (2.4.3+) + write Floating/InCell on xlsx/xlsm, with anchor/move-mode/AltText (`ws.AddImage`, `ImageAnchor`, `ImageMoveMode`, `Worksheet.Images`) / 图片：xlsx/xlsm 浮动图片读回（2.4.3+）与浮动/内嵌写回（锚点/移动方式/AltText）
- Conditional formatting (2.4.3+): cellIs/expression/colorScale/dataBar + (2.4.4) containsText/beginsWith/endsWith/timePeriod/blank/error/unique/duplicate/top10/above-average/below-average read+write on xlsx/xlsm (`Worksheet.ConditionalFormats`) / 条件格式：cellIs/expression/colorScale/dataBar + containsText 等长尾类型读写
- CSV separator options + auto-detect (2.4.3+): `ExcelReadOptions.Separator` / `ExcelWriteOptions.Separator` (comma/semicolon/tab) / CSV 分隔符选项与自动探测（逗号/分号/Tab）
- Capability degradation reporting (2.4.2+): `ExcelWriteOptions.OnDegradation` for explicit drops on xls/xlsb/csv / 能力降级回调：写出到 xls/xlsb/csv 时的显式降级上报

## Docs / 文档

- 📖 [Usage Guide / 使用手册](docs/USAGE.en.md) · [中文使用手册](docs/USAGE.zh-CN.md)
- 📝 [Changelog / 更新日志](docs/CHANGELOG.md)
