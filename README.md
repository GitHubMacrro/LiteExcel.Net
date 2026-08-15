# LiteExcel

A lightweight, zero-dependency .NET library for reading and writing Excel xlsx/xlsm/csv files, plus reading legacy xls/xlsb files (xls read/write, xlsb read).
轻量级 xlsx/xlsm/csv 读写库（另支持 xls 读写、xlsb 读取），零第三方依赖，AOT 友好。

**Version / 版本**: 2.2.5

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
- xlsx / xlsm / csv read+write, xls read+write, xlsb read / 多格式支持（xlsx/xlsm/csv 读写、xls 读写、xlsb 读取）
- Streaming read/write for large files / 大文件流式读写
- Read/Write, styles, merged cells, auto filter, row height, comments, data validation, append, Stream, List\<T\>/DataTable
- 读/写、样式、合并、自动筛选、行高、批注、数据验证、追加、Stream、List\<T\>/DataTable

## Docs / 文档

- 📖 [Usage Guide / 使用手册](docs/USAGE.en.md) · [中文使用手册](docs/USAGE.zh-CN.md)
- 📝 [Changelog / 更新日志](docs/CHANGELOG.md)
