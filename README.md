# LiteExcel

A lightweight, zero-dependency .NET library for reading and writing Excel xlsx files.
轻量级 xlsx 读写库，零第三方依赖，AOT 友好。

**Version / 版本**: 2.1.3

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

var sheet = new SheetData
{
    SheetName = "Sheet1",
    Headers = new() { "Name", "Age" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Zhang San"), Cell.FromNumber(25) },
    },
};
XlsxWriter.Write("output.xlsx", sheet);

var read = XlsxReader.Read("output.xlsx", 0);
```

## Features / 特性

- Zero dependencies / 零依赖
- AOT friendly / AOT 友好
- net48 + net8.0
- Read/Write, styles, merged cells, auto filter, row height, comments, data validation, append, Stream, List\<T\>/DataTable
- 读/写、样式、合并、自动筛选、行高、批注、数据验证、追加、Stream、List\<T\>/DataTable

## Docs / 文档

- 📖 [Usage Guide / 使用手册](docs/USAGE.en.md) · [中文使用手册](docs/USAGE.zh-CN.md)
- 📝 [Changelog / 更新日志](docs/CHANGELOG.md)
