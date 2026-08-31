# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet Downloads](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

A lightweight, zero-dependency .NET library for reading and writing xlsx / xlsm / xlsb / xls / csv without installing Excel.
无需安装 Excel 即可读写 xlsx / xlsm / xlsb / xls / csv 的轻量级 .NET 库，零第三方依赖，AOT 友好。

## Language / 语言

- [English README](README.en.md)
- [中文 README](README.zh-CN.md)

## Preview / 效果预览

Files written by LiteExcel, opened in Excel / 以下均为 LiteExcel 写出的文件在 Excel 中打开的效果：

[![Conditional formatting / 条件格式](https://raw.githubusercontent.com/GitHubMacrro/LiteExcel.Net/main/docs/screenshots/conditional.png)](https://raw.githubusercontent.com/GitHubMacrro/LiteExcel.Net/main/docs/screenshots/conditional.png)

[![Excel tables and filters / 超级表与筛选](https://raw.githubusercontent.com/GitHubMacrro/LiteExcel.Net/main/docs/screenshots/table_filter.png)](https://raw.githubusercontent.com/GitHubMacrro/LiteExcel.Net/main/docs/screenshots/table_filter.png)

More screenshots are in the [中文 README](README.zh-CN.md) and the [usage guide §15 / §16](docs/USAGE.zh-CN.md). / 更多截图见[中文 README](README.zh-CN.md) 与[使用手册 §15 / §16](docs/USAGE.zh-CN.md)。

## Docs / 文档

- [Usage Guide / 使用手册](docs/USAGE.en.md) · [中文使用手册](docs/USAGE.zh-CN.md)
- [Changelog / 更新日志](docs/CHANGELOG.md)

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

- Zero dependencies, built only on the .NET base class library / 零依赖，仅用 .NET 基础类库
- net48 + net8.0, all public APIs Native AOT / trim compatible / 双目标，公开 API 兼容 Native AOT 与裁剪
- One object model across five formats / 一套对象模型覆盖五种格式
- Styles, merge, filter, row/column sizing, comments, validation, hyperlinks, freeze, images, conditional formatting, tables, formulas, passwords / 样式、合并、筛选、行高列宽、批注、数据验证、超链接、冻结、图片、条件格式、超级表、公式、文件密码
- Open-then-save preserves macros, charts, pivot tables for xlsx / xlsm / xlsb / 打开再保存透传保留宏、图表、透视表
- Large-file streaming read/write / 大文件流式读写
- Capabilities the target format lacks are reported, never silently dropped / 目标格式不支持的能力显式上报，不静默丢弃

## Capability Matrix / 能力矩阵

Legend / 图例：☑️ supported / 支持 · ❌ not supported / 不支持 · text = partial / 文字表示部分支持

| Capability / 能力 | xlsx | xlsm | xlsb | xls | csv |
|---|---|---|---|---|---|
| Cell read/write / 数据读写 | ☑️ | ☑️ | ☑️ | ☑️ | text only / 纯文本 |
| Styles & number formats / 样式与数字格式 | ☑️ | ☑️ | number format only / 仅数字格式 | number format only / 仅数字格式 | ❌ |
| Layout / 表格布局 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Auto filter / 自动筛选 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Comments / 批注 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Data validation / 数据验证 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Hyperlinks / 超链接 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Freeze panes / 冻结窗格 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| Images / 图片 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Conditional formatting / 条件格式 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Tables / 超级表 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| Formulas / 公式 | ☑️ | ☑️ | read only / 仅读取 | read only / 仅读取 | ❌ |
| File passwords / 文件密码 | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| Charts / pivot tables / 图表 / 透视表 | passthrough / 原样保留 | passthrough / 原样保留 | passthrough / 原样保留 | ❌ | ❌ |
| Large-file streaming / 大文件流式 | ☑️ | ☑️ | ❌ | ❌ | ❌ |

> Full details in the usage guide / 完整能力明细见使用手册 §20.1：[中文](docs/USAGE.zh-CN.md) · [English](docs/USAGE.en.md)

## License
- [License / 许可证](LICENSE)