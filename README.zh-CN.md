# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet 下载](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

轻量级 .NET 库，无需安装 Excel 即可读写 xlsx / xlsm / xlsb / xls / csv 五种格式。零第三方依赖，net48 与 net8.0 双目标，AOT 友好。

> [English README](README.en.md)

## 效果预览

以下均为 LiteExcel 写出的文件在 Excel 中打开的实际效果：

[![条件格式效果](docs/screenshots/conditional.png)](docs/screenshots/conditional.png)

[![超级表与筛选效果](docs/screenshots/table_filter.png)](docs/screenshots/table_filter.png)

[![图片与冻结窗格效果](docs/screenshots/image_freeze.png)](docs/screenshots/image_freeze.png)

<details>
<summary>更多效果：样式与数字格式 · 批注与数据验证 · 合并与超链接</summary>

[![样式与数字格式](docs/screenshots/style_number.png)](docs/screenshots/style_number.png)

[![批注与数据验证](docs/screenshots/comment_validation.png)](docs/screenshots/comment_validation.png)

[![合并与超链接](docs/screenshots/merge_link.png)](docs/screenshots/merge_link.png)

</details>

## 详细文档

- [使用手册（中文）](docs/USAGE.zh-CN.md)：完整 API 参考与全部功能示例
- [更新日志](docs/CHANGELOG.md)：版本变更记录
- [English README](README.en.md)

## 特性

- 零依赖，仅用 .NET 基础类库，引用即用，部署包里没有额外原生组件。
- net48 与 net8.0 双目标，公开 API 全部兼容 Native AOT 与裁剪，经原生可执行文件实测。
- 一套对象模型覆盖五种格式，同样的代码换个格式参数就能写出 xls 或 csv。
- 覆盖常用办公能力：样式、数字格式、合并、筛选、行高列宽、批注、数据验证、超链接、冻结窗格、图片、条件格式、超级表、命名区域、公式、文件密码、大文件流式。
- 打开再保存保留未改动部件，xlsx / xlsm / xlsb 的宏、图表、透视表原样透传。
- 文件级安全：打开密码与修改密码，工作表与工作簿保护可带密码。
- 大文件流式读写不占内存。
- 写 xls / xlsb / csv 时，目标格式不支持的能力逐项显式上报，不静默丢弃。

## 安装

```powershell
dotnet add package LiteExcel
```

使用本地打包的 nupkg 时，指定包所在目录作为源：

```powershell
dotnet add package LiteExcel --source .\packages
```

## 快速上手

**对象模型读写**：新建工作簿，按自然层级写入，再打开读取。

```csharp
using LiteExcel;

var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "姓名");
ws.SetValue("B1", "年龄");
ws.SetValue("A2", "张三");
ws.SetValue("B2", 25);
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
wb.SaveAs("output.xlsx");

var opened = Excel.Open("output.xlsx");
var name = opened.Worksheets[0].Cell("A2").GetString();
var age = opened.Worksheets[0].Cells[2, 2].GetDouble();
```

`List<T>` 映射、DataTable、低层 SheetData 的写法见[使用手册第 2 章](docs/USAGE.zh-CN.md#2-快速上手)与[附录 B](docs/USAGE.zh-CN.md#附录-b-低层-api-参考)。

## 能力矩阵

图例：☑️ 支持 · ❌ 不支持 · 单元格内文字表示部分支持

| 能力 | xlsx | xlsm | xlsb | xls | csv |
|---|---|---|---|---|---|
| 数据读写 | ☑️ | ☑️ | ☑️ | ☑️ | 纯文本 |
| 样式与数字格式 | ☑️ | ☑️ | 仅数字格式 | 仅数字格式 | ❌ |
| 表格布局（合并 / 行高 / 列宽） | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 自动筛选 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 批注 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 数据验证 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 超链接 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 冻结窗格 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 图片 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 条件格式 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 超级表 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 公式 | ☑️ | ☑️ | 仅读取 | 仅读取 | ❌ |
| 文件密码 | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| 图表 / 透视表 | 原样保留 | 原样保留 | 原样保留 | ❌ | ❌ |
| 大文件流式读写 | ☑️ | ☑️ | ❌ | ❌ | ❌ |

> 完整 26 项能力明细见[使用手册 §20.1](docs/USAGE.zh-CN.md#201-格式能力矩阵)。

## 兼容性

- 目标框架：net48、net8.0
- AOT：公开 API 全部兼容 Native AOT 与裁剪；`List<T>` 反射映射已标注

## 已知边界

1. **读取入口**：`Excel.Read<T>` 只支持 xlsx / xlsm；xls / xlsb / csv 用 `Excel.Open` 按扩展名路由。
2. **CSV**：单工作表、纯文本、不带样式，所有值按文本读回。
3. **密码与宏**：xls 不支持密码；含宏工作簿只能存为 xlsm 或 xlsb。
4. **图表与透视表**：只保留不编辑，xlsx / xlsm / xlsb 打开再保存原样保留，xls / csv 会丢。
5. **流式与追加**：只支持 xlsx / xlsm。

## 运行 Demo

仓库自带控制台示例，31 个 Demo 覆盖写读 / 样式 / 筛选 / 批注 / 加密 / 图片 / 条件格式等用法。在仓库根目录执行：

```powershell
dotnet run --project demo/LiteExcel.Demo
```

输出写到程序运行目录的 `Output` 文件夹，控制台会打印完整路径。

## License

MIT [LICENSE](LICENSE)。