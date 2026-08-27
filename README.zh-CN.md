# LiteExcel

[![NuGet](https://img.shields.io/nuget/v/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![NuGet 下载](https://img.shields.io/nuget/dt/LiteExcel)](https://www.nuget.org/packages/LiteExcel)
[![CI](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml/badge.svg)](https://github.com/GitHubMacrro/LiteExcel.Net/actions/workflows/ci.yml)
![.NET](https://img.shields.io/badge/.NET-8.0%20%7C%204.8-512BD4)
![License](https://img.shields.io/badge/license-MIT-green)

轻量级 xlsx/xlsm/csv 读写库（另支持 xls 读写、xlsb 读写），零第三方依赖，AOT 友好。**版本 2.4.7**

> [English README](README.en.md)

## 特性

- **零依赖**：仅用 .NET BCL（ZipArchive + XmlReader/XDocument），无任何第三方包
- **AOT 友好**：全部公开 API 兼容 Native AOT/裁剪；List\<T\> 反射映射已用 `[DynamicallyAccessedMembers]` 标注，经原生可执行文件实测（含 xls/xlsb/csv、条件格式、命名区域、图片、流式）
- **双目标**：net48 + net8.0（老 WinForms 项目与新项目都能用）
- **直觉化对象模型 API**：`Excel -> Workbook -> Worksheet -> Cell/Range/Cells` 自然层级，一行式读写
- **格式可扩展**：xlsx/xlsm/csv 读写；xls 读写（公式降级为静态值）；xlsb 读写（公式降级为静态值）
- **全功能**：读/写、样式、合并单元格、自动筛选、行高/列宽、批注、数据验证、追加、Stream、List\<T\>/DataTable 便利 API、流式读写大文件
- **文件级安全**：打开密码（Agile 加密）+ 修改密码（写保护），支持 xlsx/xlsm/xlsb，通过 `Workbook.Security`（`SetOpenPassword` / `SetModifyPassword` / `RemoveOpenPassword` / `RemoveModifyPassword`）
- **工作表/工作簿保护**（2.4.6+）：`sheetProtection` / `workbookProtection`（锁编辑/锁结构，可选密码，`ws.Protection` / `wb.Protection`）
- **超链接**：xlsx/xlsm/xlsb/xls 四格式读写（外部 URL/文件/mailto/UNC + 内部 `#Sheet1!A1` 跳转，`Cell.Hyperlink`）
- **冻结窗格**：`FreezeRows` / `FreezeColumns` 支持 xlsx/xlsb/xls 任意行列冻结，`FreezeHeader` 兼容
- **图片写回**：xlsx/xlsm 浮动图片 + 单元格内嵌图片，支持锚点/移动方式/AltText（`ws.AddImage`、`ImageAnchor`、`ImageMoveMode`；图片读取不在 2.4.0 范围）
- **一步建簿并写数据**：`Excel.Create<T>` / `Excel.Create(DataTable)` 建簿即写数据，`Worksheet.ImportData` 清空重建，`WorksheetCollection.Add<T>` 批量加表
- **命名区域读回**（2.4.5+）：`Workbook.Names`
- **公式列**（2.4.5+）：`[LiteColumn(IsFormula = true)]` / `WriteOptions.Column(..., isFormula:)`
- **超级表**（2.4.6+）：`Worksheet.AddTable` / `Tables` 读写（60 种 Excel 内置条纹样式枚举，支持任意样式名与列级格式）
- **图标集条件格式**（2.4.6+）：`ConditionalFormatType.IconSet`（17 种内置集合枚举，箭头/红绿灯/符号/星级，可自定义阈值）
- **真实文件兼容**：可正确读取 Excel/WPS 创建的 xlsx（含 Table/theme 等扩展部件）

## 安装

```powershell
dotnet add package LiteExcel
```

或本地 .nupkg 引用：

```xml
<PackageReference Include="LiteExcel" Version="2.4.2" />
```

## 快速上手（推荐：对象模型 API）

```csharp
using LiteExcel;

// 新建工作簿，自然层级写入
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "姓名");
ws.SetValue("B1", "年龄");
ws.SetValue("A2", "张三");
ws.SetValue("B2", 25);
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
wb.SaveAs("output.xlsx");

// 打开并读取
var opened = Excel.Open("output.xlsx");
var name = opened.Worksheets[0].Cell("A2").GetString();   // "张三"
var age = opened.Worksheets[0].Cells[2, 2].GetDouble();   // 25

// 修改并保存
opened.Worksheets[0].SetValue("B2", 26);
opened.Save();
```

## XlsxWriter / XlsxReader（经典 API）

```csharp
// 写出
var sheet = new SheetData
{
    SheetName = "Sheet1",
    Headers = new() { "姓名", "年龄", "生日" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("张三"), Cell.FromNumber(25), Cell.FromDate(new DateTime(2000, 1, 1)) },
        new Cell[] { Cell.FromText("李四"), Cell.FromNumber(30), Cell.FromDate(new DateTime(1995, 5, 10)) },
    },
    FreezeHeader = true,          // 冻结表头
    ColumnWidths = new() { 12, 8, 14 },  // 列宽
};
XlsxWriter.Write("output.xlsx", sheet);

// 读回
var read = XlsxReader.Read("output.xlsx", 0);
foreach (var row in read.Rows)
{
    Console.WriteLine($"{row[0].Text}, {row[1].Number}, {row[2].Date:yyyy-MM-dd}");
}
```

## 单元格类型

| 类型 | 说明 |
|---|---|
| `Text` | 文本（共享串/内联串） |
| `Number` | 数字（整数/小数，12 位以内精确） |
| `Date` | 日期（自动识别 Excel 日期格式，支持 1900/1904 系统） |
| `Boolean` | 布尔值 |
| `Empty` | 空单元格 |
| 公式结果值 | 读 `<f>` 公式字符串 + `<v>` 缓存值；写公式字符串不计算 |

## API 速查

### 对象模型 API（推荐）

| 类型 / 方法 | 说明 |
|---|---|
| `Excel.Open(path)` | 按扩展名自动识别格式打开工作簿 |
| `Excel.Open(stream, format)` | 从流打开工作簿（必须显式指定格式） |
| `Excel.Create(format)` / `Excel.Create(sheetName, format)` / `Excel.Create(sheetNames[], format)` | 新建工作簿（xlsx/xlsm/csv/xls/xlsb），支持批量添加工作表 |
| `Excel.Create<T>(data, sheetName, format, configure?)` / `Excel.Create(dataTable, ...)` | 新建工作簿并直接写入 List\<T\> / DataTable 数据（首行表头，AOT 安全） |
| `Excel.Read<T>(path, sheetName?)` | 读为 List\<T\>（反射，AOT 安全） |
| `Excel.ReadAsDataTable(path, sheetName?)` | 读为 DataTable（AOT 安全） |
| `Excel.Write(path, Workbook)` | 写出工作簿 |
| `Excel.Write(path, DataTable)` | 从 DataTable 写（AOT 安全） |
| `Excel.Write<T>(path, IEnumerable<T>)` | 从 List\<T\> 写（反射，AOT 安全） |
| `Excel.CreateWriter(path/stream)` | 流式写入大文件（逐行） |
| `Excel.StreamRows(path, sheetName, onRow)` | 流式读取大文件 |
| `Excel.GetSheetNames(path)` | 列出所有工作表名 |
| `Workbook.Worksheets / Properties / Save() / SaveAs(path, format)` | 文件级操作 |
| `Worksheet.Cell("A1") / Cell(row, col) / Range("A1:D10") / Cells` | 表级访问 |
| `Worksheet.ImportData<T>(data, configure?)` / `ImportData(dataTable)` | 清空整表并从 A1 重建（List\<T\> / DataTable） |
| `Worksheet.SetValue / Merge / Unmerge` | 表级写值与合并 |
| `Cell.GetString/GetDouble/GetDateTime/GetBoolean/TryGet*` | 单元格便利取值 |
| `Cell.SetValue / IsFormula / FromFormula` | 单元格写值与公式 |
| `Cell.Style / Cell.NumberFormat` | 指定单元格样式与数字格式（背景色/字体/对齐等） |
| `Worksheet.Cell("A2").Style = ...` / `Worksheet.Range("A2:C3").Style = ...` | 改单个/区域单元格样式 |
| `Worksheet.Comments["A2"] = "备注"` | 给指定单元格加批注 |
| `Cells[row, col] / Cells["A1"] / Cells.Range(...)` | 集合式访问 |
| `ExcelRange.Fill / Clear / Style / Merge / ToValues` | 区域操作 |

### XlsxWriter

| 方法 | 说明 |
|---|---|
| `Write(path, SheetData)` / `Write(stream, SheetData)` | 写单表 |
| `Write(path, IReadOnlyList<SheetData>)` / `Write(stream, ...)` | 写多表 |
| `Write(path/stream, sheets, WorkbookProperties)` | 写出并携带文档属性 |
| `Write(path, DataTable)` | 从 DataTable 写（AOT 安全） |
| `Write<T>(path, IEnumerable<T>)` | 从 List\<T\> 写（反射，AOT 安全） |
| `Append(path, SheetData, WorkbookProperties?)` | 追加数据并可更新文档属性 |
| `AutoColumnWidths(sheet)` | 自动估算列宽 |

### XlsxReader

| 方法 | 说明 |
|---|---|
| `GetSheetNames(path)` / `GetSheetNames(stream)` | 列出所有工作表名 |
| `Read(path, sheetIndex)` / `Read(stream, sheetIndex)` | 按索引读单表 |
| `Read(path, sheetName)` / `Read(stream, sheetName)` | 按名称读单表 |
| `ReadAll(path)` / `ReadAll(stream)` | 读取所有工作表 |
| `StreamRows(path, sheetName, onRow)` / `StreamRows(stream, ...)` | 流式读大文件 |
| `ReadWithProgress(path, sheetIndex, onProgress)` | 带进度读取 |
| `ReadAsDataTable(path)` / `ReadAsDataTable(stream)` | 读为 DataTable（AOT 安全） |
| `Read<T>(path)` | 读为 List\<T\>（反射，AOT 安全） |
| `ReadProperties(path)` / `ReadProperties(stream)` | 读取文档属性（作者/时间/标题） |

### 模型类

| 类 | 说明 |
|---|---|
| `Workbook` / `Worksheet` / `Cells` / `ExcelRange` | 对象模型 |
| `ExcelFormat` / `ExcelReadOptions` / `ExcelWriteOptions` | 格式与选项 |
| `SheetData` | 工作表数据（表头/行/样式/合并/筛选/行高/批注/验证） |
| `Cell` | 单元格 |
| `CellStyle` / `BorderStyle` / `BorderEdge` | 样式 |
| `CellRange` | 区域（合并单元格） |
| `AutoFilter` / `FilterColumn` | 自动筛选 |
| `DataValidation` | 数据验证 |
| `Hyperlink` / `Cell.Hyperlink` | 单元格超链接（外部 URL / 内部跳转） |
| `WorksheetImage` / `ImagePlacement` / `ImageAnchor` / `ImageMoveMode` | 图片（浮动 / 单元格内嵌，锚点 / 移动方式 / AltText） |
| `WorkbookSecurity` / `Workbook.Security` | 文件级安全（打开密码 / 修改密码 / 只读状态） |
| `WorkbookProperties` | 文档属性（作者/时间/标题/应用名） |
| `LiteExcelException` / `InvalidSheetNameException` | 异常 |
| `LiteColumnAttribute` | List\<T\> 列特性 |
| `WriteOptions<T>` / `ReadOptions<T>` | List\<T\> 配置 |

## 目标框架

- net48
- net8.0

## AOT 兼容性

| API | AOT |
|---|---|
| `Excel.Open` / `Workbook` / `Worksheet` / `Cell` / `ExcelRange` / `Cells` | ✅ 安全（无反射） |
| `Excel.ReadAsDataTable` / `DataTable` 读写 | ✅ 安全（无反射） |
| `SheetData` / `XlsxWriter` / `XlsxReader` 读写 | ✅ 安全（无反射） |
| `Excel.Read<T>` / `Excel.Write<T>` / `List<T>` 读写 | ✅ 安全（反射映射已标注 `[DynamicallyAccessedMembers]`） |

## 详细文档

- 📖 [使用手册](docs/USAGE.zh-CN.md) — 完整 API 参考与全部功能示例（样式/合并/筛选/批注/数据验证/Stream 等）
- 📝 [更新日志](docs/CHANGELOG.md) — 版本变更记录
- 🌐 [English README](README.en.md)
