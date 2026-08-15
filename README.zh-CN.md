# LiteExcel

轻量级 xlsx/xlsm/csv 读写库，零第三方依赖，AOT 友好。**版本 2.2.0**

> [English README](README.en.md)

## 特性

- **零依赖**：仅用 .NET BCL（ZipArchive + XmlReader/XDocument），无任何第三方包
- **AOT 友好**：低层 API 无反射；高层反射 API 标注 `[RequiresUnreferencedCode]`
- **双目标**：net48 + net8.0（老 WinForms 项目与新项目都能用）
- **直觉化高层 API**：`Excel -> Workbook -> Worksheet -> Cell/Range/Cells` 自然层级，一行式读写
- **格式可扩展**：xlsx/xlsm/csv 已支持；xlsb/xls 预留后端
- **全功能**：读/写、样式、合并单元格、自动筛选、行高/列宽、批注、数据验证、追加、Stream、List\<T\>/DataTable 便利 API、流式读写大文件
- **真实文件兼容**：可正确读取 Excel/WPS 创建的 xlsx（含 Table/theme 等扩展部件）

## 安装

```powershell
dotnet add package LiteExcel
```

或本地 .nupkg 引用：

```xml
<PackageReference Include="LiteExcel" Version="2.2.0" />
```

## 快速上手（推荐：高层 API）

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

## 低层 API（兼容保留）

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

### 高层 API（推荐）

| 类型 / 方法 | 说明 |
|---|---|
| `Excel.Open(path)` | 按扩展名自动识别格式打开工作簿 |
| `Excel.Create(format)` / `Excel.Create(sheetName, format)` | 新建工作簿（xlsx/xlsm/csv） |
| `Excel.Read<T>(path, sheetName?)` | 读为 List\<T\>（反射，不兼容 AOT） |
| `Excel.ReadAsDataTable(path, sheetName?)` | 读为 DataTable（AOT 安全） |
| `Excel.Write(path, Workbook)` | 写出工作簿 |
| `Excel.Write(path, DataTable)` | 从 DataTable 写（AOT 安全） |
| `Excel.Write<T>(path, IEnumerable<T>)` | 从 List\<T\> 写（反射，不兼容 AOT） |
| `Excel.CreateWriter(path/stream)` | 流式写入大文件（逐行） |
| `Excel.StreamRows(path, sheetName, onRow)` | 流式读取大文件 |
| `Excel.GetSheetNames(path)` | 列出所有工作表名 |
| `Workbook.Worksheets / Properties / Save() / SaveAs(path, format)` | 文件级操作 |
| `Worksheet.Cell("A1") / Cell(row, col) / Range("A1:D10") / Cells` | 表级访问 |
| `Worksheet.SetValue / Merge / Unmerge` | 表级写值与合并 |
| `Cell.GetString/GetDouble/GetDateTime/GetBoolean/TryGet*` | 单元格便利取值 |
| `Cell.SetValue / IsFormula / FromFormula` | 单元格写值与公式 |
| `Cells[row, col] / Cells["A1"] / Cells.Range(...)` | 集合式访问 |
| `ExcelRange.Fill / Clear / Style / Merge / ToValues` | 区域操作 |

### XlsxWriter（低层兼容）

| 方法 | 说明 |
|---|---|
| `Write(path, SheetData)` / `Write(stream, SheetData)` | 写单表 |
| `Write(path, IReadOnlyList<SheetData>)` / `Write(stream, ...)` | 写多表 |
| `Write(path/stream, sheets, WorkbookProperties)` | 写出并携带文档属性 |
| `Write(path, DataTable)` | 从 DataTable 写（AOT 安全） |
| `Write<T>(path, IEnumerable<T>)` | 从 List\<T\> 写（反射，不兼容 AOT） |
| `Append(path, SheetData, WorkbookProperties?)` | 追加数据并可更新文档属性 |
| `AutoColumnWidths(sheet)` | 自动估算列宽 |

### XlsxReader（低层兼容）

| 方法 | 说明 |
|---|---|
| `GetSheetNames(path)` / `GetSheetNames(stream)` | 列出所有工作表名 |
| `Read(path, sheetIndex)` / `Read(stream, sheetIndex)` | 按索引读单表 |
| `Read(path, sheetName)` / `Read(stream, sheetName)` | 按名称读单表 |
| `ReadAll(path)` / `ReadAll(stream)` | 读取所有工作表 |
| `StreamRows(path, sheetName, onRow)` / `StreamRows(stream, ...)` | 流式读大文件 |
| `ReadWithProgress(path, sheetIndex, onProgress)` | 带进度读取 |
| `ReadAsDataTable(path)` / `ReadAsDataTable(stream)` | 读为 DataTable（AOT 安全） |
| `Read<T>(path)` | 读为 List\<T\>（反射，不兼容 AOT） |
| `ReadProperties(path)` / `ReadProperties(stream)` | 读取文档属性（作者/时间/标题） |

### 模型类

| 类 | 说明 |
|---|---|
| `Workbook` / `Worksheet` / `Cells` / `ExcelRange` | 高层模型 |
| `ExcelFormat` / `ExcelReadOptions` / `ExcelWriteOptions` | 格式与选项 |
| `SheetData` | 低层工作表数据（表头/行/样式/合并/筛选/行高/批注/验证） |
| `Cell` | 单元格 |
| `CellStyle` / `BorderStyle` / `BorderEdge` | 样式 |
| `CellRange` | 区域（合并单元格） |
| `AutoFilter` / `FilterColumn` | 自动筛选 |
| `DataValidation` | 数据验证 |
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
| `SheetData` / 低层读写 | ✅ 安全（无反射） |
| `Excel.Read<T>` / `Excel.Write<T>` / `List<T>` 读写 | ⚠️ 标注 `[RequiresUnreferencedCode]`，AOT 编译时提示警告 |

## 详细文档

- 📖 [使用手册](docs/USAGE.zh-CN.md) — 完整 API 参考与全部功能示例（样式/合并/筛选/批注/数据验证/Stream 等）
- 📝 [更新日志](docs/CHANGELOG.md) — 版本变更记录
- 🌐 [English README](README.en.md)
