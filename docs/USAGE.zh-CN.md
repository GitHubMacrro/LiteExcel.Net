# LiteExcel 使用手册

**版本**：2.4.2  
**目标框架**：net48 + net8.0  
**依赖**：零第三方依赖，仅用 .NET BCL

---

## 目录

1. [安装与引用](#1-安装与引用)
2. [对象模型 API（推荐）](#2-对象模型-api推荐)
3. [快速上手](#3-快速上手)
4. [单元格与数据类型](#4-单元格与数据类型)
5. [读取](#5-读取)
6. [写出](#6-写出)
7. [样式](#7-样式)
8. [合并单元格](#8-合并单元格)
9. [自动筛选](#9-自动筛选)
10. [行高与列宽](#10-行高与列宽)
11. [单元格批注](#11-单元格批注)
12. [超链接](#12-超链接)
13. [冻结窗格](#13-冻结窗格)
14. [图片](#14-图片)
15. [数据验证（下拉列表）](#15-数据验证下拉列表)
16. [追加数据](#16-追加数据)
17. [List&lt;T&gt; 映射（反射，AOT 安全）](#17-listt-映射反射aot-安全)
18. [DataTable 便利 API（AOT 安全）](#18-datatable-便利-apiaot-安全)
19. [Stream 读写](#19-stream-读写)
20. [流式读取与进度回调](#20-流式读取与进度回调)
21. [文档属性（作者/时间/标题）](#21-文档属性作者时间标题)
22. [文件级安全（打开密码 / 修改密码）](#22-文件级安全打开密码--修改密码)
23. [错误处理](#23-错误处理)
24. [AOT 兼容性](#24-aot-兼容性)
25. [条件格式（2.4.3+）](#25-条件格式243)
26. [命名区域（Name Range）](#26-命名区域-name-range)
27. [能力降级回调（OnDegradation）（2.4.2+）](#27-能力降级回调ondegradation242)
28. [完整 API 索引](#28-完整-api-索引)

---
## 1. 安装与引用

### NuGet 安装

```powershell
dotnet add package LiteExcel
```

### csproj 直接引用

```xml
<ItemGroup>
  <PackageReference Include="LiteExcel" Version="2.4.2" />
</ItemGroup>
```

### 命名空间

所有 API 在 `LiteExcel` 命名空间下：

```csharp
using LiteExcel;
```

### 目标框架

- **net48**：老 WinForms 项目可直接引用
- **net8.0**：新项目可用，支持 AOT
- C# 12 语法（`<LangVersion>latest</LangVersion>`）

---

## 2. 对象模型 API（推荐）

从 `2.2.0` 起提供直觉化的对象模型 API，自然层级：

```text
Excel             统一门面（打开 / 新建 / 便利读写 / 流式）
  -> Workbook     工作簿（工作表集合 / 文档属性 / 保存）
      -> Worksheet 工作表（单元格 / 区域 / 合并 / 样式）
          -> Cells / Cell / ExcelRange
```

对象模型 API 基于同一套读写引擎，但把坐标、取值、保存等细节封装成更接近 Excel 的习惯用法。`XlsxReader / XlsxWriter / SheetData / Cell` 继续保留，新旧 API 可以混用，写出的文件互相兼容。

### 2.1 新建工作簿

```csharp
using LiteExcel;

// 默认 xlsx，自带一张名为 "Sheet1" 的工作表
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
```

指定格式与初始表名：

```csharp
var wbCsv = Excel.Create(ExcelFormat.Csv);              // 新建 csv 工作簿
var wb2   = Excel.Create("员工表", ExcelFormat.Xlsx);   // 新建并命名首表
```

支持格式：`Xlsx`、`Xlsm`、`Csv`、`Xls`、`Xlsb`（后两者为历史兼容格式，读取与写入均已实现）。

带数据新建（一步建簿并写入首表，返回的 `Workbook` 可与样式/冻结/条件格式/密码等能力混用）：

```csharp
// 从 List<T> 新建（首行为表头，反射映射，AOT 安全）
var wb = Excel.Create(new[] { new Person { Name = "张三", Age = 25 } }, "员工表");
wb.Worksheets[0].FreezeRows = 1;      // 建完后可继续用工作簿级能力
wb.SaveAs("out.xlsx");

// 从 DataTable 新建（表名为空时取 DataTable.TableName，再为空则 Sheet1）
var dt = new DataTable("库存表");
dt.Columns.Add("货号"); dt.Columns.Add("金额");
dt.Rows.Add("A1", 12.5m);
var wb2 = Excel.Create(dt);
wb2.SaveAs("stock.xlsx");
```

### 2.2 打开已有文件

```csharp
// 按扩展名自动识别格式（.xlsx / .xlsm / .csv）
var opened = Excel.Open("data.xlsx");

// 或显式指定格式
var forced = Excel.Open("data.csv", ExcelFormat.Csv);
```

`Excel.DetectFormat(path)` 可单独判断文件格式。

打开加密文件需提供密码：

```csharp
var opened = Excel.Open("encrypted.xlsx", new ExcelReadOptions
{
    OpenPassword = "1",        // 打开密码
    ModifyPassword = "12",     // 修改密码（可选）
});
```

> 文件级安全（打开/修改密码）的完整说明见 [§22 文件级安全](#22-文件级安全打开密码--修改密码)。

### 2.3 读写单元格

坐标为 **1-based**，即 `Cell(1, 1)` 对应 `A1`；也可以直接使用 A1 地址。

```csharp
var ws = wb.Worksheets["Sheet1"];

// 写入
ws.SetValue("A1", "姓名");
ws.SetValue("B1", "年龄");
ws.SetValue(2, 1, "张三");   // A2
ws.SetValue(2, 2, 25);       // B2

// 读取
string name = ws.Cell("A2").GetString();   // "张三"
double age = ws.Cell(2, 2).GetDouble();    // 25
```

取值方法：

| 方法 | 说明 |
|---|---|
| `GetString()` / `TryGetString(out var v)` | 文本取值 |
| `GetDouble()` / `TryGetDouble(out var v)` | 数值取值 |
| `GetDateTime()` / `TryGetDateTime(out var v)` | 日期取值 |
| `GetBoolean()` / `TryGetBoolean(out var v)` | 布尔取值 |
| `GetValue()` | 按类型返回 `object?` |
| `SetValue(object? value)` | 写值（字符串/数值/日期/布尔/公式） |

`Cell.Style`、`Cell.NumberFormat` 可读写单元格样式与数字格式；`Cell.IsFormula` 判断是否为公式。

#### 串联场景：改值 + 改样式 + 加备注 + 保存

打开已有文件后，可对**指定单元格**同时做取值、改样式、加备注操作，再保存：

```csharp
var wb = Excel.Open("report.xlsx");
var ws = wb.Worksheets["Sheet1"];

// 改 A2 的值
ws.Cell("A2").SetValue("已完成");

// 改 A2 的背景色与字体
ws.Cell("A2").Style = new CellStyle
{
    FillColor = "#FFFF00",              // 背景色
    FontName  = "微软雅黑",             // 字体
    FontSize  = 14,
    Bold      = true,
    FontColor = "#FF0000",              // 字体颜色
};

// 给 A2 加备注
ws.Comments ??= new();
ws.Comments["A2"] = "本单元格需要人工复核";

// 区域批量样式（A2:C3 每个单元格都应用）
ws.Range("A2:C3").Style = new CellStyle { FillColor = "#D9E1F2", Italic = true };

wb.Save();                              // 覆盖保存原文件
```

> 说明：样式/批注的具体 API 见 [§7 样式](#7-样式) 与 [§11 单元格批注](#11-单元格批注)。

### 2.4 集合式访问 `Cells`

```csharp
var cells = ws.Cells;

var a1 = cells[1, 1];            // 1-based 坐标
var b2 = cells["B2"];            // A1 地址
var r  = cells.Range("A1:C10");  // 区域

foreach (var cell in cells)      // 枚举所有已存储单元格
{
    Console.WriteLine($"{cell.Text} / {cell.Number}");
}

cells.SetValue("D2", "备注");
cells.Clear();                   // 清空全部
```

### 2.5 区域操作 `ExcelRange`

`Worksheet.Range(...)` 返回 `ExcelRange`（注意：类名为 `ExcelRange`，不是 `Range`，避免与 BCL 的 `System.Range` 冲突）：

```csharp
var range = ws.Range("A1:C3");            // 或 ws.Range(1, 1, 3, 3)
range.Fill(0);                            // 整片填充
range.Fill(new object?[,] { { 1, 2, 3 }, { 4, 5, 6 } });  // 填充矩阵
var values = range.ToValues();            // object?[,]
var cells = range.ToCells();              // Cell[,]
range.Clear();
range.Merge();                            // 合并该区域
range.Unmerge();
range.Style = new CellStyle { Bold = true };  // 区域样式

foreach (var cell in range) { /* 枚举区域内单元格 */ }
```

### 2.6 保存与另存

```csharp
wb.Save();                        // 保存到打开/新建时的路径；无路径则抛 LiteExcelException
wb.SaveAs("output.xlsx");         // 另存为，更新当前路径
wb.SaveAs("output.csv", ExcelFormat.Csv);  // 跨格式另存（取决于后端能力）
wb.Save(stream, ExcelFormat.Xlsx);         // 写 Stream
```

规则：

- `SaveAs` 成功后，后续 `Save()` 保存到新路径。
- 跨格式保存是否成功取决于目标格式后端；`csv` 仅支持单表数据。
- `Excel.Write(path, workbook)` 相当于“另存到指定路径”。

### 2.7 工作表管理

```csharp
var wb = Excel.Create();
wb.Worksheets.Add("Sheet2");              // 新增
wb.Worksheets.Add("Sheet3");
wb.Worksheets.Move(0, 1);                 // 移动
wb.Worksheets.Remove("Sheet2");           // 删除
wb.Worksheets.RemoveAt(0);
bool has = wb.Worksheets.Contains("Sheet3");
var names = wb.Worksheets.Names;          // ["Sheet3", ...]
```

### 2.8 文档属性

```csharp
var props = wb.Properties;
props.Creator = "LiteExcel";
props.Title = "示例报表";
wb.Save();
```

### 2.9 List\<T\> / DataTable 便利 API

```csharp
// List<T> 映射（反射，AOT 安全）
Excel.Write("out.xlsx", new[] { new Person { Name = "张三", Age = 25 } });
var list = Excel.Read<Person>("out.xlsx");

// DataTable（AOT 安全）
var dt = Excel.ReadAsDataTable("out.xlsx");
Excel.Write("out2.xlsx", dt);
```

`Excel.GetSheetNames(path)` 可列出所有工作表名。

### 2.10 流式读写大文件

```csharp
// 流式写出：逐行写入，不驻留内存
using (var writer = Excel.CreateWriter("large.xlsx"))
{
    writer.WriteRow(new object?[] { "姓名", "年龄" });
    for (int i = 0; i < 100000; i++)
        writer.WriteRow(new object?[] { $"用户{i}", i });
}

// 流式读取
Excel.StreamRows("large.xlsx", "Sheet1", row =>
{
    Console.WriteLine(row[0]?.Text);
});
```

### 2.11 公式与高级能力桥接

```csharp
ws.SetValue("A1", 1);
ws.SetValue("A2", 2);
ws.Cell("A3").SetValue(Cell.FromFormula("SUM(A1:A2)"));  // 写入公式字符串
bool isFormula = ws.Cell("A3").IsFormula;

ws.Merge("A1:B1");                    // 合并
ws.FreezeHeader = true;               // 冻结表头（等价于 FreezeRows = 1）
ws.Range("A1:B1").Style = new CellStyle { Bold = true };
```

样式、合并、批注、数据验证、自动筛选、行高列宽、超链接、冻结窗格等高级能力均可在 `Worksheet` 层直接使用，与 `SheetData` 能力一一对应。

### 2.12 格式支持矩阵

| 格式 | 读 | 写 | 说明 |
|---|---|---|---|
| `xlsx` | ✅ | ✅ | 完整读写；支持 1904 日期系统读/写；含宏工作簿写 `.xlsx` 会抛错（见下方"宏与降级"） |
| `xlsm` | ✅ | ✅ | 读写保存；宏部件 `vbaProject.bin` 与宿主 codeName 绑定（`workbookPr`/`sheetPr`）保存时保留 |
| `csv` | ✅ | ✅ | 仅表格数据，无样式/合并等 |
| `xls` | ✅ | ✅ | 读写（BIFF8，Excel 97+）；写入时公式降级为静态值；含宏工作簿写 `.xls` 会抛错（见下方"宏与降级"） |
| `xlsb` | ✅ | ✅ | 读写（BIFF12 二进制 OOXML）；公式写入按缓存值降级；支持 1904 日期系统读/写 |

> **xls 读取范围**：`Excel.Open("file.xls")` 可读取 BIFF8 工作簿的数据单元格（文本/数字/日期/布尔）、共享字符串（含跨 CONTINUE 续接）、合并单元格、列宽、行高、冻结表头。公式单元格返回缓存结果值，并解析公式文本（常见单元格引用、运算符与内置函数；数组/3D 引用等不支持的公式仅返回缓存值）。

> **xls 写入范围**：`wb.SaveAs("file.xls", ExcelFormat.Xls)` 可写出 BIFF8 工作簿，支持多工作表（中文名）、文本/数字/日期/布尔、合并单元格、列宽、行高、冻结表头、自定义数字格式。公式单元格按缓存结果值静态写出（公式文本不保留）。已用 Excel 打开验证。

> **xlsb 读取范围**：`Excel.Open("file.xlsb")` 可读取二进制 OOXML 变体的数据单元格（文本/数字/日期/布尔/错误）、共享字符串、合并单元格、列宽、行高、冻结表头、1904 日期系统。公式单元格返回缓存结果值，并解析公式文本。
>
> **xlsb 写入范围**：`wb.SaveAs("file.xlsb", ExcelFormat.Xlsb)` 可写出 BIFF12 工作簿，支持多工作表（中文名）、文本/数字/日期/布尔、共享字符串、数字格式、合并单元格、列宽、行高、冻结表头。公式单元格按缓存结果值静态写出（公式文本不保留）。已用 Excel 打开验证（无修复提示、另存后读取一致），SheetJS 交叉验证一致。

> **保存保真**：通过 `Excel.Open` 打开后修改再保存时，LiteExcel 会重建已映射的部件（工作表数据、样式、合并、批注、验证、筛选、公式等），并将**未映射的 OOXML 部件按原始字节保留**（如宏 `vbaProject.bin`、主题、绘图、图表、表格、外部链接等）。因此 `xlsm` 打开→修改→保存后宏不会丢失。此外，工作簿与工作表的 VBA 宿主代码名（`workbookPr@codeName` / `sheetPr@codeName`）也会在打开时捕获、保存时按 schema 位置写回，确保 VBA 工程中的模块绑定（`ThisWorkbook`、工作表模块、事件宏）不因宿主被重新命名而错位失效。
>
> **降级规则**：若打开后新增/删除/重命名/移动了工作表（结构发生变化），工作表级未映射关系（如绘图、超链接）不再复用到新文件，但这些部件的原始字节仍会保留为无害的未引用条目。工作簿级部件（宏、主题）不受结构变化影响。

> **1904 日期系统**：`Excel.Open` 打开 1904 日期系统的工作簿（`workbookPr@date1904` / `BrtWbProp` flags / `DATE1904` 记录）时，日期单元格会按 1904 基准（1904-01-01 = 序列 0）换算。`SaveAs` 到 xlsx/xlsb/xls 时会写回 1904 标志并保持序列一致，因此 1904 工作簿跨格式转换不会偏移 4 年。

> **加密文件识别**：带打开密码的 xlsx/xlsm/xlsb 实际是 OLE CFB 容器（内含 `EncryptionInfo`/`EncryptedPackage` 流）。2.4.0 起支持打开密码读取与密码保存（见 [§22 文件级安全](#22-文件级安全打开密码--修改密码)）；加密 `.xls`（BIFF8 `FILEPASS` 记录）会识别并抛明确异常。

> **宏与降级**：工作簿若含 VBA 宏（打开 xlsm/xlsb 捕获 `vbaProject.bin`），`SaveAs` 到 `.xlsx` 或 `.xls`（不支持宏）会抛 `LiteExcelException` 阻止静默丢失，请在创建文件前拦截。含宏工作簿请保存为 `.xlsm` 或 `.xlsb`。数据单元格的值在任意格式转换中都不会静默丢失；xls 不支持批注/数据验证/自动筛选等元数据时按文档化降级（忽略，不写错文件）。

> **Stream API**：`Excel.Open(Stream, format)` 支持五格式对象模型读取（必须显式指定格式，流无扩展名）；`Workbook.Save(Stream, format)` 支持五格式保存。输入流不会被关闭（由调用方管理）；支持不可定位的流（内部复制到内存）。`XlsxReader.StreamRows(Stream, ...)` 仍是 xlsx/xlsm 专用底层流式逐行读取。从 Stream 打开后 `Workbook.CurrentPath` 为 null，需用 `SaveAs` 指定保存路径。

### 2.13 新旧 API 对照

| 场景 | 对象模型 API | XlsxWriter / XlsxReader |
|---|---|---|
| 打开文件 | `Excel.Open(path)` | `XlsxReader.Read(path, 0)` |
| 新建/写出 | `Excel.Create()` + `SaveAs` | `XlsxWriter.Write(path, sheet)` |
| 读单表 | `Workbook.Worksheets[i]` | `XlsxReader.Read(path, i)` |
| 读为 List\<T\> | `Excel.Read<T>(path)` | `XlsxReader.Read<T>(path)` |
| 读为 DataTable | `Excel.ReadAsDataTable(path)` | `XlsxReader.ReadAsDataTable(path)` |
| 流式读 | `Excel.StreamRows(path, name, cb)` | `XlsxReader.StreamRows(path, name, cb)` |
| 流式写 | `Excel.CreateWriter(path)` | `XlsxStreamWriter` |

---
## 3. 快速上手

### 创建工作簿填充数据（高层 API）

交互式编辑推荐使用对象模型：`Excel.Create` 新建 → `SetValue` 填充 → `SaveAs` 保存 → `Open` 读回：

```csharp
using LiteExcel;

var wb = Excel.Create();                    // 默认 xlsx，自带 Sheet1
var ws = wb.Worksheets["Sheet1"];

ws.SetValue("A1", "姓名");  ws.SetValue("B1", "年龄");
ws.SetValue("A2", "张三");   ws.SetValue("B2", 28);
ws.SetValue("A3", "李四");   ws.SetValue("B3", 32);

ws.Cell("A1").Style = new CellStyle { Bold = true };   // 表头加粗
ws.FreezeHeader = true;                                  // 冻结首行

wb.SaveAs("output.xlsx");

var rb = Excel.Open("output.xlsx");        // 读回
Console.WriteLine(rb.Worksheets[0].Cell("A2").GetString());  // 张三
```

### 用自定义类快速写出（List&lt;T&gt; 映射）

批量导出用 `List&lt;T&gt;` 映射，一个类即可声明列名 / 顺序 / 格式 / 公式 / 忽略：

```csharp
public class Employee
{
    [LiteColumn(Name = "姓名", Order = 0)]
    public string Name { get; set; } = "";

    [LiteColumn(Name = "年龄", Order = 1)]
    public int Age { get; set; }

    [LiteColumn(Name = "入职日期", Order = 2, Format = "yyyy-MM-dd")]
    public DateTime HireDate { get; set; }

    [LiteColumn(Name = "薪资", Order = 3, Format = "#,##0.00")]
    public decimal Salary { get; set; }

    [LiteColumn(Name = "在职", Order = 4)]
    public bool Active { get; set; }

    [LiteColumn(Name = "平均年龄", Order = 5, IsFormula = true)]
    public string AvgAgeFormula { get; set; } = "";   // 公式列："=AVERAGE(B2:B3)" 或 "AVERAGE(B2:B3)"

    [LiteColumn(Ignore = true)]
    public string InternalId { get; set; } = "";       // 不输出
}

var list = new List<Employee>
{
    new() { Name = "张三", Age = 28, HireDate = new DateTime(2020, 3, 15), Salary = 8500.50m, Active = true, AvgAgeFormula = "=AVERAGE(B2:B3)" },
    new() { Name = "李四", Age = 32, HireDate = new DateTime(2018, 7, 1),  Salary = 12000m,    Active = false, AvgAgeFormula = "AVERAGE(B2:B3)" },
};

XlsxWriter.Write("employees.xlsx", list, opt => opt.FreezeHeader = true);
```

`List&lt;T&gt;` 支持的属性类型一览：

| C# 属性类型 | → Excel 列 | 特性 |
|---|---|---|
| `string` | 文本 | — |
| `int` / `double` / `decimal` | 数字 | `Format` |
| `DateTime` | 日期 | `Format`（显示格式） |
| `bool` | 布尔 | — |
| `string` + `IsFormula=true` | 公式 | `IsFormula` |
| `null` | 空单元格 | — |
| 任意 + `Ignore=true` | 不输出 | `Ignore` |

> 超链接、样式、合并、冻结、条件格式、图片等高级能力请用高层 API（`Excel.Create` + `Cell.Hyperlink` / `Cell.Style` 等）或低层 `SheetData`。`List&lt;T&gt;` 映射聚焦"批量纯数据"场景。

### 最简写出

```csharp
var sheet = new SheetData
{
    SheetName = "员工表",
    Headers = new() { "姓名", "年龄", "生日" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("张三"), Cell.FromNumber(25), Cell.FromDate(new DateTime(2000, 1, 1)) },
        new Cell[] { Cell.FromText("李四"), Cell.FromNumber(30), Cell.FromDate(new DateTime(1995, 5, 10)) },
    },
};

XlsxWriter.Write("output.xlsx", sheet);
```

### 最简读取

```csharp
var sheet = XlsxReader.Read("output.xlsx", sheetIndex: 0);

Console.WriteLine($"表名: {sheet.SheetName}");
Console.WriteLine($"表头: {string.Join(", ", sheet.Headers)}");
foreach (var row in sheet.Rows)
{
    foreach (var cell in row)
    {
        Console.Write($"{cell.Type} ");
    }
    Console.WriteLine();
}
```

### 读取时如何取值

```csharp
var sheet = XlsxReader.Read("output.xlsx", 0);

foreach (var row in sheet.Rows)
{
    string name = row[0].Type == CellType.Text ? row[0].Text! : "";
    int age = row[1].Type == CellType.Number ? (int)row[1].Number : 0;
    DateTime birthday = row[2].Type == CellType.Date ? row[2].Date : DateTime.MinValue;

    Console.WriteLine($"{name}, {age}, {birthday:yyyy-MM-dd}");
}
```

**判断单元格类型**：用 `cell.Type` 判断，再用对应字段取值。

| `cell.Type` | 有效字段 |
|---|---|
| `CellType.Text` | `cell.Text` |
| `CellType.Number` | `cell.Number` |
| `CellType.Date` | `cell.Date` |
| `CellType.Boolean` | `cell.Boolean` |
| `CellType.Empty` | （空单元格） |

---

## 4. 单元格与数据类型

### Cell 类

`Cell` 表示一个单元格，`Type` 属性决定哪个值字段有效。

```csharp
public sealed class Cell
{
    public CellType Type { get; set; }
    public string? Text { get; set; }          // CellType.Text 时有效
    public double Number { get; set; }           // CellType.Number 时有效
    public DateTime Date { get; set; }           // CellType.Date 时有效
    public bool Boolean { get; set; }            // CellType.Boolean 时有效
    public CellStyle? Style { get; set; }        // 单元格样式（可选）
    public string? NumberFormat { get; set; }    // 数字格式（写出用）
    public bool IsEmpty { get; }                 // 是否为空
}
```

### 创建单元格的工厂方法

```csharp
// 文本
var textCell = Cell.FromText("你好");
var emptyTextCell = Cell.FromText("");  // → CellType.Empty
var nullTextCell = Cell.FromText(null); // → CellType.Empty

// 数字（可选指定数字格式）
var numCell = Cell.FromNumber(3.14);
var moneyCell = Cell.FromNumber(9999.50, "#,##0.00");  // 千分位 + 两位小数

// 日期（可选指定格式，默认 "yyyy-MM-dd"）
var dateCell = Cell.FromDate(new DateTime(2024, 6, 1));
var dateCell2 = Cell.FromDate(DateTime.Now, "yyyy/MM/dd HH:mm:ss");

// 布尔
var boolCell = Cell.FromBoolean(true);

// 空单元格
var emptyCell = Cell.Empty;
```

### 支持的单元格类型

| 类型 | 说明 |
|---|---|
| `Text` | 文本（共享字符串或内联字符串） |
| `Number` | 数字（整数精确到 long，12 位以内无损） |
| `Date` | 日期（Excel 存为数字 + 格式码，库自动转换） |
| `Boolean` | 布尔值（`t="b"`） |
| `Empty` | 空单元格 |

### 数字格式速查

常用格式字符串：

| 格式 | 效果 |
|---|---|
| `"0"` | 整数 |
| `"0.00"` | 两位小数 |
| `"#,##0"` | 千分位整数 |
| `"#,##0.00"` | 千分位 + 两位小数 |
| `"0.00%"` | 百分比 |
| `"yyyy-MM-dd"` | 日期（默认） |
| `"yyyy/MM/dd"` | 日期 |
| `"HH:mm:ss"` | 时间 |
| `"yyyy-MM-dd HH:mm:ss"` | 日期时间 |

### 读取时日期自动识别

读取时，库会查 `styles.xml` 的数字格式 ID（numFmtId），自动把日期格式的数字转成 `CellType.Date`。内置日期格式 ID（14-22, 27-36, 45-47, 50-58）会被识别为日期。

---

## 5. 读取

### 列出所有工作表名

```csharp
var names = XlsxReader.GetSheetNames("file.xlsx");
// ["员工表", "工资表", "部门表"]
```

### 按索引读取单表

```csharp
// sheetIndex 从 0 开始，firstRowIsHeader 默认 true
var sheet = XlsxReader.Read("file.xlsx", sheetIndex: 0);

// 不把第一行当表头
var sheet = XlsxReader.Read("file.xlsx", sheetIndex: 0, firstRowIsHeader: false);
```

### 按名称读取单表

```csharp
var sheet = XlsxReader.Read("file.xlsx", sheetName: "员工表");
```

### 读取所有工作表

```csharp
var allSheets = XlsxReader.ReadAll("file.xlsx");
foreach (var sheet in allSheets)
{
    Console.WriteLine($"{sheet.SheetName}: {sheet.Rows.Count} 行");
}
```

### 流式读取大文件（不驻留内存）

```csharp
XlsxReader.StreamRows("bigfile.xlsx", "Sheet1", row =>
{
    // row 是 IReadOnlyList<Cell>，逐行回调
    foreach (var cell in row)
    {
        // 处理单元格
    }
});
```

> **注意**：`StreamRows` 会自动跳过表头行（第一行）。只处理数据行。

### 带进度的读取

```csharp
XlsxReader.ReadWithProgress("bigfile.xlsx", sheetIndex: 0, (current, total) =>
{
    Console.WriteLine($"进度: {current}/{total} ({current * 100 / total}%)");
});
```

> `current` 从 1 递增到 `total`（数据行数，不含表头）。

### 读取结果的 SheetData 结构

```csharp
public sealed class SheetData
{
    public string SheetName { get; set; }            // 表名
    public List<string> Headers { get; set; }         // 表头（firstRowIsHeader=true 时填充）
    public List<IReadOnlyList<Cell>> Rows { get; set; } // 数据行
    public List<CellRange> MergedRanges { get; set; }   // 合并单元格区域
    public AutoFilter? Filter { get; set; }              // 自动筛选
    public CellStyle? HeaderStyle { get; set; }          // 表头样式（读出时可能为空）
    public CellStyle? DefaultStyle { get; set; }         // 全表默认样式
    public Dictionary<int, CellStyle>? RowStyles { get; set; }      // 行级样式
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }  // 列级样式
    public Dictionary<int, double>? RowHeights { get; set; }       // 行高
    public Dictionary<string, string>? Comments { get; set; }      // 单元格批注
    public List<DataValidation>? Validations { get; set; }          // 数据验证
}
```

---

## 6. 写出

### 写单表

```csharp
var sheet = new SheetData { ... };
XlsxWriter.Write("output.xlsx", sheet);
```

### 写多表

```csharp
var sheets = new List<SheetData>
{
    new() { SheetName = "表1", Headers = new() { "A" }, Rows = ... },
    new() { SheetName = "表2", Headers = new() { "B" }, Rows = ... },
};
XlsxWriter.Write("multi.xlsx", sheets);
```

### 冻结表头

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    FreezeHeader = true,  // 第一行冻结（等价于 FreezeRows = 1）
};
```

### 任意行列冻结

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C" },
    Rows = ...,
    FreezeRows = 2,       // 冻结前 2 行
    FreezeColumns = 1,    // 冻结第 1 列
};
```

### 列宽设置

```csharp
var sheet = new SheetData
{
    Headers = new() { "姓名", "年龄", "备注" },
    Rows = ...,
    ColumnWidths = new() { 15, 8, 30 },  // 列宽（Excel 字符宽度单位）
};
```

### 列宽自适应

```csharp
var sheet = new SheetData { ... };

// 自动估算列宽（中文字符算 2，英文/数字算 1，范围 8~50）
XlsxWriter.AutoColumnWidths(sheet);

XlsxWriter.Write("output.xlsx", sheet);
```

> 在 `Write` 之前调用 `AutoColumnWidths`，它会填充 `sheet.ColumnWidths`。

### Sheet 名校验

写出时自动校验 Sheet 名，不合法会抛 `InvalidSheetNameException`：
- 不能为空
- 不超过 31 字符
- 不能包含 `\ / ? * [ ] :` 字符

```csharp
try
{
    XlsxWriter.Write("bad.xlsx", new SheetData { SheetName = "test[1]" });
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine(ex.Message);
}
```

---

### CSV 写出分隔符（2.4.3+）

默认 CSV 写出使用逗号分隔。切换到分号或 Tab（国内中文 Excel 常见分号）：

```csharp
Excel.Write("out.csv", wb, new ExcelWriteOptions { Separator = ';' });
```

指定后在该工作簿生命周期内永久生效（若次存仍如需保持同一选项，需在同一次 `Excel.Write` 传参）。

> 仅 CSV 支持自定义分隔符。其他格式忽略该项。
> 其他 Excel 专有能力（批注/合并/筛选/图片/样式等）在 CSV 写出会被降级上报，见 [§25 能力降级回调](#25-capability-degradation-callback-ondegradation-242)。

---

## 7. 样式

### 样式优先级（覆盖式）

```
单元格 Cell.Style  >  行 RowStyles[row]  >  列 ColumnStyles[col]  >  全表 DefaultStyle
```

> **覆盖式**：单元格有自己的 Style 就完全用单元格的，不继承行/列/全表。没有单元格样式才看行级，行级没有才看列级，列级没有才用全表默认。

### 单元格样式

```csharp
var style = new CellStyle
{
    FontName = "微软雅黑",
    FontSize = 14,
    Bold = true,
    Italic = true,
    Underline = false,
    Strikeout = false,
    FontColor = "#FF0000",      // 红色字体
    FillColor = "#FFFF00",      // 黄底
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    WrapText = true,
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
        Bottom = new BorderEdge { Style = "thin", Color = "#000000" },
        Left = new BorderEdge { Style = "medium", Color = "#000000" },
        Right = new BorderEdge { Style = "medium", Color = "#000000" },
    },
};

var cell = new Cell { Type = CellType.Text, Text = "带样式", Style = style };
```

### 修改指定单元格 / 区域样式（对象模型 API）

打开已有文件后修改**指定单元格**或**指定区域**的样式，直接对 `Cell.Style` / `ExcelRange.Style` 赋值即可（`Style` 为覆盖式替换，未设置的字段保持 Excel 默认）：

```csharp
var wb = Excel.Open("styled.xlsx");
var ws = wb.Worksheets["Sheet1"];

// 单个单元格：改背景色 + 字体 + 字体颜色
ws.Cell("A2").Style = new CellStyle
{
    FillColor = "#FFFF00",          // 背景色（#RRGGBB）
    FontName  = "微软雅黑",         // 字体名
    FontSize  = 14,                 // 字号（磅）
    Bold      = true,
    FontColor = "#FF0000",          // 字体颜色
};

// 区域：A2:C3 内每个单元格统一应用
ws.Range("A2:C3").Style = new CellStyle
{
    FillColor = "#D9E1F2",
    Italic    = true,
    HorizontalAlignment = HorizontalAlignment.Center,
};

// 同时改值和样式
ws.Cell("B2").SetValue("新值");
ws.Cell("B2").Style = new CellStyle { FillColor = "#92D050", Bold = true };

wb.Save();   // 或 wb.SaveAs("styled2.xlsx")
```

> `ExcelRange.Style` 会遍历区域内所有单元格逐个应用；`Style` 是**整体替换**而非合并增量，需要"保留已有样式只改某一项"时应先读出再复制（`CellStyle` 提供 `Clone()`）。

### 表头样式

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    HeaderStyle = new CellStyle
    {
        Bold = true,
        FontColor = "#FFFFFF",
        FillColor = "#4472C4",
        HorizontalAlignment = HorizontalAlignment.Center,
    },
    Rows = ...,
};
```

> 表头样式优先级：`HeaderStyle > ColumnStyles > DefaultStyle`

### 全表默认样式

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    DefaultStyle = new CellStyle { FontName = "Arial", FontSize = 11 },
};
```

### 行级样式

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("行0"), Cell.FromText("x") },
        new Cell[] { Cell.FromText("行1"), Cell.FromText("y") },  // 黄底
    },
    RowStyles = new()
    {
        { 1, new CellStyle { FillColor = "#FFFF00" } },  // key=0-based 行索引
    },
};
```

### 列级样式

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    ColumnStyles = new()
    {
        { 1, new CellStyle { Italic = true, FontColor = "#0000FF" } },  // key=0-based 列索引
    },
};
```

### 颜色格式

所有颜色使用 `#RRGGBB` 格式（6 位十六进制），不含 alpha 通道。

### 边框样式

边框 `Style` 可选值（Excel 标准）：

| 值 | 说明 |
|---|---|
| `"thin"` | 细线 |
| `"medium"` | 中等 |
| `"thick"` | 粗 |
| `"dotted"` | 点线 |
| `"dashed"` | 虚线 |
| `"double"` | 双线 |
| `"none"` | 无 |

### 对齐方式

| 水平对齐 | 垂直对齐 |
|---|---|
| `HorizontalAlignment.General` | `VerticalAlignment.Top` |
| `HorizontalAlignment.Left` | `VerticalAlignment.Center` |
| `HorizontalAlignment.Center` | `VerticalAlignment.Bottom` |
| `HorizontalAlignment.Right` | |

---
## 8. 合并单元格

### 写出合并

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C" },
    Rows = ...,
    MergedRanges = new()
    {
        new CellRange(0, 0, 0, 2),  // 合并第一数据行的 A1:C1
        new CellRange(1, 3, 0, 0),  // 合并 A2:A4（跨 3 行）
    },
};
```

### CellRange 说明

```csharp
public sealed class CellRange
{
    public int FirstRow { get; set; }  // 0-based 行索引
    public int LastRow { get; set; }
    public int FirstCol { get; set; }  // 0-based 列索引
    public int LastCol { get; set; }
}
```

> 行索引对应 `Rows` 列表（不包含表头行）。列索引从 0 开始。

### 读取合并

读取时自动还原到 `sheet.MergedRanges`，无需额外调用。

```csharp
var sheet = XlsxReader.Read("file.xlsx", 0);
foreach (var range in sheet.MergedRanges)
{
    Console.WriteLine($"合并: 行{range.FirstRow}-{range.LastRow} 列{range.FirstCol}-{range.LastCol}");
}
```

---

## 9. 自动筛选

### 写出筛选（方式 1：传筛选条件）

库自动遍历 Rows 求值，计算哪些行该 hidden。

```csharp
var sheet = new SheetData
{
    Headers = new() { "姓名", "城市", "分数" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("张三"), Cell.FromText("北京"), Cell.FromNumber(85) },
        new Cell[] { Cell.FromText("李四"), Cell.FromText("上海"), Cell.FromNumber(72) },
        new Cell[] { Cell.FromText("王五"), Cell.FromText("北京"), Cell.FromNumber(90) },
    },
    Filter = new AutoFilter
    {
        Range = "A1:C4",  // 筛选范围（表头 + 数据行）
        Columns = new()
        {
            new FilterColumn
            {
                ColumnIndex = 1,                    // 第 2 列（城市）
                Type = FilterType.Equals,
                Values = new() { "北京" },         // 只显示北京
            },
        },
    },
};

XlsxWriter.Write("filtered.xlsx", sheet);
// 结果：李四那行被 hidden
```

### 写出筛选（方式 2：手动指定 hidden 行）

```csharp
sheet.Filter = new AutoFilter
{
    Range = "A1:C4",
    HiddenRows = new() { 1 },  // 0-based 行索引，第 2 行 hidden
};
```

### 筛选条件类型

| `FilterType` | 说明 | `Values` | `Operator`/`MinValue`/`MaxValue` |
|---|---|---|---|
| `Equals` | 等于（多选） | 候选值列表 | 不用 |
| `Compare` | 比较/区间 | 比较值 | `Operator` 必填 |
| `Contains` | 文本包含 | 子串列表 | 不用 |
| `BeginsWith` | 以...开头 | 前缀列表 | 不用 |
| `EndsWith` | 以...结尾 | 后缀列表 | 不用 |
| `Blank` | 空白/非空白 | 空=空白，有值=非空白 | 不用 |

### Compare 操作符

| `FilterOperator` | 说明 |
|---|---|
| `GreaterThan` | 大于 |
| `GreaterThanOrEqual` | 大于等于 |
| `LessThan` | 小于 |
| `LessThanOrEqual` | 小于等于 |
| `Between` | 区间（需设 `MinValue` 和 `MaxValue`） |

### Between 示例

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.Between,
    MinValue = "60",   // 下限
    MaxValue = "90",   // 上限
},
```

### 多条件（AND 逻辑）

所有列的条件必须同时满足（AND），该行才显示。

```csharp
Columns = new()
{
    new FilterColumn { ColumnIndex = 0, Type = FilterType.Equals, Values = new() { "张三" } },
    new FilterColumn { ColumnIndex = 1, Type = FilterType.Equals, Values = new() { "北京" } },
},
// 只有 张三+北京 的行显示
```

### 读取筛选

```csharp
var sheet = XlsxReader.Read("filtered.xlsx", 0);
if (sheet.Filter is not null)
{
    Console.WriteLine($"范围: {sheet.Filter.Range}");
    Console.WriteLine($"Hidden 行: {string.Join(", ", sheet.Filter.HiddenRows)}");
}
```

---

## 10. 行高与列宽

### 设置行高

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
        new Cell[] { Cell.FromText("z"), Cell.FromText("w") },
    },
    RowHeights = new()
    {
        { 0, 30.0 },   // 第一数据行高 30 磅
        { 1, 14.25 },  // 第二数据行高 14.25 磅
    },
};
```

> `RowHeights` 的 key 是 0-based 数据行索引（对应 `Rows` 列表，不含表头行）。

### 列宽自适应

```csharp
var sheet = new SheetData { ... };

XlsxWriter.AutoColumnWidths(sheet);
XlsxWriter.Write("output.xlsx", sheet);
```

估算规则：
- 中文字符宽度算 2，英文/数字算 1
- 最小 8，最大 50
- 含表头行宽度比较

---

## 11. 单元格批注

### 写出批注

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("x"), Cell.FromText("y") },
    },
    Comments = new()
    {
        { "A1", "这是 A1 的批注" },
        { "B1", "这是 B1 的批注 <含特殊字符>" },
    },
};

XlsxWriter.Write("comments.xlsx", sheet);
```

> key 是 A1 格式的单元格引用（`列字母 + 行号`，如 `A1`、`B2`、`AA10`）。

### 对象模型 API：给指定单元格加/读批注

打开已有文件后，对指定单元格添加、修改、读取批注：

```csharp
var wb = Excel.Open("comments.xlsx");
var ws = wb.Worksheets["Sheet1"];

// 加批注（Comments 为 null 时先初始化）
ws.Comments ??= new();
ws.Comments["A2"] = "这是 A2 的批注";

// 修改已存在的批注
ws.Comments["A2"] = "批注已更新";

// 读取批注
if (ws.Comments is not null && ws.Comments.TryGetValue("A2", out var text))
    Console.WriteLine($"A2 批注：{text}");

// 删除批注
ws.Comments.Remove("A2");

wb.Save();
```

### 读取批注

```csharp
var sheet = XlsxReader.Read("comments.xlsx", 0);
if (sheet.Comments is not null)
{
    foreach (var (ref, text) in sheet.Comments)
    {
        Console.WriteLine($"{ref}: {text}");
    }
}
```

---
## 12. 超链接

### 写出超链接

超链接支持 xlsx/xlsm/xlsb/xls 四格式。通过 `Cell.Hyperlink` 属性设置：

```csharp
var sheet = new SheetData
{
    Headers = new() { "姓名", "主页" },
    Rows = new()
    {
        new Cell[]
        {
            Cell.FromText("张三"),
            new Cell { Type = CellType.Text, Text = "点击访问", Hyperlink = new Hyperlink
            {
                Target = "https://example.com",
                Tooltip = "张三的个人主页",
            }},
        },
    },
};

XlsxWriter.Write("links.xlsx", sheet);
```

### Hyperlink 属性

| 属性 | 类型 | 说明 |
|---|---|---|
| `Target` | `string` | 链接目标 URL（必填） |
| `Tooltip` | `string?` | 悬停提示文本（可选） |
| `IsInternal` | `bool` | 是否为内部链接（如 `Sheet1!A1`） |

### 对象模型 API：设置超链接

```csharp
var wb = Excel.Open("links.xlsx");
var ws = wb.Worksheets["Sheet1"];

// 设置超链接
ws.Cell("B2").Hyperlink = new Hyperlink
{
    Target = "https://example.com",
    Tooltip = "点击访问",
};

wb.Save();
```

### 读取超链接

```csharp
var sheet = XlsxReader.Read("links.xlsx", 0);
var cell = sheet.Rows[0][1];  // B2
if (cell.Hyperlink is not null)
{
    Console.WriteLine($"目标: {cell.Hyperlink.Target}");
    Console.WriteLine($"提示: {cell.Hyperlink.Tooltip}");
}
```

> 超链接支持 xlsx/xlsm/xlsb/xls 四格式读写（外部 URL/文件/mailto/UNC 与内部 `#Sheet!A1` 跳转）；csv 不支持超链接。内部链接 `IsInternal=true` 时 Target 形如 `#Sheet1!A1`。

---

## 13. 冻结窗格

### 设置冻结行/列

通过 `SheetData.FreezeRows` 和 `FreezeColumns` 属性控制：

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B", "C", "D" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("数据1"), Cell.FromText("x"), Cell.FromText("y"), Cell.FromText("z") },
        new Cell[] { Cell.FromText("数据2"), Cell.FromText("a"), Cell.FromText("b"), Cell.FromText("c") },
    },
    FreezeRows = 2,       // 冻结前 2 行
    FreezeColumns = 1,    // 冻结第 1 列
};
```

### FreezeHeader 兼容

`FreezeHeader = true` 等价于 `FreezeRows = 1`：

```csharp
var sheet = new SheetData
{
    Headers = new() { "A", "B" },
    Rows = ...,
    FreezeHeader = true,   // 等价于 FreezeRows = 1
};
```

### 对象模型 API

```csharp
var wb = Excel.Open("report.xlsx");
var ws = wb.Worksheets["Sheet1"];

// 冻结前 2 行
ws.FreezeRows = 2;

// 冻结第 1 列
ws.FreezeColumns = 1;

// 或用 FreezeHeader 兼容语法
ws.FreezeHeader = true;   // 等价于 FreezeRows = 1

wb.Save();
```

### 读取冻结

```csharp
var sheet = XlsxReader.Read("frozen.xlsx", 0);
Console.WriteLine($"冻结行数: {sheet.FreezeRows}");      // 0 表示未冻结
Console.WriteLine($"冻结列数: {sheet.FreezeColumns}");   // 0 表示未冻结
Console.WriteLine($"冻结表头: {sheet.FreezeHeader}");    // FreezeRows > 0 时 true
```

---

## 14. 图片

### 浮动图片

图片仅支持 xlsx/xlsm 格式。通过 `Worksheet.AddImage` 添加浮动图片：

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];

// 读取图片数据
byte[] imageData = File.ReadAllBytes("logo.png");

// 添加浮动图片到 A1 单元格位置
ws.AddImage(imageData, row: 1, column: 1, widthPx: 200, heightPx: 100);

wb.SaveAs("image.xlsx");
```

### 单元格内嵌图片（InCell）

```csharp
// 单元格内嵌模式（图片随单元格大小变化）
ws.AddImage(imageData, row: 1, column: 1,
    placement: ImagePlacement.InCell);
```

### ImagePlacement 枚举

| 值 | 说明 |
|---|---|
| `Floating` | 浮动图片，可指定位置和尺寸（默认） |
| `InCell` | 单元格内嵌，图片随单元格自适应 |

### 坐标与尺寸

- 坐标（row, column）为 **1-based**
- 宽高单位为像素，缺省时按图片原始尺寸
- 自动探测扩展名（png/jpg/gif/bmp）与像素尺寸

```csharp
// 省略宽高 → 使用原始图片尺寸
ws.AddImage(imageData, row: 2, column: 3, placement: ImagePlacement.Floating);

// 指定扩展名和名称
ws.AddImage(imageData, row: 1, column: 1,
    widthPx: 300, heightPx: 200,
    extension: "png", name: "产品图");
```

### 多 Sheet 混合使用

```csharp
var ws1 = wb.Worksheets["Sheet1"];
var ws2 = wb.Worksheets.Add("Sheet2");

ws1.AddImage(logoData, row: 1, column: 1, widthPx: 100, heightPx: 50);
ws2.AddImage(photoData, row: 3, column: 2, placement: ImagePlacement.InCell);
```

> 图片仅支持 xlsx/xlsm 格式。xls/xlsb/csv 不支持图片。InCell 嵌入图片在 Excel 中单元格显示为 `#VALUE!`（与 Excel 原生真实样本一致）。
> 读取：**浮动图片支持读回**（2.4.3+，xlsx/xlsm 打开后 `Images` 自动填充）；InCell 图片读回将在后续版本提供。

### 图片读回（2.4.3+）

打开含浮动图片的 xlsx/xlsm 后，`Worksheet.Images` 已回填：

```csharp
var wb = Excel.Open("hasImage.xlsx");
foreach (var img in wb.Worksheets[0].Images)
{
    Console.WriteLine($"位置: {img.Row},{img.Column}  类型: {img.Placement}  尺寸: {img.Anchor?.WidthPixels}x{img.Anchor?.HeightPixels}");
    Console.WriteLine($"altText: {img.AltText}  name: {img.Name}  bytes: {img.Data.Length}");
}
```

支持字段：Row / Column（1-based）、Placement（目前仅读回 Floating）、Name、AltText、Anchor（TopLeftCell/偏移/尺寸/MoveMode）、Extension（png/jpg/gif/bmp 自动）、Data（byte[]）。

InCell 图片读回后续版本（2.4.4）。写回保持原能力不清除已有图片。

### 图片锚点与移动方式（2.4.1+）

浮动图片支持高精度锚点，可指定左上单元格 + EMU 偏移 + 显示尺寸 + 随单元格的移动/缩放方式，以及无障碍替换文本（AltText）：

```csharp
ws.AddImage(logoData, new ImageAnchor
{
    TopLeftCell = "B2",           // 左上单元格 A1 引用
    TopLeftOffsetX = 9525,       // 水平偏移（EMU，1px≈9525）
    TopLeftOffsetY = 0,           // 垂直偏移
    WidthPixels = 200,
    HeightPixels = 120,
    MoveMode = ImageMoveMode.MoveAndSizeWithCells, // 随格移动+缩放
}, extension: "png", name: "logo", altText: "公司 Logo");
```

`ImageMoveMode` 三种模式：

| 模式 | OOXML | 行为 |
|---|---|---|
| `MoveButDontSizeWithCells`（默认） | oneCellAnchor | 随单元格移动，不缩放 |
| `MoveAndSizeWithCells` | twoCellAnchor | 随单元格移动并缩放（图片跟随格子拉伸） |
| `FixedPosition` | oneCellAnchor editAs="absolute" | 固定位置，不随单元格移动/缩放 |

> `twoCellAnchor` 的终止位置按默认列宽≈64px/行高≈20px 估算，非默认格子下图片按格子缩放。`ImageAnchor` 仅 Floating 生效；InCell 忽略锚点。

---
## 15. 数据验证（下拉列表）

### 写出数据验证

```csharp
var sheet = new SheetData
{
    Headers = new() { "姓名", "部门", "分数" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromText("IT"), Cell.FromNumber(85) },
    },
    Validations = new()
    {
        // 下拉列表
        new DataValidation
        {
            Type = DataValidationType.List,
            Sqref = "B2:B100",                          // 应用范围
            Formula1 = "\"IT,HR,Finance,Sales\"",       // 下拉选项（用双引号包裹）
            AllowBlank = true,
            PromptTitle = "部门",
            Prompt = "请从列表选择部门",
        },
        // 整数范围 0-100
        new DataValidation
        {
            Type = DataValidationType.WholeNumber,
            Sqref = "C2:C100",
            Formula1 = "0",
            Formula2 = "100",
            AllowBlank = false,
        },
    },
};

XlsxWriter.Write("validation.xlsx", sheet);
```

### 数据验证类型

| `DataValidationType` | 说明 | `Formula1` | `Formula2` |
|---|---|---|---|
| `List` | 下拉列表 | `"\"选项1,选项2\""` | 不用 |
| `WholeNumber` | 整数 | 下限 | 上限 |
| `Decimal` | 小数 | 下限 | 上限 |
| `Date` | 日期 | 起始日期 | 结束日期 |

### 读取数据验证

```csharp
var sheet = XlsxReader.Read("validation.xlsx", 0);
if (sheet.Validations is not null)
{
    foreach (var v in sheet.Validations)
    {
        Console.WriteLine($"类型: {v.Type}, 范围: {v.Sqref}, 公式: {v.Formula1}");
    }
}
```

---

## 16. 追加数据

### 追加到已有文件

```csharp
// 先写 3 行
XlsxWriter.Write("data.xlsx", new SheetData
{
    SheetName = "数据",
    Headers = new() { "ID" },
    Rows = new()
    {
        new Cell[] { Cell.FromNumber(1) },
        new Cell[] { Cell.FromNumber(2) },
        new Cell[] { Cell.FromNumber(3) },
    },
});

// 追加 2 行
XlsxWriter.Append("data.xlsx", new SheetData
{
    SheetName = "数据",  // 同名 → 追加到该 sheet
    Headers = new() { "ID" },
    Rows = new()
    {
        new Cell[] { Cell.FromNumber(4) },
        new Cell[] { Cell.FromNumber(5) },
    },
});

// 读回 → 5 行
var sheet = XlsxReader.Read("data.xlsx", 0);
Console.WriteLine(sheet.Rows.Count);  // 5
```

### 追加到不存在的 sheet

如果 `newData.SheetName` 在原文件中不存在，会作为新 sheet 加入。

### 原文件不存在

如果原文件不存在，`Append` 等同于 `Write`，直接创建新文件。

> 对已有文件执行 `Append` 时，LiteExcel 会保留现有的文档属性（作者、标题、主题、创建时间等），并自动将最后修改时间更新为当前时间。

> `Append` 会基于 LiteExcel 已读取的模型重新生成工作簿。样式、合并单元格、筛选、批注、数据验证等已支持数据会被保留；Excel Table、Theme、透视表、图表等尚未映射到 LiteExcel 模型的 OOXML 部件不保证保留。

可选传入第三个参数覆盖指定属性：

```csharp
XlsxWriter.Append("data.xlsx", moreRows, new WorkbookProperties
{
    LastModifiedBy = "张三",
    Title = "更新后的报表",
});
```

### 表头对齐

追加时，如果新数据的表头和原表头不一致：
- 新表头中不存在于原表头的列，会追加到原表头末尾
- 数据行按列名对齐（新增列补 Empty）

---

## 17. List&lt;T&gt; 映射（反射，AOT 安全）

> **注意**：List&lt;T&gt; API 使用反射，已用 `[DynamicallyAccessedMembers]` 标注，兼容 AOT/裁剪（详见 §24）。调用时请使用具体类型（如 `Excel.Write<Person>`）。

### 基本写出

```csharp
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public DateTime Birthday { get; set; }
}

var list = new List<Person>
{
    new() { Name = "张三", Age = 25, Birthday = new DateTime(2000, 1, 1) },
    new() { Name = "李四", Age = 30, Birthday = new DateTime(1995, 5, 10) },
};

XlsxWriter.Write("people.xlsx", list);
```

### 基本读取

```csharp
var list = XlsxReader.Read<Person>("people.xlsx");
```

### 建簿并写入（Excel.Create）

```csharp
// 返回可继续使用的工作簿（样式/冻结/条件格式/密码等都能加）
var wb = Excel.Create(list, "员工表");
wb.Worksheets[0].HeaderStyle = new CellStyle { Bold = true };
wb.SaveAs("people.xlsx");

// 批量加表并写入
wb.Worksheets.Add("历史", moreList);
```

### 导入到已有工作表（Worksheet.ImportData）

```csharp
var wb = Excel.Open("people.xlsx");
var ws = wb.Worksheets[0];
ws.ImportData(newList);                 // 清空现有内容，从 A1 重建（首行表头）
ws.ImportData(dataTable, includeHeader: false);   // DataTable 版，可不写表头
wb.Save();
```

`ImportData` 清空整表后重建（合并区一并清除），导入后仍可继续修改单元格样式。

### [LiteColumn] 特性

```csharp
public class Product
{
    [LiteColumn(Name = "产品编码", Order = 0)]
    public string Code { get; set; } = "";

    [LiteColumn(Name = "产品名称", Order = 1)]
    public string Name { get; set; } = "";

    [LiteColumn(Name = "单价", Order = 2, Format = "#,##0.00")]
    public decimal Price { get; set; }

    [LiteColumn(Order = 3, Format = "yyyy-MM-dd")]
    public DateTime CreatedAt { get; set; }

    [LiteColumn(Ignore = true)]
    public string? InternalRemark { get; set; }  // 不输出

    public int Stock { get; set; }  // 无特性，默认输出
}
```

| 属性 | 说明 |
|---|---|
| `Name` | 列名（默认用属性名） |
| `Order` | 列顺序（-1 按声明顺序） |
| `Format` | 数字/日期格式 |
| `IsFormula` | true 则把该字符串属性当公式写出（值可带或不带 `=`） |
| `Ignore` | true 则不输出/读取 |

### Fluent 配置（写出）

```csharp
XlsxWriter.Write("people.xlsx", list, opt => opt
    .Column(x => x.Name, "姓名")
    .Column(x => x.Age, "年龄")
    .Column(x => x.Birthday, "生日", "yyyy-MM-dd")
    .Column(x => x.SumFormula, "合计", isFormula: true)   // 公式列
    .Ignore(x => x.InternalRemark));

// 也可设 SheetName 和 FreezeHeader
XlsxWriter.Write("people.xlsx", list, opt =>
{
    opt.Column(x => x.Name, "姓名");
    opt.FreezeHeader = true;
    opt.SheetName = "员工";
});
```

### Fluent 配置（读取）

```csharp
var list = XlsxReader.Read<Person>("people.xlsx", 0, opt => opt
    .Column(x => x.Name, "Full Name")
    .Column(x => x.Age, "Years"));
```

### 字典映射

```csharp
var mapping = new Dictionary<string, string>
{
    { "Name", "姓名" },
    { "Age", "年龄" },
};

XlsxWriter.Write("people.xlsx", list, opt => opt.Map(mapping));
var list = XlsxReader.Read<Person>("people.xlsx", 0, opt => opt.Map(mapping));
```

### 支持的数据类型

自动转换：`int`/`long`/`double`/`float`/`decimal`/`DateTime`/`bool`/`string` 及其可空版本。

### 超级表（Table / ListObject）

超级表是 Excel 的"插入 ▸ 表"——带条纹样式、表头筛选、结构化范围的区域。

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets[0];

// 先准备数据（首行为表头）
ws.SetValue("A1", "姓名"); ws.SetValue("B1", "薪资"); ws.SetValue("C1", "入职日期");
for (int i = 2; i <= 501; i++) { ws.SetValue($"A{i}", $"员工{i}"); ws.SetValue($"B{i}", i * 100); }

// 建表（区域至少 2 行：表头 + 1 行数据）
var tbl = ws.AddTable("A1:C501", "员工表");          // 默认 TableStyleMedium9
ws.AddTable("E1:F10", "历史表", TableStyleStyle.Medium6);  // 或指定枚举

// 列格式（用户自定义）：映射到列 dxf
tbl.Column("薪资").NumberFormat = "#,##0.00";
tbl.Column("入职日期").NumberFormat = "yyyy/m/d";
tbl.Column("姓名").Style = new CellStyle { FontColor = "#FF0000", Bold = true };

// 样式开关
tbl.ShowRowStripes = true;      // 斑马线（默认 true）
tbl.ShowFirstColumn = false;    // 首列强调
tbl.ShowLastColumn = false;     // 末列强调
tbl.ShowColumnStripes = false;  // 列条纹
tbl.AutoFilter = true;          // 表头筛选下拉（默认 true）
tbl.CustomStyleName = "TableStyleDark3";   // 或用任意样式名字符串（优先生效）

wb.SaveAs("table.xlsx");

// 读回
foreach (var t in Excel.Open("table.xlsx").Worksheets[0].Tables)
    Console.WriteLine($"{t.Name} {t.Ref} {t.StyleName}");

// 删除
ws.RemoveTable("员工表");
```

> **条纹样式机理**：条纹长什么样由 Excel 程序内置渲染，**文件里只存样式名字符串**。Excel 内置 60 种样式名：`TableStyleLight1~21`、`TableStyleMedium1~28`、`TableStyleDark1~11`（`None` 表示无样式）。名字写错/大小写错时 Excel 静默退化为无样式（不报错）；库侧用字符串重载写了未知名字会经 `OnDegradation` 上报提示。

> **命名规则**：表名全簿唯一；不能以数字开头、不能含空格、不能是单元格地址（如 `C1`）；允许中文。

> **支持范围**：仅 xlsx/xlsm 写出与读回；xls/xlsb/csv 走 `OnDegradation` 降级上报（不写表）。暂不支持：汇总行、结构化引用（`表名[#All]`）、自定义条纹配色定义。

---
## 18. DataTable 便利 API（AOT 安全）

> DataTable 自带 schema，无需反射，AOT 安全。

### 从 DataTable 写出

```csharp
var dt = new DataTable("订单");
dt.Columns.Add("OrderID", typeof(int));
dt.Columns.Add("Customer", typeof(string));
dt.Columns.Add("Amount", typeof(decimal));
dt.Columns.Add("Date", typeof(DateTime));

dt.Rows.Add(1001, "Alice", 599.99m, new DateTime(2024, 6, 1));
dt.Rows.Add(1002, "Bob", 1299.50m, new DateTime(2024, 6, 15));

XlsxWriter.Write("orders.xlsx", dt, "Orders");
```

### 读为 DataTable

```csharp
var dt = XlsxReader.ReadAsDataTable("orders.xlsx", sheetIndex: 0);
// 或按名
var dt = XlsxReader.ReadAsDataTable("orders.xlsx", "Orders");

foreach (DataRow row in dt.Rows)
{
    Console.WriteLine($"#{row["OrderID"]} | {row["Customer"]} | {row["Amount"]:C}");
}
```

### 建簿并导入

```csharp
// 一步建簿并写 DataTable（表名空则取 DataTable.TableName，再空则 Sheet1）
var wb = Excel.Create(dt);
wb.SaveAs("orders.xlsx");

// 或导入到已有工作表（清空整表后从 A1 重建；includeHeader=false 不写列名行）
var opened = Excel.Open("orders.xlsx");
opened.Worksheets[0].ImportData(dt, includeHeader: false);
opened.Save();
```

---

## 19. Stream 读写

所有读写 API 都有 Stream 重载，适合内存流/网络流场景。

### Stream 写出

```csharp
using var ms = new MemoryStream();
XlsxWriter.Write(ms, sheet);
// ms 现在包含 xlsx 字节
File.WriteAllBytes("output.xlsx", ms.ToArray());
```

### Stream 读取

```csharp
using var fs = File.OpenRead("output.xlsx");
var sheet = XlsxReader.Read(fs, sheetIndex: 0);
```

### Stream 写出多表

```csharp
using var ms = new MemoryStream();
XlsxWriter.Write(ms, new[] { sheet1, sheet2 });
```

### Stream 读取全部表

```csharp
using var fs = File.OpenRead("output.xlsx");
var allSheets = XlsxReader.ReadAll(fs);
```

### Stream 流式读取

```csharp
using var fs = File.OpenRead("bigfile.xlsx");
XlsxReader.StreamRows(fs, "Sheet1", row => { /* 处理行 */ });
```

### Stream 读为 DataTable

```csharp
using var fs = File.OpenRead("output.xlsx");
var dt = XlsxReader.ReadAsDataTable(fs, 0);
```

> **注意**：Stream 重载不会关闭传入的 Stream（`leaveOpen: true`），调用方负责 Stream 生命周期。

---

## 20. 流式读取与进度回调

### StreamRows（逐行回调）

```csharp
XlsxReader.StreamRows("bigfile.xlsx", "Sheet1", row =>
{
    // row 类型是 IReadOnlyList<Cell>
    // 第一行（表头）会被自动跳过
    string name = row[0].Text!;
    int age = (int)row[1].Number;
    // 处理数据...
});
```

### ReadWithProgress（带进度）

```csharp
XlsxReader.ReadWithProgress("bigfile.xlsx", 0, (current, total) =>
{
    // current: 当前处理到第几行（从 1 开始）
    // total: 数据行总数（不含表头）
    Console.WriteLine($"进度: {current}/{total}");
});

// 也可读出结果
var sheet = XlsxReader.ReadWithProgress("bigfile.xlsx", 0, (current, total) =>
{
    if (current % 100 == 0 || current == total)
        Console.WriteLine($"读取 {current}/{total}");
});
```

---

## 21. 文档属性（作者/时间/标题）

文件属性对话框里显示的信息（作者、最后保存者、创建时间、标题等）。

### 写出时携带文档属性

```csharp
var props = new WorkbookProperties
{
    Creator = "张三",                                    // 作者
    LastModifiedBy = "李四",                              // 最后保存者
    Created = DateTime.Now,                               // 创建时间
    Modified = DateTime.Now,                              // 修改时间
    Title = "月度报表",                                   // 标题
    Subject = "财务",                                     // 主题
    Application = "MyApp",                                // 应用程序名（可选）
};

XlsxWriter.Write("report.xlsx", sheet, props);
```

> `Application` 为 null 时，默认取宿主程序集名（`Assembly.GetEntryAssembly().GetName().Name`）。

### 读取文档属性

```csharp
var props = XlsxReader.ReadProperties("report.xlsx");
Console.WriteLine($"作者: {props.Creator}");
Console.WriteLine($"最后保存者: {props.LastModifiedBy}");
Console.WriteLine($"创建时间: {props.Created}");
Console.WriteLine($"修改时间: {props.Modified}");
Console.WriteLine($"标题: {props.Title}");
Console.WriteLine($"主题: {props.Subject}");
Console.WriteLine($"应用程序: {props.Application}");
```

> 文件无 docProps 时不抛异常，返回空对象（所有字段为 null）。

### 不带属性的写出（向后兼容）

```csharp
// 不传 props，不生成 docProps（行为与旧版一致）
XlsxWriter.Write("output.xlsx", sheet);
```

### WorkbookProperties 字段说明

| 字段 | 类型 | 对应 XML |
|---|---|---|
| `Creator` | `string?` | dc:creator |
| `LastModifiedBy` | `string?` | cp:lastModifiedBy |
| `Created` | `DateTime?` | dcterms:created |
| `Modified` | `DateTime?` | dcterms:modified |
| `Title` | `string?` | dc:title |
| `Subject` | `string?` | dc:subject |
| `Application` | `string?` | app.xml Application |

---
## 22. 文件级安全（打开密码 / 修改密码）

### 打开加密文件

读取带打开密码的 xlsx/xlsm/xlsb 文件时，需通过 `ExcelReadOptions` 提供密码：

```csharp
var wb = Excel.Open("encrypted.xlsx", new ExcelReadOptions
{
    OpenPassword = "1",           // 打开密码
    ModifyPassword = "12",        // 修改密码（可选）
});
```

> 仓库内加密样本（`files/` 目录下）的密码约定：打开密码 = `1`，修改密码 = `12`。例如 `打开修改都需要密码.xlsx`（打开=1、修改=12）、`12.*`（仅修改=12）、`*.`（仅打开=1）。

### 读取安全状态

打开后可通过 `Workbook.Security` 查询文件安全状态：

```csharp
var security = wb.Security;

bool hasOpenPwd = security.HasOpenPassword;     // 是否有打开密码
bool hasModPwd  = security.HasModifyPassword;   // 是否有修改密码（写保护）
bool hasModAcc  = security.HasModifyAccess;     // 是否已获得修改授权（乐观授权）
bool isReadOnly = security.IsReadOnly;          // 是否只读
bool canSave    = security.CanSave;             // 当前能否保存
```

### 设置密码

```csharp
// 设置打开密码（保存时对文件加密）
wb.Security.SetOpenPassword("mySecret");

// 设置修改密码（写保护，fileSharing，非 zip 加密）
wb.Security.SetModifyPassword("myModSecret");

wb.SaveAs("protected.xlsx");
```

### 移除密码

```csharp
// 移除打开密码
wb.Security.RemoveOpenPassword();

// 移除修改密码（需已获得修改授权，即打开时提供了 ModifyPassword）
wb.Security.RemoveModifyPassword();

wb.SaveAs("plain.xlsx");
```

### 密码继承与授权规则

- 打开加密文件后，`SaveAs` **默认继承打开密码**；如需移除，保存前调用 `Security.RemoveOpenPassword()`。
- 修改密码是写保护（`<fileSharing>`），不是 zip 加密；读取时提供 `ModifyPassword` 即视为已授权（乐观授权，不校验样本值）。
- 密码**绝不会**出现在异常消息、日志或测试输出中。
- 仅支持 xlsx/xlsm/xlsb；csv/xls 不支持密码。

### 工作表 / 工作簿保护（2.4.6+）

工作表保护（`sheetProtection`）与工作簿保护（`workbookProtection`）是**编辑限制**（区别于上面的文件级密码）：

```csharp
// 工作表保护：锁定后按标志位限制编辑，可选密码
wb.Worksheets[0].Protection = new SheetProtection
{
    Enabled = true,
    SelectLockedCells = true,
    SelectUnlockedCells = false,   // 不允许选定未锁定单元格
    Sort = false,                  // 不允许排序
    AutoFilter = false,            // 不允许自动筛选
};
wb.Worksheets[0].Protection.SetPassword("pw");

// 工作簿保护：锁定结构（禁止增删/移动/重命名工作表）
wb.Protection = new WorkbookProtection
{
    Enabled = true,
    LockStructure = true,
    LockWindows = false,
};
wb.Protection.SetPassword("wb");

wb.SaveAs("protected.xlsx");
```

- 读回：`wb.Worksheets[i].Protection` / `wb.Protection` 自动填充；`VerifyPassword(pwd)` 校验保护密码。
- 密码以 SHA-512 + salt 哈希写出（与 `fileSharing` 同机制），不含明文。
- 仅支持 xlsx/xlsm 写出；xls/xlsb 走既有降级上报。
- 已由 Excel COM 真实验证：打开后 `ProtectContents` / `ProtectStructure` 为 True。

---

## 23. 错误处理

### LiteExcelException

所有面向用户的错误统一抛 `LiteExcelException`：

```csharp
try
{
    var sheet = XlsxReader.Read("not-an-xlsx.txt", 0);
}
catch (LiteExcelException ex)
{
    Console.WriteLine($"读取失败: {ex.Message}");
    // "这不是有效的 xlsx 文件"
}
```

### InvalidSheetNameException

写出时 Sheet 名不合法：

```csharp
try
{
    XlsxWriter.Write("bad.xlsx", new SheetData { SheetName = "sheet[1]" });
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine(ex.Message);
}
```

### 常见错误

| 场景 | 异常类型 | 提示信息 |
|---|---|---|
| 非 xlsx 文件 | `LiteExcelException` | "这不是有效的 xlsx 文件" |
| Sheet 名不存在 | `LiteExcelException` | "找不到工作表：{name}（共有 {n} 张表）" |
| Sheet 名非法 | `InvalidSheetNameException` | 含具体原因 |
| Sheet 索引越界 | `ArgumentOutOfRangeException` | "工作表索引超出范围" |
| 空表列表 | `ArgumentException` | "至少需要一张工作表" |

> 异常消息目前为中文。

---

## 24. AOT 兼容性

全部公开 API 均兼容 Native AOT / 裁剪。库以 `IsAotCompatible` 编译，并经过两类实测验证：

- **编译期**：仓库内置 `tests/LiteExcel.AotSmoke`（`PublishAot` + `TrimmerRootAssembly Include="LiteExcel"`），ILC 遍历库全部代码路径做裁剪分析，publish 零 IL 警告。
- **运行期**：以原生 AOT 可执行文件实测覆盖 xlsx/xlsb/xls/csv 读写、`[LiteColumn]`、公式列、Fluent 表达式、可空/decimal、DataTable、`Excel.Create<T>`/`ImportData`/`Add<T>`、条件格式、命名区域、图片（Floating/InCell）、流式读写、降级回调、超链接、冻结窗格，全部断言通过。

| API | AOT 说明 |
|---|---|
| 对象模型（`Workbook` / `Worksheet` / `Cell` / `ExcelRange` / `Cells`） | 无反射 |
| 流式（`Excel.CreateWriter` / `Excel.StreamRows`） | 无反射 |
| `SheetData` 读写（`Read` / `Write` / `Append` / `AutoColumnWidths`） | 无反射 |
| DataTable（`ReadAsDataTable` / `Write(path, DataTable)` / `Create(DataTable)` / `ImportData(DataTable)`） | 无反射 |
| List&lt;T&gt; 映射（`Excel.Read<T>` / `Excel.Write<T>` / `Excel.Create<T>` / `Worksheet.ImportData<T>` / `WorksheetCollection.Add<T>`） | 使用反射，已用 `[DynamicallyAccessedMembers]` 标注，AOT/裁剪安全 |
| `WriteOptions.Column<TProp>` / `Ignore<TProp>` | 只读表达式树的 `Member.Name`，从不 `.Compile()`，无成员反射 |

> **调用约束**：`List<T>` 泛型 API 请使用具体类型（如 `Excel.Write<Employee>`）。若在你的未标注开放泛型方法中转发，编译器会提示 IL2091 警告，需给对应类型参数补上标注——这是正确的传染行为。

> **读取入口边界**：`Excel.Read<T>` / `XlsxReader.Read` 仅支持 xlsx/xlsm（OpenXML 文本格式）；**xls / xlsb / csv 的读取请走 `Excel.Open(path)`**（门面按扩展名路由到对应后端）。

> **InvariantGlobalization**：`InvariantGlobalization=true`（AOT 项目常见配置）下，19 项原生可执行文件断言全部通过，含 xls/xlsb 读写。注意 xls 的 ANSI 文本在 Invariant 模式下按 Latin1 解码，非当前系统代码页的字符（如中文 GBK 文本）可能失真——这是 xls 二进制格式的固有限制，与 AOT 无关。

---

## 25. 条件格式（2.4.3+）

条件格式支持四类：单元格比较 / 公式 / 色阶 / 数据条。

### 单元格比较

```csharp
var ws = wb.Worksheets[0];
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.CellIs,
    Sqref = "B2:B100",
    Operator = ConditionalOperator.GreaterThan,
    Formula = "60",
    Style = new CellStyle { FontColor = "#FF0000" }, // 预警红
});
```

### 公式条件

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.Expression,
    Sqref = "A2:E100",
    Formula = "MOD(ROW(),2)=0",
    Style = new CellStyle { FillColor = "#F2F2F2" },
});
```

### 色阶（低/中/高 2~3 色）

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.ColorScale,
    Sqref = "D2:D100",
    ColorScale = new ColorScaleInfo
    {
        LowColor = "#FF0000",
        MidColor = "#FFFF00",
        HighColor = "#00FF00",
    },
});
```

### 数据条

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.DataBar,
    Sqref = "E2:E100",
    DataBar = new DataBarInfo
    {
        Color = "#63C384",
        ShowValue = true,
        MinLengthPercent = 0,
        MaxLengthPercent = 100,
    },
});
```

### 读取

xlsx/xlsm 打开含条件格式的工作簿，`Worksheet.ConditionalFormats` 已回填：**xls/xlsb/csv 不支持条件格式**，写出那会按 [§26 能力降级回调](#26-能力降级回调ondegradation242) 上报。

### 长尾类型（2.4.4+）

除 `cellIs`/`expression`/`colorScale`/`dataBar` 外，追加支持：

| 类型 | 枚举 | 说明 | 关键字段 |
|---|---|---|---|
| 文本包含 | `ContainsText` | 含指定文本 | `Text` |
| 文本开头 | `BeginsWith` | 以指定文本开头 | `Text` |
| 文本结尾 | `EndsWith` | 以指定文本结尾 | `Text` |
| 文本不含 | `NotContainsText` | 不含指定文本 | `Text` |
| 时间周期 | `TimePeriod` | 昨天/今天/上周等 | `TimePeriod`（如 `"thisMonth"`） |
| 空单元格 | `Blanks` | 空白 | — |
| 非空单元格 | `NoBlanks` | 非空白 | — |
| 错误 | `Errors` | 错误值 | — |
| 非错误 | `NoErrors` | 非错误值 | — |
| 唯一值 | `Unique` | 本范围内唯一 | — |
| 重复值 | `Duplicate` | 本范围内重复 | — |
| 前 N 项 | `Top10` | 前 N 项或前 N% | `Rank` + `Percent` |
| 高于平均 | `AboveAverage` | 高于本范围平均分 | — |
| 低于平均 | `BelowAverage` | 低于本范围平均分 | — |

示例：

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.ContainsText,
    Sqref = "B2:B100",
    Text = "urgent",
    Style = new CellStyle { FontColor = "#FF0000" },
});
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.Top10,
    Sqref = "A2:A100",
    Rank = 3,        // 前 3 项；Percent=true 则前 3%
});
```

### 图标集（iconSet，2.4.6+）

在单元格内显示箭头 / 红绿灯 / 符号 / 星级等图标的条件可视化：

```csharp
// 最简：三色箭头，百分比均分阈值（0/33/67）
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.IconSet,
    Sqref = "B2:B10",
    IconSet = new IconSetInfo(),
});

// 指定集合 + 自定义阈值
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Type = ConditionalFormatType.IconSet,
    Sqref = "C2:C10",
    IconSet = new IconSetInfo
    {
        Style = IconSetStyle.FourRating,
        Percent = true,
        ShowValue = true,
        Thresholds = new double[] { 25, 50, 75 },   // 图标数 - 1 个；为空则均分
    },
});

// 读回
foreach (var cf in wb.Worksheets[0].ConditionalFormats)
    if (cf.Type == ConditionalFormatType.IconSet && cf.IconSet is not null)
        Console.WriteLine($"{cf.Sqref} {cf.IconSet.Style}");
```

> **机理**：图标长什么样由 Excel 程序内置渲染，文件只存集合名字符串（同条纹样式）。Excel 内置 17 种集合：`3Arrows` / `3ArrowsGray` / `3Flags` / `3TrafficLights1` / `3TrafficLights2` / `3Signs` / `3Symbols` / `3Symbols2` / `4Arrows` / `4ArrowsGray` / `4RedToBlack` / `4Rating` / `4TrafficLights` / `5Arrows` / `5ArrowsGray` / `5Rating` / `5Quarters`。枚举 `IconSetStyle` 覆盖全部；`CustomStyleName` 可写任意集合名字符串。
>
> **阈值约定**：默认按图标数均分（3 图标=0/33/67，4=0/25/50/75，5=0/20/40/60/80）；`Thresholds` 自定义（数量 = 图标数 - 1）。`Percent=false` 时阈值为绝对数值（`cfvo type="num"`）。xls/xlsb/csv 不支持条件格式，走降级上报。

---

## 26. 命名区域（Name Range）

工作簿中的命名区域（如 `Score = S1!$B$2:$B$6`）通过 `Workbook.Names` 读取。包含：

- `Name` — 名称（如 `"Score"`）
- `Reference` — 原始范围文本（如 `"S1!$B$2:$B$6"`，含 sheet 限定）
- `LocalSheetId` — 并非 `-1` 时表示该名称为工作表局部；`-1` 表示工作簿全局
- `IsLocalSheet` — 布尔快速地判断是否为局部名称

```csharp
var wb = Excel.Open("data.xlsx");
foreach (var nr in wb.Names)
{
    Console.WriteLine($"Name={nr.Name} Ref={nr.Reference} Local={nr.LocalSheetId}");
}
```

写入：不修改 `Workbook.Names` 的前提下，保存时原始 XML 原样回写（内容不会丢）。

---

## 27. 能力降级回调（OnDegradation）（2.4.2+）

保存到不支持某能力的格式时（如 xls/xlsb/csv），默认会静默丢弃这些能力。开启回调可直接收到被丢弃项的清单，不再沉默。

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets[0];
ws.SetValue("A1", "x");
ws.Filter = new AutoFilter { Range = "A1:A10" };   // 筛选
ws.Comments = new Dictionary<string, string> { ["A1"] = "note" }; // 批注
ws.FreezeRows = 1;                                  // 冻结
ws.Cell("B1").Style = new CellStyle { Bold = true };

var options = new ExcelWriteOptions
{
    OnDegradation = d =>
        Console.WriteLine($"[{d.Capability}] {d.SheetName} => {d.Message}")
};
Excel.Write("out.csv", wb, options);
```

默认 `null`：与历史版本行为完全一致（零破坏）。

### 降级能力枚举

`DegradationCapability`：Comments、DataValidation、AutoFilter、Images、DocumentProperties、NamedRanges、Styles、MergedCells、FreezePanes、Hyperlinks、RowHeights、ColumnWidths、Formulas、Charts、PivotTables、RichData。

### 各格式降级矩阵（2.4.2）

| 能力 | xlsx/xlsm | xlsb | xls | csv |
|---|---|---|---|---|
| 批注 | 支持 | 降级上报 | 降级上报 | 降级上报 |
| 数据验证 | 支持 | 降级上报 | 降级上报 | 降级上报 |
| 自动筛选 | 支持 | 降级上报 | 降级上报 | 降级上报 |
| 图片 | 支持（仅写回） | 降级上报 | 降级上报 | 降级上报 |
| 单元格样式 | 支持 | **仅数字格式保留**，其余上报 | **仅数字格式保留**，其余上报 | 降级上报 |
| 超链接 | 支持 | 支持 | 支持 | 降级上报 |
| 公式 | 支持 | 仅缓存值文本 | 仅缓存值文本 | 降级上报（值文本） |
| 行高/列宽/合并/冻结 | 支持 | 支持 | 支持 | 降级上报 |
| 图表/透视表/主题 | 保真透传 | 保真透传 | 不支持（永不做） | 不支持 |

`TargetFormat` 字段为 `ExcelFormat`；库内部不会修改 `DegradationInfo`。

---

## 28. 完整 API 索引

### Excel（门面）

| 方法 | 返回 | 说明 |
|---|---|---|
| `Open(path, options?)` / `Open(stream, format, options?)` | `Workbook` | 按扩展名识别格式打开 |
| `DetectFormat(path)` | `ExcelFormat` | 按扩展名识别格式 |
| `Create(format = Xlsx)` / `Create(sheetName, format)` / `Create(string[], format)` | `Workbook` | 新建工作簿（空壳 / 命名 / 批量表） |
| `Create<T>(data, sheetName, format, configure?)` | `Workbook` | 建簿并写入 List&lt;T&gt;（首行表头，反射，AOT 安全） |
| `Create(DataTable, sheetName?, format)` | `Workbook` | 建簿并写入 DataTable（表名兜底 TableName → Sheet1） |
| `Write(path, Workbook, options?)` | `void` | 写出工作簿 |
| `Write(path, SheetData, options?)` | `void` | 写出单表（低层模型） |
| `Write(path, DataTable, sheetName?, options?)` | `void` | 从 DataTable 写（AOT 安全） |
| `Write<T>(path, IEnumerable<T>, sheetName?, configure?)` | `void` | 从 List&lt;T&gt; 写（反射，AOT 安全） |
| `Read<T>(path, sheetName?, configure?)` | `List<T>` | 读为 List&lt;T&gt;（仅 xlsx/xlsm） |
| `ReadAsDataTable(path, sheetName?, firstRowIsHeader?)` | `DataTable` | 读为 DataTable |
| `GetSheetNames(path)` | `List<string>` | 列出所有工作表名 |
| `StreamRows(path, sheetName, onRow)` | `void` | 流式读取（自动跳过首行） |
| `CreateWriter(path)` / `CreateWriter(stream)` | `XlsxStreamWriter` | 流式写入（仅 xlsx/xlsm） |

### Worksheet

| 成员 | 说明 |
|---|---|
| `Cell(row, col)` / `Cell("A1")` / `Range(...)` / `Cells` | 单元格 / 区域访问 |
| `SetValue(row, col / address, value)` | 写值 |
| `ImportData<T>(data, configure?)` | 清空整表并从 A1 重建（List&lt;T&gt;，首行表头） |
| `ImportData(DataTable, includeHeader = true)` | 清空整表并从 A1 重建（DataTable） |
| `AddImage(byte[] data, row, col, ...)` | 添加图片（Floating/InCell） |
| `FreezeRows` / `FreezeColumns` / `FreezeHeader` | 冻结窗格 |
| `Merge(...)` / `Unmerge(...)` / `MergedRanges` | 合并单元格 |
| `HeaderStyle` / `DefaultStyle` / `RowStyles` / `ColumnStyles` | 样式 |
| `RowHeights` / `ColumnWidths` | 行高 / 列宽 |
| `Comments` / `Validations` / `Filter` / `ConditionalFormats` / `Images` | 批注 / 数据验证 / 筛选 / 条件格式 / 图片 |
| `ToSheetData()` | 转回低层模型（写回用） |

### XlsxReader

| 方法 | 返回 | 说明 |
|---|---|---|
| `GetSheetNames(string path)` | `List<string>` | 列出所有工作表名 |
| `GetSheetNames(Stream stream)` | `List<string>` | 从流列出 |
| `Read(string path, int sheetIndex, bool firstRowIsHeader = true)` | `SheetData` | 按索引读 |
| `Read(string path, string sheetName, bool firstRowIsHeader = true)` | `SheetData` | 按名读 |
| `Read(Stream stream, int sheetIndex, bool firstRowIsHeader = true)` | `SheetData` | 从流按索引读 |
| `Read(Stream stream, string sheetName, bool firstRowIsHeader = true)` | `SheetData` | 从流按名读 |
| `ReadAll(string path)` | `List<SheetData>` | 读全部表 |
| `ReadAll(Stream stream)` | `List<SheetData>` | 从流读全部 |
| `StreamRows(string path, string sheetName, Action<IReadOnlyList<Cell>> onRow)` | `void` | 流式逐行读 |
| `StreamRows(Stream stream, string sheetName, Action<IReadOnlyList<Cell>> onRow)` | `void` | 从流式逐行读 |
| `ReadWithProgress(string path, int sheetIndex, Action<int,int> onProgress)` | `void` | 带进度读取 |
| `ReadAsDataTable(string path, int sheetIndex = 0, bool firstRowIsHeader = true)` | `DataTable` | 读为 DataTable |
| `ReadAsDataTable(string path, string sheetName, bool firstRowIsHeader = true)` | `DataTable` | 按名读为 DataTable |
| `ReadAsDataTable(Stream stream, int sheetIndex = 0, bool firstRowIsHeader = true)` | `DataTable` | 从流读为 DataTable |
| `ReadAsDataTable(Stream stream, string sheetName, bool firstRowIsHeader = true)` | `DataTable` | 从流按名读 |
| `Read<T>(string path, int sheetIndex = 0, Action<ReadOptions<T>>? configure = null)` | `List<T>` | 读为 List&lt;T&gt;（反射，AOT 安全） |
| `Read<T>(string path, string sheetName, Action<ReadOptions<T>>? configure = null)` | `List<T>` | 按名读为 List&lt;T&gt; |
| `ReadProperties(string path)` / `ReadProperties(Stream stream)` | `WorkbookProperties` | 读取文档属性（作者/时间/标题） |

### XlsxWriter

| 方法 | 说明 |
|---|---|
| `Write(string path, SheetData data)` | 写单表 |
| `Write(string path, IReadOnlyList<SheetData> sheets)` | 写多表 |
| `Write(Stream stream, SheetData data)` | 写单表到流 |
| `Write(Stream stream, IReadOnlyList<SheetData> sheets)` | 写多表到流 |
| `Write(path/stream, sheets, WorkbookProperties? properties)` | 写出并携带文档属性 |
| `Write(string path, DataTable table, string sheetName = "Sheet1")` | 从 DataTable 写 |
| `Write<T>(string path, IEnumerable<T> data, Action<WriteOptions<T>>? configure = null)` | 从 List&lt;T&gt; 写（反射，AOT 安全） |
| `Append(string path, SheetData? newData, WorkbookProperties? updateProperties = null)` | 追加数据并可更新文档属性 |
| `AutoColumnWidths(SheetData sheet)` | 自动估算列宽 |

### CellRef（工具类）

| 方法 | 说明 |
|---|---|
| `CellRef.Parse(string cellRef)` | "A1" → (row=0, col=0) |
| `CellRef.ToString(int row, int col)` | (0, 0) → "A1" |
| `CellRef.ColToLetter(int col)` | 0 → "A", 26 → "AA" |
| `CellRef.LetterToCol(string letters)` | "A" → 0, "AA" → 26 |

### 模型类

| 类 | 说明 |
|---|---|
| `SheetData` | 工作表数据 |
| `Cell` | 单元格 |
| `CellStyle` | 样式 |
| `BorderStyle` / `BorderEdge` | 边框 |
| `CellRange` | 区域（合并单元格） |
| `AutoFilter` / `FilterColumn` | 自动筛选 |
| `DataValidation` | 数据验证 |
| `Hyperlink` | 超链接（Target / Tooltip / IsInternal） |
| `WorksheetImage` | 工作表图片 |
| `WorkbookSecurity` | 文件安全状态（打开密码/修改密码/只读/可保存） |
| `LiteExcelException` / `InvalidSheetNameException` | 异常 |
| `LiteColumnAttribute` | List&lt;T&gt; 列特性 |
| `WorkbookProperties` | 文档属性（作者/时间/标题/应用名） |
| `WriteOptions<T>` / `ReadOptions<T>` | List&lt;T&gt; 配置 |

### 枚举

| 枚举 | 值 |
|---|---|
| `CellType` | Text, Number, Date, Boolean, Empty |
| `HorizontalAlignment` | General, Left, Center, Right |
| `VerticalAlignment` | Top, Center, Bottom |
| `FilterType` | Equals, Compare, Contains, BeginsWith, EndsWith, Blank |
| `FilterOperator` | GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between |
| `DataValidationType` | List, WholeNumber, Decimal, Date |
| `ImagePlacement` | Floating, InCell |