# LiteExcel 使用手册

> 本手册反映 LiteExcel 当前主线版本的全部公开能力（对象模型 API + 低层 API）
> 术语约定：**对象模型 API** 指以 `Excel` → `Workbook` → `Worksheet` → `Cell`/`Cells`/`Range` 为主线的日常用法；**低层 API** 指 `SheetData` / `XlsxReader` / `XlsxWriter` / `CsvBackend` / `XlsxStreamWriter` 等裸数据入口，见附录 B。

---

## 目录

**第一部分 入门**
- [1. 安装与引用](#1-安装与引用)
- [2. 快速上手（一段最小可运行的完整读写）](#2-快速上手一段最小可运行的完整读写)

**第二部分 对象模型API正文**
- [3. 文件导航：打开 / 创建 / 保存 / 格式（含工作表管理、文档属性）](#3-文件导航打开-创建-保存-格式含工作表管理文档属性)
- [4. 单元格与取值（Cell / Cells / Range / SetValue）](#4-单元格与取值cell-cells-range-setvalue)
- [5. 数据类型与转换（文本 / 数字 / 日期 / 布尔 / 公式 / 可空 / Byte[]）](#5-数据类型与转换文本-数字-日期-布尔-公式-可空-byte)
- [6. 样式（CellStyle / Border / 对齐 / 换行）](#6-样式cellstyle-border-对齐-换行)
- [7. 合并单元格](#7-合并单元格)
- [8. 自动筛选](#8-自动筛选)
- [9. 行高与列宽（含 AutoColumnWidths）](#9-行高与列宽含-autocolumnwidths)
- [10. 批注](#10-批注)
- [11. 超链接（外部 / 内部）](#11-超链接外部-内部)
- [12. 冻结窗格](#12-冻结窗格)
- [13. 图片（Floating / InCell / 读回）](#13-图片floating-incell-读回)
- [14. 数据验证](#14-数据验证)
- [15. 条件格式（cellIs / expression / colorScale / dataBar / 长尾 / iconSet）](#15-条件格式cellis-expression-colorscale-databar-长尾-iconset)
- [16. 超级表（Table / ListObject，样式枚举 + 任意样式名 + 列格式）](#16-超级表table-listobject样式枚举-任意样式名-列格式)
- [17. 命名区域](#17-命名区域)
- [18. 文件级密码（打开 / 修改）](#18-文件级密码打开-修改)
- [19. 工作表 / 工作簿保护](#19-工作表-工作簿保护)

**第三部分 多格式与平台**
- [20. 多格式行为（xlsx/xlsm 全能 + xls/xlsb/csv 限制与降级）](#20-多格式行为xlsxxlsm-全能-xlsxlsbcsv-限制与降级)
- [21. 流式读取 / 进度回调 / 追加数据](#21-流式读取-进度回调-追加数据)
- [22. 降级回调 OnDegradation](#22-降级回调-ondegradation)
- [23. AOT 兼容性（DAM、IsAotCompatible、验证方式与成果摘要）](#23-aot-兼容性damisaotcompatible验证方式与成果摘要)

**第四部分 注意事项**
- [24. 异常处理](#24-异常处理)
- [25. 大文件注意事项](#25-大文件注意事项)

**附录**
- 附录 A 对象模型速查（类 / 成员索引表）
- 附录 B 低层 API 参考（SheetData / XlsxReader / XlsxWriter / CsvBackend / 流式）

---

## 第一部分 入门

第 1–2 章：安装引用与最小可运行示例，新读者从这里开始。

### 1. 安装与引用

#### NuGet 安装（推荐）

从 NuGet 发布包安装，适用于大多数生产项目：

```powershell
dotnet add package LiteExcel
```

也可在 Visual Studio 的「管理 NuGet 程序包」中搜索 `LiteExcel` 安装。

#### 从源码本地引用

未发包或需联调库源码时，通过 csproj 项目引用引入：

```xml
<ItemGroup>
  <ProjectReference Include="..\src\LiteExcel\LiteExcel.csproj" />
</ItemGroup>
```

> **说明**：生产项目请优先使用 NuGet 包；源码引用仅用于本地联调库源码或尚未发版本时。

#### 命名空间

所有类型都位于 `LiteExcel` 命名空间：

```csharp
using LiteExcel;
```

#### 目标框架

库同时面向 **net48** 与 **net8.0**。net8.0 目标额外声明 `IsAotCompatible=true`，全部公开 API 兼容 Native AOT / 裁剪（详见第 23 章）。

---

### 2. 快速上手（一段最小可运行的完整读写）

下面用对象模型 API 完成「新建 → 写值 → 保存 → 打开 → 读回」的完整闭环：

```csharp
using LiteExcel;

var wb = Excel.Create();                       // 新建工作簿（默认含 Sheet1）
var ws = wb.Worksheets["Sheet1"];

ws.SetValue("A1", "Name");                     // 表头
ws.SetValue("B1", "Age");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B2", 25);
ws.SetValue("A3", "Li Si");
ws.SetValue("B3", 30);

wb.SaveAs("people.xlsx");                      // 保存到磁盘

// 读回
var opened = Excel.Open("people.xlsx");
var sheet = opened.Worksheets[0];
Console.WriteLine(sheet.Cell("A2").GetString());   // 输出: Zhang San
Console.WriteLine(sheet.Cell("B2").GetDouble());   // 输出: 25
```

输出：

```
Zhang San
25
```

---

## 第二部分 对象模型API正文

第 3–19 章：对象模型主线能力，日常读写全部集中于此。

### 3. 文件导航：打开 / 创建 / 保存 / 格式（含工作表管理、文档属性）

#### 3.1 打开已有文件

`Excel.Open` 按扩展名自动识别格式，支持 xlsx / xlsm / xls / xlsb / csv：

```csharp
var wb = Excel.Open("report.xlsx");            // 自动识别
var wb2 = Excel.Open("data.csv");              // 自动识别为 Csv
var wb3 = Excel.Open("legacy.xls");            // 自动识别为 Xls
```

也可以显式指定格式（适用于扩展名与内容不一致的场景）：

```csharp
var wb = Excel.Open("data.bin", ExcelFormat.Xlsx);
```

输出：（本示例无控制台输出）

#### 3.2 从流打开

流没有扩展名，**必须显式指定格式**。输入流不会被关闭（由调用方管理）；不可定位的流（如网络流）内部会复制到内存：

```csharp
using var fs = File.OpenRead("report.xlsx");
var wb = Excel.Open(fs, ExcelFormat.Xlsx);
// 打开后 CurrentPath 为 null，需用 SaveAs 指定保存路径
wb.SaveAs("copy.xlsx");
```

输出：

```
已写入 copy.xlsx
```

从流读取全部表：`Excel.Open(stream, ExcelFormat.Xlsx).Worksheets` 直接遍历多表工作表。

#### 3.3 读取选项 `ExcelReadOptions`

`Open` 的第二参数可传读取选项，含密码、合并填充、CSV 分隔符等：

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",        // 打开密码（文件加密）
    ModifyPassword = "write",       // 修改密码（写保护，用于获得编辑权限）
    FillMergedCells = true,         // 把合并区左上角的值展开到整个区域
    Separator = ';',                // 仅 CSV 生效；null 时自动探测
});
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `OpenPassword` | `string?` | 打开密码（文件加密），解密带密码的 xlsx/xlsm/xlsb |
| `ModifyPassword` | `string?` | 修改密码（写保护），提供后获得编辑/保存权限 |
| `FillMergedCells` | `bool` | 把合并区左上角的值展开到整个合并区域，默认 `false` |
| `Separator` | `char?` | 仅 CSV 生效；`null` 时自动探测 |
| `ReadStyles` | `bool` | 是否读取样式，默认 `true` |
| `LeaveOpen` | `bool` | Stream 重载读取完成后是否保持输入流打开，默认 `false` |

输出：（本示例无控制台输出）

#### 3.4 写入选项 `ExcelWriteOptions`

`Excel.Write` 的第二参数可传写入选项：

```csharp
Excel.Write("out.xlsx", wb, new ExcelWriteOptions
{
    Overwrite = true,           // 目标文件已存在时是否覆盖，默认 true
    AutoFitColumns = true,      // 写出前自动估算列宽，默认 false
    FreezeHeader = true,        // 写出时冻结表头，默认 false
    Properties = new WorkbookProperties { Creator = "Me" },  // 覆盖文档属性
    OnDegradation = info => Console.WriteLine(info.Capability),  // 降级回调
    Separator = ';',            // 仅 CSV 生效；null 时默认逗号
});
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `Overwrite` | `bool` | 目标文件已存在时是否覆盖，默认 `true` |
| `AutoFitColumns` | `bool` | 写出前自动估算列宽，默认 `false` |
| `FreezeHeader` | `bool` | 写出时冻结表头，默认 `false` |
| `Properties` | `WorkbookProperties?` | 覆盖工作簿文档属性 |
| `OnDegradation` | `Action<DegradationInfo>?` | 能力降级回调（写出到不支持某能力的格式时逐项上报，默认 `null`） |
| `Separator` | `char?` | 仅 CSV 生效；`null` 时默认逗号 |
| `LeaveOpen` | `bool` | Stream 重载写入完成后是否保持输出流打开，默认 `false` |

输出：

```
已写入 out.xlsx
```

#### 3.5 新建工作簿

`Excel.Create` 有多个重载：

```csharp
var wb1 = Excel.Create();                    // 空簿，默认 Sheet1
Console.WriteLine(wb1.Worksheets[0].Name);   // 打印验证：默认表名
var wb2 = Excel.Create("Data");              // 指定首个工作表名
var wb3 = Excel.Create(new[] { "Q1", "Q2", "Q3" });   // 批量建表
var wb4 = Excel.Create(ExcelFormat.Xlsm);    // 指定格式
```

一步建簿并写数据（List\<T\>  DataTable）：

```csharp
var people = new List<Person> { new() { Name = "A", Age = 1 } };
var wb5 = Excel.Create(people, "People");    // 首行为表头

var dt = new System.Data.DataTable("T");
dt.Columns.Add("X");
dt.Rows.Add("v");
var wb6 = Excel.Create(dt);                  // sheetName 为空时用 TableName
```

输出：

```
Sheet1
```

#### 3.6 保存与另存

`Workbook.Save` 保存到当前路径（新建簿无路径时抛 `LiteExcelException`）；`SaveAs` 指定路径：

```csharp
wb.Save();                    // 保存到 CurrentPath
wb.SaveAs("out.xlsx");        // 另存，沿用当前格式
wb.SaveAs("out.xlsm", ExcelFormat.Xlsm);   // 另存并转格式
wb.Save(new FileStream("s.xlsx", FileMode.Create), ExcelFormat.Xlsx);  // 存到流
```

输出：

```
已写入 out.xlsx / out.xlsm / s.xlsx
```

> ⚠️ **重要**：本库**不创建 / 不编辑图表（Chart）与数据透视表（PivotTable）**。如果你的数据包含这类元素，**不要**用本库保存/另存覆盖源文件，否则这些元素会被**丢弃**（保存为无图表/无透视表的新文件）。你可以在其他工具中另存一份副本后再处理。

- `SaveAs(path, format)` 要求路径扩展名与格式匹配，否则抛 `LiteExcelException`（避免写出内容与扩展名不一致、Excel 无法打开的文件）。
- 含 VBA 宏的工作簿不允许保存为不支持宏的格式（xlsx / xls），会提前报错。
- 文件级密码（打开 / 修改）仅支持 xlsx / xlsm / xlsb，保存为 csv / xls 时若带密码会报错。

#### 3.7 格式枚举 `ExcelFormat`

```csharp
public enum ExcelFormat { Xlsx, Xlsm, Xlsb, Xls, Csv }
```

- `Excel.DetectFormat(path)` 按扩展名返回格式。
- `Workbook.Format` 返回当前工作簿格式。

输出：（本示例无控制台输出）

#### 3.8 列出工作表名

```csharp
var names = Excel.GetSheetNames("report.xlsx");   // List<string>
using var stream = File.OpenRead("report.xlsx");
var names2 = Excel.GetSheetNames(stream);         // 仅 xlsx/xlsm（zip 容器的 XML 元数据）
```

> `Excel.GetSheetNames(path)` 对 xlsx / xlsm 直接读工作簿元数据（轻量）；对 **xlsb / xls / csv** 会走 `Excel.Open` 按格式解析，也能正确返回表名（xlsb 返回其内部 sheet 名，csv 返回单表 "Sheet1"）。`GetSheetNames(Stream)` 只支持 xlsx / xlsm；需要从流取 xlsb 表名时用 `Excel.Open(stream, ExcelFormat.Xlsb).Worksheets.Names`。

输出：（返回一个List<string>类型）

#### 3.9 工作表管理 `Worksheets`

`Workbook.Worksheets` 为 `WorksheetCollection`，支持增删移动、索引与名称访问：

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];         // 按名称访问（不存在抛 LiteExcelException）
var first = wb.Worksheets[0];             // 按索引访问（0-based）
Console.WriteLine(wb.Worksheets.Count);   // 数量
Console.WriteLine(string.Join(", ", wb.Worksheets.Names));   // 所有表名
Console.WriteLine(wb.Worksheets.Contains("Sheet1"));         // true

wb.Worksheets.Add("Data");                // 新增空表
wb.Worksheets.Add("People", peopleList);  // 新增表并写 List<T>（首行表头）
wb.Worksheets.Add("T", dataTable);        // 新增表并写 DataTable（首行列名）
wb.Worksheets.Move(0, 2);                 // 移动顺序（0-based）
wb.Worksheets.Remove("Data");             // 按名称删除（存在返回 true）
wb.Worksheets.RemoveAt(0);                // 按索引删除
```

遍历工作表：

```csharp
foreach (var sheet in wb.Worksheets)
    Console.WriteLine(sheet.Name);
```

输出：

```
1
Sheet1
True
Sheet1
T
```

> ⚠️ 表名校验规则见第 24 章：不合法表名（含 `\ / ? * [ ] :`、超 31 字符等）在保存时抛 `InvalidSheetNameException`。

#### 3.10 文档属性 `WorkbookProperties`

> ⚠️ 文档属性仅支持 **xlsx / xlsm / xlsb**；不支持 xls（OLE 属性集未实现）。写出 xls 时属性会**静默丢失**，经 `OnDegradation` 上报。

`Workbook.Properties` 对应 xlsx 包内的 `docProps/core.xml` 与 `docProps/app.xml`：

```csharp
var wb = Excel.Create();
wb.Properties.Creator = "JackZ";
wb.Properties.LastModifiedBy = "Me";
wb.Properties.Created = DateTime.Now;
wb.Properties.Modified = DateTime.Now;
wb.Properties.Title = "季度报告";
wb.Properties.Subject = "财务";
wb.Properties.Application = "LiteExcel";
wb.SaveAs("props.xlsx");
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `Creator` | `string?` | 作者（dc:creator） |
| `LastModifiedBy` | `string?` | 最后保存者（cp:lastModifiedBy） |
| `Created` | `DateTime?` | 创建时间 |
| `Modified` | `DateTime?` | 最后修改时间 |
| `Title` | `string?` | 标题 |
| `Subject` | `string?` | 主题 |
| `Application` | `string?` | 应用程序名；`null` 时写出取宿主程序集名 |

写出时覆盖属性（`ExcelWriteOptions.Properties`）：

```csharp
Excel.Write("out.xlsx", wb, new ExcelWriteOptions
{
    Properties = new WorkbookProperties { Creator = "Bot", Title = "Auto" },
});
```

读回文档属性：

```csharp
var opened = Excel.Open("props.xlsx");
Console.WriteLine(opened.Properties.Creator);    // JackZ
Console.WriteLine(opened.Properties.Title);      // 季度报告
```

输出：

```
JackZ
季度报告
```

---


### 4. 单元格与取值（Cell / Cells / Range / SetValue）

#### 4.1 按坐标 / 地址访问单元格

坐标统一 **1-based**。`Worksheet.Cell` 提供按行列或 A1 地址访问：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

var c1 = ws.Cell(1, 1);        // A1
var c2 = ws.Cell("B3");        // B3
Console.WriteLine(c1.IsEmpty);   // True（尚未赋值）
```

`Cell` 是引用类型，访问不存在的单元格返回一个 `Empty` 占位；一旦赋值即落入网格。

输出：

```
True
```

#### 4.2 设置值 `SetValue`

`SetValue` 越界自动扩展网格；`null` / `DBNull` 写空单元格：

```csharp
ws.SetValue(1, 1, "Header");
ws.SetValue("A2", 42);
ws.SetValue("B2", null);       // 清空 B2
Console.WriteLine(ws.Cell("A2").GetString());   // 42
```

输出：

```
42
```

#### 4.3 集合式访问 `Cells`

`Worksheet.Cells` 提供整表入口，支持索引器、区域提取、枚举与批量清空：

```csharp
ws.Cells[1, 1] = "A1 via cells";      // 索引器（行列）
ws.Cells["B1"] = "B1 via cells";      // 索引器（A1 地址）
ws.Cells.SetValue("C1", 3.14);        // 便捷写值

var range = ws.Cells.Range("A1:D10"); // 提取区域
foreach (var cell in ws.Cells)        // 枚举网格中已有单元格
    Console.WriteLine(cell.GetString());
ws.Cells.Clear();                     // 清空整表值（不删行列）
```

输出：

```
A1 via cells
B1 via cells
3.14
```

#### 4.4 区域操作 `ExcelRange`

`Worksheet.Range` 返回连续矩形区域（1-based，含端点），支持批量读写、样式、合并、清空、枚举：

```csharp
var r = ws.Range("A1:D10");          // 或 ws.Range(1, 1, 10, 4)
Console.WriteLine(r.Address);        // "A1:D10"
Console.WriteLine($"{r.RowCount} x {r.ColumnCount}");  // 10 x 4

r.Fill("x");                          // 整区填相同值
r.Fill(new object?[,] { { "a", "b" }, { "c", "d" } }); // 二维数组写入
var vals = r.ToValues();              // object?[,] 读回
var cells = r.ToCells();              // Cell[,] 读回
r.Style = new CellStyle { Bold = true };  // 整区统一样式
r.Merge();                            // 合并该区域
r.Unmerge();                          // 取消合并
r.Clear();                            // 清空区域内值

var single = r.Cell(0, 0);            // 区域内相对偏移（0-based）
```

输出：

```
A1:D10
10 x 4
```

#### 4.5 单元格读取方法

`Cell` 提供强类型与 Try 风格读取：

```csharp
var cell = ws.Cell("A1");
string? s = cell.GetString();        // 文本 / 数字 / 日期 / 布尔按惯例格式化
double d = cell.GetDouble();         // 类型不匹配抛 InvalidCastException
DateTime dt = cell.GetDateTime();
bool b = cell.GetBoolean();
object? raw = cell.GetValue();       // 原始对象，Empty 返回 null

bool ok = cell.TryGetString(out var s2);   // 空单元格返回 false
bool ok2 = cell.TryGetDouble(out double d2);
bool ok3 = cell.TryGetDateTime(out DateTime dt2);
bool ok4 = cell.TryGetBoolean(out bool b2);
```

输出：（单元格原值）

---

### 5. 数据类型与转换（文本 / 数字 / 日期 / 布尔 / 公式 / 可空 / Byte[]）

#### 5.1 单元格类型 `CellType`

```csharp
public enum CellType { Text, Number, Date, Boolean, Empty }
```

`Cell.Type` 决定哪个值字段有效：`Text` / `Number` / `Date` / `Boolean`。`IsEmpty` 表示空单元格。

输出：（本示例无控制台输出）

#### 5.2 工厂方法

```csharp
var t = Cell.FromText("hello");
var n = Cell.FromNumber(42, "#,##0.00");      // 可带数字格式
var d = Cell.FromDate(new DateTime(2024, 1, 1));  // 默认格式 yyyy-MM-dd
var b = Cell.FromBoolean(true);
var f = Cell.FromFormula("SUM(A1:A3)");        // 公式单元格
var e = Cell.Empty;
Console.WriteLine(d.GetString());   // 默认日期格式 yyyy-MM-dd
```

输出：

```
2024-01-01
```

#### 5.3 自动类型转换

`SetValue(object?)` 会按 CLR 类型自动映射：

| CLR 类型 | 单元格类型 |
|---|---|
| `bool` | `Boolean` |
| `DateTime` | `Date`（默认 `yyyy-MM-dd`） |
| `sbyte/byte/short/ushort/int/uint/long/ulong/float/double/decimal` | `Number` |
| `null` / `DBNull` | `Empty` |
| 其他（含 `string`） | `Text` |

```csharp
ws.SetValue("A1", 3.14);        // Number
ws.SetValue("A2", true);        // Boolean
ws.SetValue("A3", DateTime.Now); // Date
ws.SetValue("A4", "text");      // Text
ws.SetValue("A5", null);        // Empty

Console.WriteLine(ws.Cell("A1").Type);   // Number
Console.WriteLine(ws.Cell("A2").Type);   // Boolean
Console.WriteLine(ws.Cell("A3").Type);   // Date
Console.WriteLine(ws.Cell("A4").Type);   // Text
Console.WriteLine(ws.Cell("A5").Type);   // Empty
```

输出：

```
Number
Boolean
Date
Text
Empty
```

#### 5.4 可空类型

`SetValue` 接受 `object?`，可空值类型（`int?` / `DateTime?` 等）装箱后按底层值处理；`null` 写空单元格：

```csharp
int? maybe = null;
ws.SetValue("A1", maybe);        // 写空
maybe = 7;
ws.SetValue("A2", maybe);        // Number 7
Console.WriteLine(ws.Cell("A1").IsEmpty);     // True（null 写空）
Console.WriteLine(ws.Cell("A2").GetDouble()); // 7
```

输出：

```
True
7
```

#### 5.5 数字格式速查

`NumberFormat` 使用 Excel 格式代码字符串，常见示例：

| 格式代码 | 效果 |
|---|---|
| `"0"` | 整数 |
| `"0.00"` | 两位小数 |
| `"#,##0.00"` | 千分位 + 两位小数 |
| `"0%"` | 百分比 |
| `"yyyy/m/d"` / `"yyyy-MM-dd"` | 日期 |
| `"hh:mm"` | 时间 |
| `"@"` | 文本 |

```csharp
ws.Cell("A1").SetValue(12345.678);
ws.Cell("A1").NumberFormat = "#,##0.00";   // 显示 12,345.68
Console.WriteLine(ws.Cell("A1").GetString());  // 按格式读回
```

输出：

```
12,345.68
```

#### 5.6 读取时日期自动识别

读取 xlsx / xlsm / xlsb 时，单元格数字格式为 Excel 内置日期格式（ID 14-22、27-36、45-47、50-58 等）时，自动读为 `CellType.Date`：

```csharp
var opened = Excel.Open("dates.xlsx");
var cell = opened.Worksheets[0].Cell("A1");
if (cell.Type == CellType.Date)
    Console.WriteLine(cell.GetDateTime().ToString("yyyy-MM-dd"));
```

打开时捕获的 1904 日期系统标志（`Date1904`）会在保存时写回对应格式标志，保证日期序列值基准一致。

输出：

```
2024-01-01
```

#### 5.7 公式

`Cell.Formula` 与缓存值字段分离（公式串不再占用 `Text`，避免覆盖文本公式的缓存结果值）。旧代码把公式写进 `Text` 且设 `IsFormula=true` 的写法仍兼容：

```csharp
var cell = ws.Cell("C1");
cell.Formula = "SUM(A1:B1)";     // 仅写公式字符串，不计算结果
cell.IsFormula = true;           // 写出时按公式处理（兼容垫片）

// 或直接给单元格赋一个公式 Cell（SetValue 支持传 Cell，会复制其内容）
ws.Cell("C2").SetValue(Cell.FromFormula("A1*2"));
Console.WriteLine(cell.Formula);             // SUM(A1:B1)
Console.WriteLine(ws.Cell("C2").Formula);    // A1*2
```

输出：

```
SUM(A1:B1)
A1*2
```

List\<T\> 映射中可用 `[LiteColumn(IsFormula = true)]` 把字符串属性当作公式列（见第 5.9 节）。

#### 5.8 Byte[]

`SetValue` 遇到非数值类型一律按 `Text` 处理（`value.ToString()`）。**二进制数据请走图片 API**（`Worksheet.AddImage`，见第 13 章）或自行编码为文本。库本身不把 `byte[]` 映射为二进制单元格类型。

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", Convert.ToBase64String(new byte[] { 1, 2, 3 }));  // 自行编码为文本
Console.WriteLine(ws.Cell("A1").GetString());                       // AQID
```

输出：

```
AQID
```

#### 5.9 List\<T\> 映射与 `[LiteColumn]`

`[LiteColumn]` 特性控制 List\<T\> 映射时的列名 / 顺序 / 格式 / 忽略 / 公式：

```csharp
public class Person
{
    [LiteColumn(Name = "姓名", Order = 0)]
    public string Name { get; set; } = "";

    [LiteColumn(Name = "年龄", Order = 1, Format = "0")]
    public int Age { get; set; }

    [LiteColumn(Name = "总额", Order = 2, Format = "#,##0.00", IsFormula = true)]
    public string Total { get; set; } = "";   // 值可带或不带前导 "="

    [LiteColumn(Ignore = true)]
    public string Secret { get; set; } = "";  // 不输出
}
```

```csharp
var people = new List<Person> { new() { Name = "张三", Age = 30, Total = "=100*2", Secret = "隐藏" } };
var wb = Excel.Create(people, "People");   // 表头：姓名 | 年龄 | 总额；Secret 列忽略
wb.SaveAs("people.xlsx");
```

输出：

```
已写入 people.xlsx
```

List\<T\> 映射自动转换的 CLR 类型：

| CLR 类型 | 单元格类型 | 说明 |
|---|---|---|
| `int` / `long` / `short` / `byte` | `Number` | 整数 |
| `double` / `float` / `decimal` | `Number` | 小数 |
| `DateTime` | `Date` | 日期时间 |
| `bool` | `Boolean` | 布尔值 |
| `string` | `Text` | 文本 |

以上类型均支持可空版本（`int?` / `DateTime?` 等），null 写为空单元格。

#### 5.10 List\<T\> Fluent 配置（WriteOptions\<T\> / ReadOptions\<T\>）

除了 `[LiteColumn]` 特性，还支持 Fluent API 与字典映射，适合临时调整列名 / 格式 / 忽略 / 公式：

```csharp
// 写出时 Fluent 配置
Excel.Write("people.xlsx", people, "Employees", opt => opt
    .Column(p => p.Name, "姓名")                    // 指定列名
    .Column(p => p.Age, "年龄", format: "0")        // 指定列名 + 数字格式
    .Column(p => p.Total, "总额", isFormula: true)  // 公式列（值可带或不带前导 "="）
    .Ignore(p => p.Secret)                          // 忽略属性
);

// 读取时 Fluent 配置
var list = Excel.Read<Person>("people.xlsx", "Employees", opt => opt
    .Column(p => p.Name, "姓名")                    // 指定表头名 -> 属性映射
    .Column(p => p.Age, "年龄")
);

// 字典映射（老项目常见）；configure 用命名参数，sheetName 取默认 "Sheet1"
Excel.Write("people.xlsx", people, configure: opt => opt
    .Map(new Dictionary<string, string> { { "Name", "姓名" }, { "Age", "年龄" } })
);
```

输出：（直接输出一个people.xlsx文件）

#### 5.11 DataTable 便利 API

DataTable 自带列结构，**无需反射**（不触发反射映射），AOT 安全。首行自动写为列名：

```csharp
var dt = new DataTable("订单");
dt.Columns.Add("OrderID", typeof(int));
dt.Columns.Add("Customer", typeof(string));
dt.Columns.Add("Amount", typeof(decimal));
dt.Columns.Add("Date", typeof(DateTime));
dt.Rows.Add(1001, "Alice", 599.99m, new DateTime(2024, 6, 1));
dt.Rows.Add(1002, "Bob", 1299.50m, new DateTime(2024, 6, 15));

Excel.Write("orders.xlsx", dt, "Orders");   // 一步写出

var back = Excel.ReadAsDataTable("orders.xlsx", "Orders");   // 读回（首行为表头）
foreach (DataRow row in back.Rows)
    Console.WriteLine($"#{row["OrderID"]} | {row["Customer"]} | {row["Amount"]:0.00}");

var wb = Excel.Create(dt);        // 一步建簿；sheetName 缺省取 DataTable.TableName（再空则 Sheet1）
Console.WriteLine("sheet: " + wb.Worksheets[0].Name);
wb.SaveAs("orders2.xlsx");

var opened = Excel.Open("orders.xlsx");
// 导入到已有工作表：清空现有内容后从 A1 重建；includeHeader=false 不写列名行
opened.Worksheets[0].ImportData(dt, includeHeader: false);
opened.SaveAs("orders3.xlsx");
Console.WriteLine("imported rows: " + Excel.ReadAsDataTable("orders3.xlsx", "Orders", firstRowIsHeader: false).Rows.Count);
```

重要参数：

| API | 参数 | 类型 | 说明 |
|---|---|---|---|
| `Excel.Write` | `sheetName` | `string` | 目标工作表名，默认 `"Sheet1"` |
| `Excel.Write` | `options` | `ExcelWriteOptions?` | 写入选项（见 3.4） |
| `Excel.Create` | `sheetName` | `string?` | 空则取 `DataTable.TableName`，再空则 `"Sheet1"` |
| `Excel.ReadAsDataTable` | `sheetName` | `string?` | 目标工作表，null 读第一张表 |
| `Excel.ReadAsDataTable` | `firstRowIsHeader` | `bool` | 首行是否作为列名，默认 `true` |
| `ImportData` | `includeHeader` | `bool` | 是否写列名行，默认 `true`；导入会清空整表后从 A1 重建 |

输出：

```
#1001 | Alice | 599.99
#1002 | Bob | 1299.50
sheet: 订单
imported rows: 2
```

> ⚠️ DataTable 路径不经过反射映射（`[LiteColumn]` / Fluent 配置不适用）；首行始终为列名（`includeHeader=false` 除外）。

---

### 6. 样式（CellStyle / Border / 对齐 / 换行）

#### 6.1 单元格样式 `CellStyle`

颜色统一使用 `"#RRGGBB"` 格式：

```csharp
var style = new CellStyle
{
    FontName = "Arial",
    FontSize = 12,
    Bold = true,
    Italic = true,
    Underline = false,
    Strikeout = false,
    FontColor = "#FF0000",
    FillColor = "#FFFF00",
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Center,
    WrapText = true,
    NumberFormat = "#,##0.00",   // 用于 dxf 读回（超级表列格式 / 条件格式）
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
    },
};
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `FontName` | `string?` | 字体名，如 `"Arial"` |
| `FontSize` | `double` | 字号，默认 11 |
| `Bold` | `bool` | 加粗 |
| `Italic` | `bool` | 斜体 |
| `Underline` | `bool` | 下划线 |
| `Strikeout` | `bool` | 删除线 |
| `FontColor` | `string?` | 字体颜色，`"#RRGGBB"` 格式 |
| `FillColor` | `string?` | 填充颜色，`"#RRGGBB"` 格式 |
| `HorizontalAlignment` | `HorizontalAlignment` | 水平对齐，默认 `General` |
| `VerticalAlignment` | `VerticalAlignment` | 垂直对齐，默认 `Bottom` |
| `WrapText` | `bool` | 自动换行 |
| `NumberFormat` | `string?` | 数字格式代码（用于 dxf 读回） |
| `Border` | `BorderStyle?` | 边框 |

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("A1").Style = style;
Console.WriteLine(ws.Cell("A1").Style.Bold);   // True
```

输出：

```
True
```

#### 6.2 边框 `BorderStyle` / `BorderEdge`

```csharp
var style = new CellStyle
{
    Border = new BorderStyle
    {
        Top = new BorderEdge { Style = "thin", Color = "#000000" },
        Bottom = new BorderEdge { Style = "double" },
        Left = new BorderEdge { Style = "medium", Color = "#333333" },
        Right = new BorderEdge { Style = "dashed" },
    },
};
```

`BorderEdge.Style` 为字符串，常用值：`thin` / `medium` / `thick` / `double` / `dashed` / `dotted` 等。

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("A1").Style = style;
Console.WriteLine(ws.Cell("A1").Style.Border.Top.Style);   // thin
```

输出：

```
thin
```

#### 6.3 对齐与换行

```csharp
public enum HorizontalAlignment { General, Left, Center, Right }
public enum VerticalAlignment { Top, Center, Bottom }
```

`WrapText = true` 启用单元格内自动换行。

输出：（本示例无控制台输出）

#### 6.4 设置单元格 / 区域样式（对象模型 API）

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// 单格
ws.Cell("A1").Style = new CellStyle { Bold = true, FillColor = "#D9E1F2" };
ws.Cell("A1").NumberFormat = "yyyy/m/d";

// 区域统一样式
ws.Range("A1:C10").Style = new CellStyle { HorizontalAlignment = HorizontalAlignment.Center };
Console.WriteLine(ws.Cell("A1").NumberFormat);   // yyyy/m/d
```

输出：

```
yyyy/m/d
```

#### 6.5 表头样式 `HeaderStyle`

作用于 `SheetData.Headers` 表头行。

> ⚠️ **生效条件**：`HeaderStyle` 仅在有独立表头行（List\<T\> / DataTable 写入，或低层 `SheetData.Headers`（见附录 B.1））时生效。由于**对象模型是"整表网格"模型**（`ws.SetValue` 写入的所有行都算数据行，没有独立表头行），直接 `ws.HeaderStyle = ...` **不会**产生效果。对象模型下要给首行加样式，用 `RowStyles` 指定第 0 行：

```csharp
// 方式一：低层 / List<T> / DataTable 路径（有 Headers）
ws.HeaderStyle = new CellStyle { Bold = true, FillColor = "#4472C4", FontColor = "#FFFFFF" };
```

```csharp
// 方式二：对象模型网格（首行当表头 → 用 RowStyles 第 0 行）
ws.SetValue("A1", "Name"); ws.SetValue("B1", "Age");
ws.RowStyles = new Dictionary<int, CellStyle>
{
    { 0, new CellStyle { Bold = true, FillColor = "#4472C4", FontColor = "#FFFFFF" } },
};
Console.WriteLine(ws.RowStyles[0].Bold);   // True
```

输出：

```
True
```

#### 6.6 全表默认样式 `DefaultStyle`

优先级最低：

```csharp
ws.DefaultStyle = new CellStyle { FontName = "Consolas", FontSize = 10 };
Console.WriteLine(ws.DefaultStyle.FontName);   // Consolas
```

输出：

```
Consolas
```

#### 6.7 行级样式 `RowStyles`

key 为 **0-based 行索引**：

```csharp
ws.RowStyles = new Dictionary<int, CellStyle>
{
    { 1, new CellStyle { FillColor = "#FCE4D6" } },   // 第 2 行（0-based 1）
};
Console.WriteLine(ws.RowStyles[1].FillColor);   // #FCE4D6
```

输出：

```
#FCE4D6
```

#### 6.8 列级样式 `ColumnStyles`

key 为 **0-based 列索引**：

```csharp
ws.ColumnStyles = new Dictionary<int, CellStyle>
{
    { 2, new CellStyle { HorizontalAlignment = HorizontalAlignment.Right } },  // 第 3 列
};
Console.WriteLine(ws.ColumnStyles[2].HorizontalAlignment);   // Right
```

输出：

```
Right
```

#### 6.9 样式优先级（覆盖式）

写出时按如下优先级解析（**行列级样式优先级更明确**）：

- **数据行**：`Cell.Style` > `RowStyle` > `ColumnStyle` > `DefaultStyle`
- **表头行**：`HeaderStyle` > `ColumnStyle` > `DefaultStyle`

```csharp
// 示例：单元格样式覆盖行样式，行样式覆盖列样式，列样式覆盖默认样式
ws.DefaultStyle = new CellStyle { FontSize = 10 };
ws.ColumnStyles = new Dictionary<int, CellStyle> { { 0, new CellStyle { Bold = true } } };
ws.RowStyles = new Dictionary<int, CellStyle> { { 0, new CellStyle { Italic = true } } };
ws.Cell("A1").Style = new CellStyle { Underline = true };
// A1 最终：Underline（单元格） + Italic（行） + Bold（列） + FontSize 10（默认）
```

输出：（本示例无控制台输出）

---

### 7. 合并单元格

#### 7.1 写出合并

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.SetValue("A1", "Merged Title");
ws.Merge("A1:D1");                    // A1 地址
ws.Merge(2, 1, 2, 3);                 // 1-based 行列（合并 A2:C2）
ws.Merge("A3:B4");

// 区域方式
ws.Range("C5:E5").Merge();
Console.WriteLine(ws.MergedRanges.Count);   // 4
```

输出：

```
4
```

#### 7.2 取消合并

```csharp
ws.Unmerge("A1:D1");
ws.Range("C5:E5").Unmerge();
```

输出：（本示例无控制台输出）

#### 7.3 读取合并区域

`Worksheet.MergedRanges` 返回 `IReadOnlyList<CellRange>`（**0-based**，与低层模型一致（见附录 B.1））：

```csharp
var opened = Excel.Open("merged.xlsx");
var ws = opened.Worksheets[0];
foreach (var m in ws.MergedRanges)
    Console.WriteLine($"{m.FirstRow},{m.FirstCol} - {m.LastRow},{m.LastCol}");
```

输出：

```
0,0 - 0,3
1,0 - 1,2
2,0 - 3,1
4,2 - 4,4
```

#### 7.4 合并区域填充

读取时设置 `FillMergedCells = true` 会把左上角的值展开到整个合并区域：

```csharp
var wb = Excel.Open("merged.xlsx", new ExcelReadOptions { FillMergedCells = true });
// 合并区非左上角单元格现在也有值
```

输出：（本示例无控制台输出）

---

### 8. 自动筛选

#### 8.1 写出筛选

`Worksheet.Filter` 为 `AutoFilter` 对象，含 `Range`、每列条件 `Columns` 与 `HiddenRows`。Excel 的筛选区域**第一行始终是表头**——所以首先写入表头行，数据从第 2 行开始：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// 表头行（筛选区域的第 1 行）
ws.SetValue("A1", "Name");
ws.SetValue("B1", "Type");
ws.SetValue("C1", "Score");

// 数据从第 2 行开始，共 541 行 → 区域 A1:C542
for (int r = 2; r <= 542; r++)
{
    ws.SetValue(r, 1, $"Name {r - 1}");
    ws.SetValue(r, 2, r % 2 == 0 ? "Active" : "Inactive");
    ws.SetValue(r, 3, (r - 1) * 10);
}

ws.Filter = new AutoFilter
{
    Range = "A1:C542",
    Columns = new List<FilterColumn>
    {
        new FilterColumn
        {
            ColumnIndex = 1,                 // 0-based 列索引（第 2 列）
            Type = FilterType.Equals,
            Values = new List<string> { "Active" },
        },
    },
};
```

`AutoFilter` 关键成员：

| 参数 | 类型 | 说明 |
|---|---|---|
| `Range` | `string` | 筛选区域 A1 风格引用，首行为表头 |
| `Columns` | `List<FilterColumn>` | 每列筛选条件，0 基列索引 |
| `HiddenRows` | `HashSet<int>` | 手动隐藏的 0 基行索引集合（可选，见 8.6） |

`FilterColumn` 关键成员：

| 参数 | 类型 | 说明 |
|---|---|---|
| `ColumnIndex` | `int` | 0 基列索引（第 1 列 = 0） |
| `Type` | `FilterType` | 条件类型，见 8.2 |
| `Values` | `List<string>` | 匹配值集合 |
| `Operator` | `FilterOperator` | `Type = Compare` 时的比较操作符，见 8.3 |
| `MinValue` / `MaxValue` | `string?` | `Between` 的下 / 上界 |

输出：（本示例无控制台输出）

> ⚠️ 不要把数据写进筛选区域的第一行——Excel 会把该行当作表头（显示筛选箭头），数据从第 2 行起才能正确参与筛选。

#### 8.2 筛选条件类型 `FilterType`

`FilterColumn.Type` 指定某列的筛选条件类型，枚举取值如下：

```csharp
public enum FilterType { Equals, Compare, Contains, BeginsWith, EndsWith, Blank }
```

输出：（本示例无控制台输出）

#### 8.3 Compare 操作符 `FilterOperator`

当 `FilterColumn.Type = FilterType.Compare` 时，用 `Operator` 指定比较关系：

```csharp
public enum FilterOperator { GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, Between }
```

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.GreaterThan,
    Values = new List<string> { "500" },
};
```

输出：（本示例无控制台输出）

#### 8.4 Between 示例

区间筛选用 `Operator = FilterOperator.Between` 并配合 `MinValue` / `MaxValue`：

```csharp
new FilterColumn
{
    ColumnIndex = 2,
    Type = FilterType.Compare,
    Operator = FilterOperator.Between,
    MinValue = "100",
    MaxValue = "500",
};
```

输出：（本示例无控制台输出）

#### 8.5 多条件（AND 逻辑）

多个 `FilterColumn` 同时生效（同一行需满足所有列条件）：

```csharp
ws.Filter = new AutoFilter
{
    Range = "A1:D542",
    Columns = new List<FilterColumn>
    {
        new FilterColumn { ColumnIndex = 1, Type = FilterType.Equals, Values = new() { "Active" } },
        new FilterColumn { ColumnIndex = 2, Type = FilterType.Compare, Operator = FilterOperator.GreaterThan, Values = new() { "500" } },
    },
};
```

输出：（本示例无控制台输出）

#### 8.6 手动指定隐藏行

`HiddenRows` 为 0-based 行索引集合（相对 `Rows`）：

```csharp
ws.Filter = new AutoFilter
{
    Range = "A1:D542",
    HiddenRows = new HashSet<int> { 1, 3, 5 },   // 隐藏第 2、4、6 行
};
```

输出：（本示例无控制台输出）

#### 8.7 读取筛选

打开文件后通过 `Worksheet.Filter` 读取筛选区域与各列条件：

```csharp
var opened = Excel.Open("filtered.xlsx");
var filter = opened.Worksheets[0].Filter;
if (filter is not null)
{
    Console.WriteLine(filter.Range);
    foreach (var col in filter.Columns)
        Console.WriteLine($"{col.ColumnIndex}: {col.Type} {string.Join(",", col.Values)}");
}
```

输出：

```
A1:C542
1: Equals Active
```
---

### 9. 行高与列宽（含 AutoColumnWidths）

#### 9.1 设置行高

`Worksheet.RowHeights` 为 `Dictionary<int, double>`，key 为 **0-based 行索引**，单位 **磅（point）**：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Tall row");
ws.RowHeights = new Dictionary<int, double> { { 0, 30.0 } };   // 第 1 行高 30 磅
```

输出：（本示例无控制台输出）

#### 9.2 设置列宽

`Worksheet.ColumnWidths` 为 `Dictionary<int, double>`，key 为 **0-based 列索引**：

```csharp
ws.ColumnWidths = new Dictionary<int, double>
{
    { 0, 20.0 },
    { 1, 15.0 },
};
```

输出：（本示例无控制台输出）

#### 9.3 列宽自适应 `AutoColumnWidths`

`Worksheet.AutoColumnWidths()` 按表内现有内容估算每列宽度（中文字符算 2，英文 / 数字算 1，范围 `[8, 50]`），结果写入 `ColumnWidths`：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Name");
ws.SetValue("A2", "Zhang San");
ws.SetValue("B1", "Description");
ws.SetValue("B2", "A very long description that should widen the column");
ws.AutoColumnWidths();
// 现在 ws.ColumnWidths 已按内容估算
```

读回验证：

```csharp
var opened = Excel.Open("autowidth.xlsx");
var widths = opened.Worksheets[0].ColumnWidths;
if (widths is not null)
    foreach (var kv in widths)
        Console.WriteLine($"Col {kv.Key}: {kv.Value:F1}");
```

输出：

```
Col 0: 9.0
Col 1: 50.0
```

> ⚠️ `AutoColumnWidths` 为估算值（中文字符按 2、英文/数字按 1，钳制在 `[8, 50]`），与 Excel 实际渲染宽度可能有细微差异。

#### 9.4 写出时自动适配

`Excel.Write` 的 `ExcelWriteOptions.AutoFitColumns = true` 会在写出前对每张表自动估算列宽：

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", "自动适配列宽");
Excel.Write("out.xlsx", wb, new ExcelWriteOptions { AutoFitColumns = true });
```

输出：已写入 out.xlsx

---

### 10. 批注

#### 10.1 写出批注

`Worksheet.Comments` 为 `Dictionary<string, string>`，key 为 **A1 格式单元格引用**，value 为批注文本：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string>
{
    { "A1", "This is a comment on A1" },
    { "B1", "Note for B1 <with special chars>" },
};
```

输出：（本示例无控制台输出）

> ⚠️ 批注仅支持 xlsx / xlsm；写出到 xls / xlsb / csv 时按降级机制丢弃（见第 22 章）。批注写回依赖 OOXML VML legacyDrawing，需用真实 Excel 打开验证。

#### 10.2 读回批注

```csharp
var opened = Excel.Open("comments.xlsx");
var comments = opened.Worksheets[0].Comments;
if (comments is not null && comments.TryGetValue("A1", out var text))
    Console.WriteLine(text);   // 输出: This is a comment on A1
```

输出：

```
This is a comment on A1
```

#### 10.3 对象模型 API：给指定单元格加 / 读批注

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.Cell("C5").SetValue("data");

// 加批注
ws.Comments ??= new Dictionary<string, string>();
ws.Comments["C5"] = "审核通过";

// 读批注
var opened = Excel.Open("comments2.xlsx");
string? note = null;
opened.Worksheets[0].Comments?.TryGetValue("C5", out note);
Console.WriteLine(note);   // 输出: 审核通过
```

输出：

```
审核通过
```

---

### 11. 超链接（外部 / 内部）

#### 11.1 写出超链接

`Cell.Hyperlink` 为 `Hyperlink` 对象，支持外部链接（URL / 文件路径）与工作簿内部跳转：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// 外部链接
ws.Cell("A1").SetValue("Example");
ws.Cell("A1").Hyperlink = new Hyperlink
{
    Target = "https://example.com",
    Tooltip = "Visit Example",
    IsInternal = false,
};

// 内部跳转（Target 以 '#' 开头）
ws.Cell("B1").SetValue("Go to Sheet2");
ws.Cell("B1").Hyperlink = new Hyperlink
{
    Target = "#Sheet2!A1",
    IsInternal = true,
};
```

输出：（本示例无控制台输出）

#### 11.2 Hyperlink 属性

`Cell.Hyperlink` 为 `Hyperlink` 对象，成员如下：

- `Target`：链接目标。内部链接格式如 `#SheetName!A1`；外部为完整 URL 或文件路径。
- `Tooltip`：鼠标悬停提示文本（可选）。
- `IsInternal`：是否工作簿内部跳转。

输出：（本示例无控制台输出）

#### 11.3 读回超链接

打开文件后通过 `Cell.Hyperlink` 读取超链接信息：

```csharp
var opened = Excel.Open("links.xlsx");
var cell = opened.Worksheets[0].Cell("A1");
if (cell.Hyperlink is { } h)
    Console.WriteLine($"{h.Target} internal={h.IsInternal} tooltip={h.Tooltip}");
```

输出：

```
https://example.com internal=False tooltip=Visit Example
```

超链接在 xlsx / xlsm / xlsb / xls 四格式读写均支持。

---

### 12. 冻结窗格

#### 12.1 设置冻结行 / 列

`Worksheet.FreezeRows` / `FreezeColumns` 为 1-based 冻结行 / 列数（0 = 不冻结）：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.FreezeRows = 2;       // 冻结前 2 行
ws.FreezeColumns = 3;    // 冻结前 3 列
```

输出：（本示例无控制台输出）

#### 12.2 FreezeHeader 兼容

`FreezeHeader = true` 等价于 `FreezeRows = 1`：

```csharp
ws.FreezeHeader = true;   // 冻结首行
```

输出：（本示例无控制台输出）

#### 12.3 对象模型 API

通过属性直接设置冻结行列（等价于 12.1）：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.FreezeRows = 1;
ws.FreezeColumns = 1;
```

输出：（本示例无控制台输出）

#### 12.4 读回冻结

打开文件后读取 `FreezeRows` / `FreezeColumns`：

```csharp
var opened = Excel.Open("frozen.xlsx");
var ws = opened.Worksheets[0];
Console.WriteLine($"{ws.FreezeRows} rows, {ws.FreezeColumns} cols");
```

输出：

```
2 rows, 3 cols
```

冻结窗格在 xlsx / xlsb / xls 三格式任意行列冻结均支持。

---

### 13. 图片（Floating / InCell / 读回）

图片仅支持 xlsx / xlsm。`Worksheet.AddImage` 添加图片，`Worksheet.Images` 读回。

#### 13.1 浮动图片（Floating）

以 `row/column` 左上角为锚点，默认按图片原始尺寸显示：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
byte[] png = File.ReadAllBytes("logo.png");

ws.AddImage(png, 1, 1);                            // 锚点 A1，原始尺寸
ws.AddImage(png, 1, 3, 120, 60);                   // 指定显示尺寸（像素）
```

`AddImage`（row/column 重载）重要参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `data` | `byte[]` | 图片二进制（PNG/JPEG/GIF/BMP） |
| `row` / `column` | `int` | 1 基左上角锚点行列 |
| `widthPx` / `heightPx` | `double?` | 显示尺寸（像素），null = 图片原始尺寸 |
| `placement` | `ImagePlacement` | `Floating` / `InCell`，默认 `Floating` |

输出：（本示例无控制台输出）

> ⚠️ 图片仅支持 xlsx / xlsm（见第 20 章格式支持矩阵）。

#### 13.2 单元格内嵌图片（InCell）

Excel 365 InCell 图片（richData 体系）：

```csharp
ws.AddImage(png, 2, 1, placement: ImagePlacement.InCell);
```

输出：（本示例无控制台输出）

> ⚠️ InCell 图片基于 Excel 365 richData 体系（写回为 richData 部件），老版本 Excel 可能无法识别。

#### 13.3 图片放置枚举 `ImagePlacement`

`ImagePlacement` 决定图片是嵌入单元格还是浮动：

```csharp
public enum ImagePlacement { InCell, Floating }
```

输出：（本示例无控制台输出）

#### 13.4 高精度锚点 `ImageAnchor` 与移动方式 `ImageMoveMode`

`ImageAnchor` 提供左上单元格 + EMU 偏移 + 显示尺寸 + 移动方式，写回时优先于 `Row`/`Column`：

```csharp
var anchor = new ImageAnchor
{
    TopLeftCell = "B2",
    TopLeftOffsetX = 100,        // EMU 偏移（1px ≈ 9525）
    TopLeftOffsetY = 50,
    WidthPixels = 200,
    HeightPixels = 100,
    MoveMode = ImageMoveMode.MoveAndSizeWithCells,
};
ws.AddImage(png, anchor, name: "Chart", altText: "季度图表");
```

```csharp
public enum ImageMoveMode
{
    MoveAndSizeWithCells,        // 随单元格移动并缩放
    MoveButDontSizeWithCells,    // 随单元格移动但不缩放（默认）
    FixedPosition,               // 固定位置
}
```

`ImageAnchor` 重要参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `TopLeftCell` | `string` | 左上单元格 A1 引用 |
| `TopLeftOffsetX` / `TopLeftOffsetY` | `int` | 左上偏移（EMU，1px≈9525） |
| `WidthPixels` / `HeightPixels` | `double` | 显示尺寸（像素） |
| `MoveMode` | `ImageMoveMode` | 移动 / 缩放方式，默认 `MoveButDontSizeWithCells` |

`AddImage`（anchor 重载）重要参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `data` | `byte[]` | 图片二进制 |
| `anchor` | `ImageAnchor` | 高精度锚点 |
| `name` | `string?` | 图片名称（可选） |
| `altText` | `string?` | 无障碍替换文本（可选） |

输出：（本示例无控制台输出）

> ⚠️ `ImageAnchor` 仅对 Floating 图片生效；InCell 请用 row/column 重载（`Anchor` 会被忽略）。

#### 13.5 图片读回

打开含图片的文件后，`Worksheet.Images` 自动填充：

```csharp
var opened = Excel.Open("with_images.xlsx");
foreach (var img in opened.Worksheets[0].Images)
{
    Console.WriteLine($"{img.CellAddress} {img.Placement} {img.Data.Length} bytes");
    // img.Data 为原始图片字节，img.Row/Column 为锚点，img.Extension 为扩展名
}
```

输出：

```
A1 Floating 70 bytes
C1 Floating 70 bytes
A2 InCell 70 bytes
```

`WorksheetImage` 关键成员：`Data`（字节）、`Extension`、`Row`/`Column`（1-based 锚点）、`Placement`、`WidthPx`/`HeightPx`、`Name`、`Anchor`、`AltText`、`CellAddress`（只读 A1 引用）。

#### 13.6 多 Sheet 混合使用

不同工作表可分别使用 Floating 与 InCell 放置，互不干扰：

```csharp
byte[] img = File.ReadAllBytes("logo.png");
var wb = Excel.Create();

var wsBanner = wb.Worksheets[0];
wsBanner.Name = "Banner";
wsBanner.AddImage(img, 1, 1);                                 // 浮动图片

var wsEmbed = wb.Worksheets.Add("Embed");
wsEmbed.AddImage(img, 1, 1, placement: ImagePlacement.InCell); // 单元格内嵌图片

wb.SaveAs("multi_images.xlsx");

// 读回验证
var opened = Excel.Open("multi_images.xlsx");
foreach (var s in opened.Worksheets)
    foreach (var im in s.Images)
        Console.WriteLine($"{s.Name}: {im.CellAddress} {im.Placement} {im.Data.Length} bytes");
```

输出：

```
Banner: A1 Floating 70 bytes
Embed: A1 InCell 70 bytes
```

> ⚠️ 图片仅支持 xlsx / xlsm（见第 20 章格式支持矩阵）。InCell 图片老版本 Excel 可能无法识别。
---

### 14. 数据验证

#### 14.1 写出数据验证

`Worksheet.Validations` 为 `List<DataValidation>`，通过对象初始化器逐条配置验证规则；数据验证仅 xlsx / xlsm 写出（其余格式经降级上报丢弃）：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.Validations = new List<DataValidation>
{
    // 下拉列表（逗号分隔，用引号包裹）
    new DataValidation
    {
        Type = DataValidationType.List,
        Sqref = "A1:A10",
        Formula1 = "\"Active,Inactive,Pending\"",
        AllowBlank = true,
        PromptTitle = "请选择",
        Prompt = "从下拉列表选择状态",
    },
    // 整数区间验证
    new DataValidation
    {
        Type = DataValidationType.WholeNumber,
        Sqref = "B1:B10",
        Formula1 = "1",
        Formula2 = "100",
    },
};
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `Type` | `DataValidationType` | 验证类型（见 14.2） |
| `Sqref` | `string` | 应用范围（A1 风格，如 `A1:A10`） |
| `Formula1` | `string` | 列表验证为引号包裹的逗号分隔项；区间验证为下限 |
| `Formula2` | `string?` | 区间验证上限（非区间可省略） |
| `AllowBlank` | `bool` | 是否允许空值（默认 false） |
| `PromptTitle` / `Prompt` | `string?` | 选中单元格时的输入提示标题 / 正文 |

输出：（本示例无控制台输出）

#### 14.2 数据验证类型 `DataValidationType`

`DataValidationType` 决定验证规则类别，配合 `Formula1` / `Formula2` 使用：

```csharp
public enum DataValidationType { List, WholeNumber, Decimal, Date }
```

- `List`：下拉列表验证，`Formula1` 用引号包裹的逗号分隔列表。
- `WholeNumber` / `Decimal` / `Date`：数值 / 日期验证，`Formula1` 为下限、`Formula2` 为上限（区间验证）。

输出：（本示例无控制台输出）

#### 14.3 读回数据验证

打开含数据验证的文件后，`Worksheet.Validations` 自动回填，遍历即可打印每条规则：

```csharp
var opened = Excel.Open("validations.xlsx");
var validations = opened.Worksheets[0].Validations;
if (validations is not null)
    foreach (var v in validations)
        Console.WriteLine($"{v.Type} {v.Sqref} {v.Formula1} {v.Formula2}");
```

输出：

```
List A1:A10 Active,Inactive,Pending 
WholeNumber B1:B10 1 100
```

---

### 15. 条件格式（cellIs / expression / colorScale / dataBar / 长尾 / iconSet）

条件格式在 xlsx / xlsm 读写。`Worksheet.ConditionalFormats` 为 `List<ConditionalFormat>`。

#### 15.1 单元格值比较（cellIs）

`ConditionalFormatType.CellIs` 按 `ConditionalOperator` 与固定值比较，命中后套用 `Style`：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "B2:B10",
    Type = ConditionalFormatType.CellIs,
    Operator = ConditionalOperator.GreaterThan,
    Formula = "100",
    Style = new CellStyle { Bold = true, FillColor = "#FFC7CE" },
});
```

`ConditionalOperator`：`LessThan` / `LessThanOrEqual` / `Equal` / `NotEqual` / `GreaterThan` / `GreaterThanOrEqual` / `Between` / `NotBetween`。Between 用 `Formula`（下限）+ `Formula2`（上限）。

`ConditionalFormat` 通用成员：

| 参数 | 类型 | 说明 |
|---|---|---|
| `Sqref` | `string` | 应用范围（A1 风格，可含多个区域，如 `A1:A100 D2:D9`） |
| `Type` | `ConditionalFormatType` | 规则类型（见本章各子节） |
| `Operator` | `ConditionalOperator` | 仅 `CellIs` 有效（默认 `GreaterThan`） |
| `Formula` / `Formula2` | `string?` | 比较目标 / Between 上限 |
| `Style` | `CellStyle?` | 命中时的样式（字体 / 填充 / 边框，不含对齐与数字格式） |

输出：（本示例无控制台输出）

#### 15.2 公式条件（expression）

`ConditionalFormatType.Expression` 用返回 TRUE / FALSE 的公式判定，公式用相对引用（当前单元格起算）：

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "A1:A100",
    Type = ConditionalFormatType.Expression,
    Formula = "MOD(ROW(),2)=0",     // 偶数行高亮
    Style = new CellStyle { FillColor = "#D9E1F2" },
});
```

输出：（本示例无控制台输出）

#### 15.3 色阶（colorScale）

`ConditionalFormatType.ColorScale` 按数值高低在低 / 高色间渐变，`MidColor` 非空时变为 3 色刻度：

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "C2:C10",
    Type = ConditionalFormatType.ColorScale,
    ColorScale = new ColorScaleInfo
    {
        LowColor = "F8696B",
        HighColor = "63BE7B",
        MidColor = "FFEB84",        // 设 nonNull 时启用 3 色刻度
    },
});
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `LowColor` | `string` | 低值颜色（`#RRGGBB` 或 `RRGGBB`） |
| `HighColor` | `string` | 高值颜色 |
| `MidColor` | `string?` | 中间色；非空时为 3 色刻度，否则 2 色 |

输出：（本示例无控制台输出）

#### 15.4 数据条（dataBar）

`ConditionalFormatType.DataBar` 在单元格内绘制与数值成比例的条形：

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "D2:D10",
    Type = ConditionalFormatType.DataBar,
    DataBar = new DataBarInfo
    {
        Color = "638EC6",
        ShowValue = true,
        MinLengthPercent = 0,
        MaxLengthPercent = 100,
    },
});
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `Color` | `string` | 条形颜色（默认 Excel 蓝 `638EC6`） |
| `ShowValue` | `bool` | 是否同时显示数值（默认 true；false 只显示条形） |
| `MinLengthPercent` / `MaxLengthPercent` | `int` | 最短 / 最长条形长度百分比（0–100） |

输出：（本示例无控制台输出）

#### 15.5 长尾文本 / 空值 / 错误 / 重复 / 前 N / 平均线

以下各类型为条件格式的长尾能力，命中规则与专用属性见注释：

```csharp
// 包含指定文本
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "A2:A100",
    Type = ConditionalFormatType.ContainsText,
    Text = "urgent",
    Style = new CellStyle { Bold = true },
});

// 以指定文本开头 / 结尾 / 不包含
// Type = BeginsWith / EndsWith / NotContainsText，同样用 Text

// 文本长度比较
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "B2:B100",
    Type = ConditionalFormatType.TextLength,
    Operator = ConditionalOperator.GreaterThan,
    Formula = "10",
});

// 时间周期
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "C2:C100",
    Type = ConditionalFormatType.TimePeriod,
    TimePeriod = "today",   // yesterday/today/tomorrow/lastWeek/thisWeek/nextWeek/lastMonth/thisMonth/nextMonth
});

// 空 / 非空 / 错误 / 非错误
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "D2:D100", Type = ConditionalFormatType.Blanks });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "D2:D100", Type = ConditionalFormatType.NoBlanks });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "E2:E100", Type = ConditionalFormatType.Errors });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "E2:E100", Type = ConditionalFormatType.NoErrors });

// 唯一 / 重复值
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "F2:F100", Type = ConditionalFormatType.Unique });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "F2:F100", Type = ConditionalFormatType.Duplicate });

// 前 N 项 / 前 N%
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "G2:G100",
    Type = ConditionalFormatType.Top10,
    Rank = 10,
    Percent = false,
});

// 高于 / 低于平均
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "H2:H100", Type = ConditionalFormatType.AboveAverage });
ws.ConditionalFormats.Add(new ConditionalFormat { Sqref = "H2:H100", Type = ConditionalFormatType.BelowAverage });
```

输出：（本示例无控制台输出）

#### 15.6 图标集（iconSet）

`IconSetInfo` 提供 17 个内置集合枚举 + 任意集合名 + 阈值：

```csharp
ws.ConditionalFormats.Add(new ConditionalFormat
{
    Sqref = "I2:I100",
    Type = ConditionalFormatType.IconSet,
    IconSet = new IconSetInfo
    {
        Style = IconSetStyle.ThreeArrows,       // 默认三色箭头
        Percent = true,
        ShowValue = true,
        // Thresholds 为空时按图标数均分百分比
        // Thresholds = new double[] { 33, 66 },
    },
});
```

`IconSetStyle` 枚举（17 种）：`ThreeArrows` / `ThreeArrowsGray` / `ThreeFlags` / `ThreeTrafficLights` / `ThreeTrafficLights2` / `ThreeSigns` / `ThreeSymbols` / `ThreeSymbols2` / `FourArrows` / `FourArrowsGray` / `FourRedToBlack` / `FourRating` / `FourTrafficLights` / `FiveArrows` / `FiveArrowsGray` / `FiveRating` / `FiveQuarters`。也可用 `CustomStyleName` 指定任意集合名字符串（非空时优先生效）。

| 参数 | 类型 | 说明 |
|---|---|---|
| `Style` | `IconSetStyle` | 内置集合（默认 `ThreeArrows`） |
| `CustomStyleName` | `string?` | 任意集合名字符串；非空时优先生效 |
| `Percent` | `bool` | 阈值按百分比（true）还是绝对数值（false），默认 true |
| `ShowValue` | `bool` | 单元格内是否同显数值，默认 true |
| `Thresholds` | `double[]?` | 自定义阈值（图标数 - 1 个，升序）；为空则按图标数均分 |

输出：（本示例无控制台输出）

#### 15.7 读回条件格式

打开含条件格式的文件后，`Worksheet.ConditionalFormats` 自动回填，遍历即可打印每条规则：

```csharp
var opened = Excel.Open("cf.xlsx");
var cfs = opened.Worksheets[0].ConditionalFormats;
foreach (var cf in cfs)
    Console.WriteLine($"{cf.Type} {cf.Sqref} {cf.Formula}");
```

输出：

```
CellIs B2:B10 100
Expression A1:A100 MOD(ROW(),2)=0
ColorScale C2:C10 
DataBar D2:D10 
IconSet I2:I100 
```

`ConditionalFormat` 其他成员：`Priority`（优先级，默认按注册顺序自动编号）、`Style`（条件满足时样式，不包含对齐与数字格式）。

---

### 16. 超级表（Table / ListObject，样式枚举 + 任意样式名 + 列格式）

超级表在 xlsx / xlsm 读写。`Worksheet.AddTable` 创建，`Worksheet.Tables` 读回。

#### 16.1 创建超级表

覆盖区首行作为表头列名，至少需要表头 + 1 行数据：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];

// 先写表头 + 数据
ws.SetValue("A1", "Product");
ws.SetValue("B1", "Price");
ws.SetValue("A2", "Apple");
ws.SetValue("B2", 3.5);
ws.SetValue("A3", "Banana");
ws.SetValue("B3", 2.0);

var table = ws.AddTable("A1:B3", "Products");     // 默认 Medium9 样式
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `refAddress` | `string` | 表覆盖区域（A1 风格，首行恒为表头） |
| `name` | `string` | 表名（全簿唯一；允许中文，不能以数字开头、不能含空格、不能撞单元格地址） |
| `style` | `TableStyleStyle?` | 内置样式枚举；缺省默认 `Medium9`（另有 `string styleName` 重载，见 16.3） |

返回 `XlTable`，可直接对其设置列格式等属性。

输出：（本示例无控制台输出）

#### 16.2 样式枚举 `TableStyleStyle`

`TableStyleStyle` 枚举内置 60 个条纹名（Light 1-21 / Medium 1-28 / Dark 1-11）+ `None`；样式外观由 Excel 内置渲染，文件仅保存样式名：

```csharp
var table = ws.AddTable("A1:B3", "Products", TableStyleStyle.Medium2);
```

输出：（本示例无控制台输出）

#### 16.3 任意样式名 `CustomStyleName`

`string` 重载 `AddTable(ref, name, styleName)` 可传任意样式名字符串（含 Excel 未来新增样式名）：

```csharp
var table = ws.AddTable("A1:B3", "Products", "TableStyleMedium9");
// 不在 60 个内置名内时 Excel 打开退化为无样式（经 OnDegradation 上报）
```

输出：（本示例无控制台输出）

> ⚠️ 样式名不在 60 个内置名内时，Excel 打开会静默退化为无样式，经 `OnDegradation` 回调上报（见第 22 章）。

#### 16.4 表属性

`XlTable` 成员：`Name`（全簿唯一，允许中文，不能以数字开头、不能含空格、不能撞单元格地址）、`Ref`、`Style`、`CustomStyleName`、`ShowRowStripes`（默认 true）、`ShowFirstColumn`、`ShowLastColumn`、`ShowColumnStripes`、`AutoFilter`（默认 true）、`TotalsRowShown`（读回保留）、`HeaderStyle`、`Columns`。返回的 `XlTable` 可直接读取这些属性：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
ws.SetValue("A1", "Product");
ws.SetValue("B1", "Price");
ws.SetValue("A2", "Apple");
ws.SetValue("B2", 3.5);

var table = ws.AddTable("A1:B2", "Products", TableStyleStyle.Medium2);
Console.WriteLine($"{table.Name} {table.Ref} {table.Style} 行条纹={table.ShowRowStripes} 筛选={table.AutoFilter}");
```

输出：

```text
Products A1:B2 Medium2 行条纹=True 筛选=True
```

#### 16.5 列格式（`XlTableColumn`）

`table.Column(name)` 按列名取列（大小写不敏感），设置 `Style`（font/fill/border → dxf）与 `NumberFormat`：

```csharp
var table = ws.AddTable("A1:B3", "Products");
table.Column("Price").NumberFormat = "#,##0.00";
table.Column("Price").Style = new CellStyle { Bold = true };
```

| 参数 | 类型 | 说明 |
|---|---|---|
| `Column(name)` 的 `name` | `string` | 列名（= 表头单元格文本，大小写不敏感；不存在抛 `LiteExcelException`） |
| `NumberFormat` | `string?` | 该列数字格式（如 `"#,##0.00"`） |
| `Style` | `CellStyle?` | 该列样式（字体 / 填充 / 边框，写出映射到 dxf） |

输出：（本示例无控制台输出）

#### 16.6 删除超级表

`RemoveTable(name)` 按表名（大小写不敏感）删除超级表，存在则返回 `true`，否则 `false`：

```csharp
bool removed = ws.RemoveTable("Products");   // 存在则删除并返回 true
```

输出：（本示例无控制台输出）

#### 16.7 读回超级表

打开含超级表的文件后，`Worksheet.Tables` 自动回填（含样式、列格式），遍历即可打印：

```csharp
var opened = Excel.Open("tables.xlsx");
foreach (var t in opened.Worksheets[0].Tables)
{
    Console.WriteLine($"{t.Name} {t.Ref} 样式={t.CustomStyleName ?? t.Style.ToString()}");
    foreach (var col in t.Columns)
        Console.WriteLine($"  {col.Name} fmt={col.NumberFormat}");
}
```

输出：

```
Products A1:B3 样式=TableStyleMedium2
  Product fmt=
  Price fmt=#,##0.00
```

---

### 17. 命名区域

> ⚠️ 命名区域仅支持 **xlsx / xlsm**（从 `workbook.xml` 的 `definedNames` 读回）。xls / xlsb 未实现：写出时命名区域会**静默丢失**，经 `OnDegradation` 上报。

#### 17.1 读回命名区域

打开含命名区域的文件后，`Workbook.Names` 自动填充（全局 + sheet-local），遍历即可打印：

```csharp
var opened = Excel.Open("names.xlsx");
foreach (var nr in opened.Names)
    Console.WriteLine($"{nr.Name} = {nr.Reference} local={nr.LocalSheetId}");
```

`NamedRange` 成员：`Name`、`Reference`（如 `Sheet1!$A$1:$C$9`）、`LocalSheetId`（-1 表示全局名称）、`IsLocalSheet`。

输出：

```
MyRange = Sheet1!$A$1:$C$9 local=-1
LocalRange = Sheet1!$B$2 local=0
```

#### 17.2 写出保留

命名区域在打开后保存时**原样保留**（`workbook.xml` 的 `definedNames` 透传），不会因编辑丢失：

```csharp
var opened = Excel.Open("names.xlsx");
opened.Worksheets[0].SetValue("A1", "edited");
opened.Save();   // 命名区域仍保留
```

输出：已写入 names.xlsx

---

### 18. 文件级密码（打开 / 修改）

文件级安全通过 `Workbook.Security`（`WorkbookSecurity`）管理，支持 xlsx / xlsm / xlsb。密码本体仅存储于对象内部，不对外暴露明文。

#### 18.1 打开加密文件

打开密码（Agile 加密）在 `ExcelReadOptions.OpenPassword` 提供；未提供时若文件已加密，抛出明确异常：

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions { OpenPassword = "secret" });
```

输出：（本示例无控制台输出）

> ⚠️ 打开 / 修改密码仅支持 xlsx / xlsm / xlsb；打开加密文件时若未提供密码（或错误），读取会抛异常。

#### 18.2 读取安全状态

`Workbook.Security`（`WorkbookSecurity`）暴露只读安全状态属性，配合打开选项读取：

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",
    ModifyPassword = "write",
});

var sec = wb.Security;
Console.WriteLine(sec.HasOpenPassword);      // true
Console.WriteLine(sec.HasModifyPassword);    // true
Console.WriteLine(sec.HasModifyAccess);      // true（提供了修改密码）
Console.WriteLine(sec.IsReadOnly);           // false
Console.WriteLine(sec.CanSave);              // true
```

| 属性 | 类型 | 说明 |
|---|---|---|
| `HasOpenPassword` | `bool` | 文件是否有打开密码（文件加密） |
| `HasModifyPassword` | `bool` | 文件是否有修改密码（写保护） |
| `HasModifyAccess` | `bool` | 是否已获得修改权限（提供了正确修改密码） |
| `IsReadOnly` | `bool` | 有修改密码但未获修改权限时为只读 |
| `CanSave` | `bool` | 是否允许保存（`!IsReadOnly`） |

输出：

```
True
True
True
False
True
```

#### 18.3 设置密码

`Workbook.Security` 提供设置打开 / 修改密码的方法，随后保存即生效：

```csharp
var wb = Excel.Create();
wb.Security.SetOpenPassword("secret");       // 打开密码（文件加密）
wb.Security.SetModifyPassword("write");      // 修改密码（写保护），默认建议只读
wb.Security.SetModifyPassword("write", readOnlyRecommended: false);  // 不提示只读
wb.SaveAs("secured.xlsx");
```

| 方法 | 参数 | 说明 |
|---|---|---|
| `SetOpenPassword` | `password` | 设置打开密码（文件加密），覆盖旧值；空 / 空白视为移除 |
| `SetModifyPassword` | `password` | 设置修改密码（写保护），覆盖旧值；空 / 空白视为移除 |
| `SetModifyPassword` | `readOnlyRecommended` | 是否建议以只读方式打开（默认 true） |
| `RemoveOpenPassword` | — | 移除打开密码（下次保存为无打开密码文件） |
| `RemoveModifyPassword` | — | 移除修改密码（要求已获修改权限） |
| `ClearAll` | — | 清空全部文件级密码（要求已获修改权限） |

输出：已写入 secured.xlsx

> ⚠️ 密码本体仅存储于 `WorkbookSecurity` 对象内部，不对外暴露明文；错误消息与日志不含密码。

#### 18.4 移除密码

打开含密码文件（提供正确密码获得授权）后，调用移除方法再保存即可去密码：

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions { OpenPassword = "secret", ModifyPassword = "write" });
wb.Security.RemoveOpenPassword();            // 下次保存为无打开密码
wb.Security.RemoveModifyPassword();          // 下次保存为无修改保护
wb.Security.ClearAll();                      // 清空全部文件级密码
wb.SaveAs("plain.xlsx");
```

输出：已写入 plain.xlsx

> ⚠️ `RemoveModifyPassword` / `ClearAll` 要求已获得修改权限（否则抛 `LiteExcelException`），防止未授权剥离 / 替换写保护。

#### 18.5 修改密码权限与只读

文件设置了修改密码但未提供时以只读方式打开，可通过安全状态属性确认：

```csharp
var wb = Excel.Open("readonly.xlsx");        // 该文件设置了修改密码但未提供
Console.WriteLine(wb.Security.IsReadOnly);   // true
Console.WriteLine(wb.Security.CanSave);      // false
```

输出：

```
True
False
```

- 文件设置了修改密码但未提供（或提供错误）时，工作簿以**只读**方式打开，`IsReadOnly = true`、`CanSave = false`，保存会抛 `LiteExcelException`。
- 提供正确的 `ModifyPassword` 即获得编辑授权（`HasModifyAccess = true`）。
- `SetModifyPassword` / `RemoveModifyPassword` / `ClearAll` 要求已获得修改权限，否则抛异常（防止未授权剥离 / 替换写保护）。
- 打开时捕获的原 `fileSharing` 在保存时透传保留；用户显式设置新修改密码时重新生成。

#### 18.6 保真回写

打开加密文件后 `SaveAs` 默认继承密码，无需重新设置：

```csharp
var wb = Excel.Open("secured.xlsx", new ExcelReadOptions
{
    OpenPassword = "secret",
    ModifyPassword = "write",
});
wb.SaveAs("secured_copy.xlsx");   // 默认继承打开密码
```

输出：已写入 secured_copy.xlsx

> ⚠️ `ModifyPasswordTouched`（用户显式改动过修改密码）时不透传原 fileSharing，按新设置的修改密码重新生成。含 VBA 宏的工作簿保存为 xlsx / xls 会报错（格式不支持宏）。
---

### 19. 工作表 / 工作簿保护

#### 19.1 工作表保护 `SheetProtection`

`Worksheet.Protection` 控制受保护工作表中允许 / 禁止的操作，可选密码（SHA-512 + salt 哈希）：

```csharp
var ws = Excel.Create().Worksheets["Sheet1"];
var p = new SheetProtection
{
    Enabled = true,
    // 允许的操作（默认 false 表示禁止）：
    SelectLockedCells = true,
    SelectUnlockedCells = true,
    // FormatCells / FormatColumns / FormatRows / InsertColumns / InsertRows /
    // InsertHyperlinks / DeleteColumns / DeleteRows / Sort / AutoFilter / PivotTables
    Objects = true,                // 默认 true 允许编辑对象
    Scenarios = true,              // 默认 true 允许编辑方案
};
p.SetPassword("protect123");      // 可选密码（方法，不能写在 initializer 里）
ws.Protection = p;
```

读回：

```csharp
var opened = Excel.Open("protected.xlsx");
var p = opened.Worksheets[0].Protection;
if (p is not null)
{
    Console.WriteLine(p.Enabled);
    Console.WriteLine(p.HasPassword);
    Console.WriteLine(p.VerifyPassword("protect123"));   // 对从文件读取的哈希有效
    p.RemovePassword();              // 移除保护密码（null/空白同样视为移除）
}
```

`SheetProtection` 参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `Enabled` | `bool` | 是否启用保护（写出 `sheetProtection` 的前提） |
| `SelectLockedCells` / `SelectUnlockedCells` | `bool` | 是否允许选定锁定 / 未锁定单元格（默认 true） |
| `Objects` / `Scenarios` | `bool` | 是否允许编辑对象 / 方案（默认 true） |
| `FormatCells`…`PivotTables` | `bool` | 允许的编辑操作（默认 false 禁止） |
| `SetPassword` / `RemovePassword` | 方法 | 设置 / 移除保护密码（null / 空白视为移除） |
| `VerifyPassword` | 方法 | 验证密码（仅对从文件读取的哈希有效） |

输出：

```
True
False
True
```

> ⚠️ 读回时 `HasPassword` 恒为 `False`——密码以 SHA-512 + salt 哈希落盘，库不把明文读回内存。是否设保护密码应通过 `VerifyPassword(...)` 判断，而不是 `HasPassword`。

#### 19.2 工作簿保护 `WorkbookProtection`

`Workbook.Protection` 锁定工作簿结构 / 窗口，可选密码：

```csharp
var wb = Excel.Create();
var p2 = new WorkbookProtection
{
    Enabled = true,
    LockStructure = true,   // 禁止插入/删除/移动/隐藏/重命名工作表
    LockWindows = false,
};
p2.SetPassword("wbpass");     // 可选密码（方法，不能写在 initializer 里）
wb.Protection = p2;
// p2.RemovePassword() 移除工作簿保护密码
wb.SaveAs("wbprotected.xlsx");
```

读回：

```csharp
var opened = Excel.Open("wbprotected.xlsx");
var p = opened.Protection;
if (p is not null)
    Console.WriteLine($"{p.Enabled} structure={p.LockStructure} hasPwd={p.HasPassword}");
```

`WorkbookProtection` 参数：

| 参数 | 类型 | 说明 |
|---|---|---|
| `Enabled` | `bool` | 是否启用保护（写出 `workbookProtection` 的前提） |
| `LockStructure` | `bool` | 禁止插入 / 删除 / 移动 / 隐藏 / 重命名工作表（默认 true） |
| `LockWindows` | `bool` | 锁定窗口（默认 false） |
| `SetPassword` / `RemovePassword` | 方法 | 设置 / 移除保护密码（null / 空白视为移除） |
| `VerifyPassword` | 方法 | 验证密码（仅对从文件读取的哈希有效） |

输出：

```
True structure=True hasPwd=False
```

> ⚠️ 同 19.1：读回的 `hasPwd=False` 不代表未设密码——原因见上（明文不入内存，用 `VerifyPassword` 判断）。

---

## 第三部分 多格式与平台

第 20–23 章：多格式行为与降级、流式与追加、AOT 兼容性，跨格式与平台差异集中于此。

### 20. 多格式行为（xlsx/xlsm 全能 + xls/xlsb/csv 限制与降级）

#### 20.1 格式能力矩阵

下表列出每个能力在各格式下的支持情况；其中 xls / xlsb / csv 不支持的能力在写出时经 `ExcelWriteOptions.OnDegradation` 上报（见第 22 章）。

| 能力 | xlsx | xlsm | xlsb | xls | csv |
|---|---|---|---|---|---|
| 单元格值 / 表头 | ☑️ | ☑️ | ☑️ | ☑️ | ☑️ |
| 样式（字体/颜色/边框/对齐/换行） | ☑️ | ☑️ | 仅 NumberFormat | 仅 NumberFormat | ❌ |
| 数字格式 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 合并单元格 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 自动筛选 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 行高 / 列宽 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 批注 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 数据验证 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 超链接 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 冻结窗格 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 图片（浮动/InCell） | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 条件格式 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 超级表 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 命名区域 | ☑️ | ☑️ | ❌ | ❌ | ❌ |
| 文档属性 | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| 打开密码 / 修改密码 | ☑️ | ☑️ | ☑️ | ❌ | ❌ |
| 公式 | ☑️ | ☑️ | ☑️ | ☑️ | ❌ |
| 图表 / 透视表 | 只保真 | 只保真 | 只保真 | ❌ | ❌ |

写出到 csv 时接通降级回调，观察被丢弃的能力：

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };

Excel.Write("matrix.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
        Console.WriteLine($"[降级] {info.Capability} -> {info.TargetFormat} @ {info.SheetName}: {info.Message}"),
});
```

输出：

```
[降级] Comments -> Csv @ Sheet1: CSV 不支持批注，工作表 'Sheet1' 的批注已丢弃。
```

#### 20.2 xls / xlsb 读回为静态值

xls / xlsb 读回时，样式降级为仅保留 `NumberFormat`（规避 BIFF 手写风险）；批注 / 数据验证 / 条件格式 / 图片 / 超级表等高级能力不读回。**这些降级在写出时会通过 `OnDegradation` 显式上报**（见第 22 章）。

读取 xls 文件（读回为静态值，不再携带除数字格式外的样式信息）：

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", "hello");
Excel.Write("roundtrip.xls", wb);

var reopened = Excel.Open("roundtrip.xls");
var ws = reopened.Worksheets["Sheet1"];
Console.WriteLine($"{ws.Name}: {ws.Cell("A1").GetString()}");
```

输出：

```
Sheet1: hello
```

#### 20.3 CSV 行为

- CSV 仅支持单工作表工作簿（写出多表抛 `NotSupportedException`）。
- 读取时首行不拆分为表头（`CsvBackend.Read`（见附录 B.4）返回原始行）。
- 分隔符：读取自动探测（逗号 > 分号 > Tab，仅统计引号外），`ExcelReadOptions.Separator` 可固定；写出默认逗号，`ExcelWriteOptions.Separator` 可指定。
- CSV 不支持样式 / 合并 / 批注 / 数据验证 / 超链接 / 图片 / 条件格式 / 超级表 / 命名区域 / 文档属性 / 公式 / 密码。

写入再读回（读回时首行作为数据行，不拆分为表头）：

```csharp
var sheet = new SheetData
{
    SheetName = "People",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
    },
};
Excel.Write("people.csv", sheet);

var wb = Excel.Open("people.csv");
var csv = wb.Worksheets[0].ToSheetData();
for (int r = 0; r < csv.Rows.Count; r++)
    Console.WriteLine($"{csv.Rows[r][0].GetString()}: {csv.Rows[r][1].GetString()}");
```

输出：

```
Name: Score
Alice: 95
Bob: 88
```

#### 20.4 加密文件格式限制

文件级密码（打开 / 修改）仅支持 xlsx / xlsm / xlsb；保存为 csv / xls 时若带密码会抛 `LiteExcelException`：

```csharp
var wb = Excel.Create();
wb.Security.SetOpenPassword("1");
try
{
    wb.SaveAs("enc.csv", ExcelFormat.Csv);
}
catch (LiteExcelException ex)
{
    Console.WriteLine(ex.Message);
}
```

输出：

```
无法写出 Csv：Csv 格式不支持文件级密码（打开密码/修改密码）。请使用 xlsx/xlsm/xlsb 保存，或先移除密码。
```

#### 20.5 保真回写

打开 xlsx / xlsm / xlsb 时，未映射的 OOXML 部件（宏 / 主题 / 绘图 / 图表 / 透视表等）被捕获并在保存时按二进制透传，避免静默删除。改表名不再丢 drawing 关联；追加数据不再丢宏 / 图表。

```csharp
var wb = Excel.Open("macro.xlsm");   // 打开包含宏的 xlsm
wb.Worksheets[0].Name = "重命名表";  // 改表名不丢 drawing 关联
wb.SaveAs("macro_copy.xlsm");
```

输出：

```
已写入 macro_copy.xlsm
```

---

### 21. 流式读取 / 进度回调 / 追加数据

#### 21.1 流式读取 `StreamRows`

逐行回调，不驻留内存，适合大文件。**仅支持 xlsx / xlsm**：

```csharp
var sheet = new SheetData
{
    SheetName = "Scores",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
    },
};
Excel.Write("scores.xlsx", sheet);

Excel.StreamRows("scores.xlsx", "Scores", row =>
{
    foreach (var cell in row)
        Console.Write($"{cell.GetString()} | ");
    Console.WriteLine();
});
```

输出：

```
Alice | 95 |
Bob | 88 |
```

#### 21.2 带进度读取 `ReadWithProgress`

先快速扫描总数据行数，再流式逐行读取，`current` 从 1 递增到 `total`（数据行数，不含表头）：

```csharp
var sheet = new SheetData
{
    SheetName = "Scores",
    Headers = new() { "Name", "Score" },
    Rows = new()
    {
        new Cell[] { Cell.FromText("Alice"), Cell.FromNumber(95) },
        new Cell[] { Cell.FromText("Bob"), Cell.FromNumber(88) },
        new Cell[] { Cell.FromText("Cara"), Cell.FromNumber(72) },
    },
};
Excel.Write("progress.xlsx", sheet);

Excel.ReadWithProgress("progress.xlsx", 0, (current, total) =>
    Console.WriteLine($"{current}/{total}"));
```

输出：

```
1/3
2/3
3/3
```

#### 21.3 追加数据 `Append`

`Excel.Append(path, SheetData, WorkbookProperties?)`（`SheetData` 见附录 B.1）：同名 sheet 合并列后追加行；不同名则作为新 sheet 加入；文件不存在时创建。**仅支持 xlsx / xlsm**：

```csharp
// 先写 3 行
var sheet1 = new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID" },
    Rows = new()
    {
        new Cell[] { Cell.FromNumber(1) },
        new Cell[] { Cell.FromNumber(2) },
        new Cell[] { Cell.FromNumber(3) },
    },
};
Excel.Write("append.xlsx", sheet1);

// 追加 2 行
var appendData = new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID" },
    Rows = new()
    {
        new Cell[] { Cell.FromNumber(4) },
        new Cell[] { Cell.FromNumber(5) },
    },
};
Excel.Append("append.xlsx", appendData);

// 读回验证
var read = Excel.ReadAsDataTable("append.xlsx");
Console.WriteLine(read.Rows.Count);
```

输出：

```
5
```

追加不改变既有 sheet 顺序，工作表级保留 rels 可继续复用（xlsm 追加不丢宏）。

**追加时表头对齐**：同名 sheet 追加时，新表头中不存在于原表头的列会**追加到原表头末尾**；数据行按列名对齐到合并后的列位置；缺失列补 `Empty`：

```csharp
// 原文件表头：ID | Name
Excel.Write("align.xlsx", new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID", "Name" },
    Rows = new() { new Cell[] { Cell.FromNumber(1), Cell.FromText("Alice") } },
});

// 追加的表头含新列 Price → 合并为 ID | Name | Price；数据行按列名对齐
Excel.Append("align.xlsx", new SheetData
{
    SheetName = "Data",
    Headers = new() { "ID", "Price" },
    Rows = new() { new Cell[] { Cell.FromNumber(2), Cell.FromNumber(9.5) } },
});

var sheet = XlsxReader.Read("align.xlsx", 0);   // 低层读回（见附录 B.2）
Console.WriteLine("headers: " + string.Join(" | ", sheet.Headers));
var ws2 = Excel.Open("align.xlsx").Worksheets["Data"];
Console.WriteLine($"C3 = {ws2.Cell("C3").Number}");   // Price 对齐到第 3 列
Console.WriteLine($"B3 类型 = {ws2.Cell("B3").Type}"); // 缺失列补 Empty
```

输出：

```
headers: ID | Name | Price
C3 = 9.5
B3 类型 = Empty
```

#### 21.4 流式写入 `CreateWriter`

`Excel.CreateWriter` 返回 `XlsxStreamWriter`（见附录 B.5），逐行写大文件，不驻留内存。**仅支持 .xlsx / .xlsm 扩展名**；使用后必须 `Dispose` / `Close` 完成文件：

```csharp
using var writer = Excel.CreateWriter("big_out.xlsx");
for (int i = 0; i < 1_000_000; i++)
    writer.WriteRow(new object?[] { i, $"row {i}", i * 1.5, i % 2 == 0 });
// using 结束自动 Close
```

输出：

```
已写入 big_out.xlsx
```

也可写入流（`LeaveOpen` 由调用方管理）：

```csharp
using var ms = new MemoryStream();
using (var writer = Excel.CreateWriter(ms))
    writer.WriteRow(new object?[] { 1, "a" });
ms.Position = 0;
var read = XlsxReader.Read(ms, 0);
```

`XlsxStreamWriter` 支持样式 / 公式 / 超链接随行写入（styles.xml 与 sheet rels 在 Close 时统一写出）；合并 / 筛选 / 图片等高级能力不支持。超链接数量极大时内存不再恒定（内部缓冲全部超链接引用）。

---

### 22. 降级回调 OnDegradation

`ExcelWriteOptions.OnDegradation` 为可选回调（默认 null，不注册则行为与历史版本完全一致，无破坏性）。写出到不支持某能力的格式（xls / xlsb / csv）时，对被静默丢弃的能力逐项回调：

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };   // csv 不支持批注

Excel.Write("out.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
    {
        Console.WriteLine($"[降级] {info.Capability} -> {info.TargetFormat} @ {info.SheetName}: {info.Message}");
    },
});
```

输出：

```
[降级] Comments -> Csv @ Sheet1: CSV 不支持批注，工作表 'Sheet1' 的批注已丢弃。
```

#### 22.1 能力枚举 `DegradationCapability`

`DegradationCapability` 枚举列出所有可被降级上报的能力：

```csharp
foreach (var cap in Enum.GetNames(typeof(DegradationCapability)))
    Console.WriteLine(cap);
```

输出：

```
Comments
DataValidation
AutoFilter
Images
DocumentProperties
NamedRanges
Styles
MergedCells
FreezePanes
Hyperlinks
RowHeights
ColumnWidths
Formulas
Charts
PivotTables
RichData
ConditionalFormatting
Tables
```

#### 22.2 降级信息 `DegradationInfo`

`DegradationInfo` 携带单次降级事件的完整描述：`Capability`（被丢弃能力）、`SheetName`（工作簿级能力为 null）、`TargetFormat`（目标格式）、`Message`（人类可读说明）。

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", "x");
ws.Comments = new Dictionary<string, string> { { "A1", "note" } };

Excel.Write("deg.csv", wb, new ExcelWriteOptions
{
    OnDegradation = info =>
    {
        Console.WriteLine($"Capability={info.Capability}");
        Console.WriteLine($"SheetName={info.SheetName}");
        Console.WriteLine($"TargetFormat={info.TargetFormat}");
        Console.WriteLine($"Message={info.Message}");
    },
});
```

输出：

```
Capability=Comments
SheetName=Sheet1
TargetFormat=Csv
Message=CSV 不支持批注，工作表 'Sheet1' 的批注已丢弃。
```

#### 22.3 样式降级细节

xls / xlsb 写出时，完整样式（字体 / 颜色 / 边框 / 对齐 / 换行）降级为仅保留 NumberFormat，经 `DegradationCapability.Styles` 上报：

```csharp
var wb = Excel.Create();
var ws = wb.Worksheets["Sheet1"];
ws.SetValue("A1", 1);
ws.Cell("A1").Style = new CellStyle { Bold = true, FillColor = "FFFF00" };

Excel.Write("style.xls", wb, new ExcelWriteOptions
{
    OnDegradation = info => Console.WriteLine($"{info.Capability}: {info.Message}"),
});
```

输出：

```
Styles: xls 仅支持数字格式，工作表 'Sheet1' 的完整样式（字体/颜色/边框/对齐/换行）已降级。
```

---

### 23. AOT 兼容性（DAM、IsAotCompatible、验证方式与成果摘要）

#### 23.1 DAM 标注

List\<T\> 反射映射 API 已用 `[DynamicallyAccessedMembers]` 标注，AOT / 裁剪安全：

- `Excel.Create<T>` / `Excel.Write<T>` / `Excel.Read<T>` 标注 `PublicProperties`（读取还含 `PublicParameterlessConstructor`）。
- `Worksheet.ImportData<T>` / `WorksheetCollection.Add<T>` 标注 `PublicProperties`。
- `XlsxReader.Read<T>` / `XlsxWriter.Write<T>` 同样标注（见附录 B.2 / B.3）。

以第 5 章的 `Person` 为例（映射无需额外配置，公开属性由库内标注保留）：

```csharp
var people = new List<Person> { new() { Name = "张三", Age = 30 } };
Excel.Write("people.xlsx", people);

var read = Excel.Read<Person>("people.xlsx");
Console.WriteLine(read.Count);
```

输出：

```
1
```

#### 23.2 IsAotCompatible

net8.0 目标在 csproj 声明 `IsAotCompatible=true`，全部公开 API 兼容 Native AOT / 裁剪：

```csharp
var wb = Excel.Create("Sheet1");
wb.Worksheets["Sheet1"].SetValue("A1", "x");
Excel.Write("aot.xlsx", wb);

var reopened = Excel.Open("aot.xlsx");
Console.WriteLine(reopened.Worksheets.Count);
```

输出：

```
1
```

#### 23.3 验证方式与成果摘要

- 经原生 AOT 可执行文件实测，全部公开 API 通过。
- AOT 零 IL 警告 + 运行期断言通过。
- 注意：`Excel.Read<T>` / `XlsxReader.Read`（见附录 B.2）仅支持 xlsx / xlsm；**xls / xlsb / csv 必须走 `Excel.Open(path)`**（按扩展名路由后端）。按需读表名时用 `Excel.Open` 处理非 zip 格式。

非 zip 格式读取入口：

```csharp
var wb = Excel.Open("data.xls");
Console.WriteLine(string.Join(",", wb.Worksheets.Names));
```

输出：

```
Sheet1
```

#### 23.4 InvariantGlobalization

全局不变量（常见于 AOT / 容器）：

- 发布时加 `<InvariantGlobalization>true</InvariantGlobalization>` 不会影响本库任何功能；读取侧 `Encoding.GetEncoding` 与写入侧 `CultureInfo.InvariantCulture` 均通过验证。
- **前提**：基准日期边界（1900/1904 日期系统）与 xls ANSI 字符串需走 `Latin1`，非当前系统代码页的字符可能失真——这是 BIFF8 的固有限制，与 AOT 无关。

不变量下日期写出 / 读回（固定 `yyyy-MM-dd` 数字格式）：

```csharp
var wb = Excel.Create();
wb.Worksheets["Sheet1"].SetValue("A1", new DateTime(2024, 1, 2));
Excel.Write("inv.csv", wb);

var reopened = Excel.Open("inv.csv");
var data = reopened.Worksheets[0].ToSheetData();
Console.WriteLine(data.Rows[0][0].GetString());
```

输出：

```
2024-01-02
```

---

## 第四部分 注意事项

第 24–25 章：异常处理与大文件注意事项，上线前建议通读。

### 24. 异常处理

#### 24.1 异常分层

- `LiteExcelException`：库所有异常的基类。
- `LiteXlsxException`：旧异常名称的兼容别名（`[Obsolete]`，请改用 `LiteExcelException`）。
- `InvalidSheetNameException`：当 Sheet 名不合法（为空、超过 31 字符、包含非法字符）时抛出，含 `SheetName` 属性。

非法 Sheet 名在写出时抛出（`SheetName` 携带原始名称）：

```csharp
var sheet = new SheetData { SheetName = "非法?名称" };
try
{
    Excel.Write("bad.xlsx", sheet);
}
catch (InvalidSheetNameException ex)
{
    Console.WriteLine($"非法 Sheet 名：{ex.SheetName}");
}
```

输出：

```
非法 Sheet 名：非法?名称
```

#### 24.2 常见异常场景

| 场景 | 异常 |
|---|---|
| 文件不存在 | `FileNotFoundException` |
| 路径为空 | `ArgumentException` |
| 工作表名重复 / 找不到 | `LiteExcelException` |
| Sheet 名非法 | `InvalidSheetNameException` |
| 保存路径扩展名与格式不匹配 | `LiteExcelException` |
| 新建簿未指定路径就 Save | `LiteExcelException` |
| 只读工作簿（有修改密码未授权）保存 | `LiteExcelException` |
| 带密码保存为 csv / xls | `LiteExcelException` |
| 含宏保存为 xlsx / xls | `LiteExcelException` |
| 流式读取 / 追加非 xlsx/xlsm | `LiteExcelException` |
| CSV 多表写出 | `NotSupportedException` |
| 单元格类型不匹配强类型读取 | `InvalidCastException` |

典型捕获顺序（先具体异常后基类）：

```csharp
var wb = Excel.Create();
try
{
    wb.Save();   // 新建簿未指定路径就保存
}
catch (LiteExcelException ex)
{
    Console.WriteLine(ex.Message);
}
```

输出：

```
当前工作簿没有目标路径，请使用 SaveAs 指定保存位置
```

#### 24.3 建议

```csharp
try
{
    var wb = Excel.Open("report.xlsx", new ExcelReadOptions { OpenPassword = "secret" });
    wb.SaveAs("out.xlsx");
}
catch (LiteExcelException ex)
{
    Console.WriteLine($"LiteExcel: {ex.Message}");
}
catch (Exception ex)
{
    Console.WriteLine($"Other: {ex.Message}");
}
```

输出：

```
已写入 out.xlsx
```

---

### 25. 大文件注意事项

- **流式读取**：`Excel.StreamRows`（见附录 B.2）逐行回调，不驻留内存，适合大文件（仅 xlsx / xlsm）。
- **流式写入**：`Excel.CreateWriter` / `XlsxStreamWriter`（见附录 B.5）逐行写大文件，不驻留内存（仅 xlsx / xlsm）。使用后必须 `Dispose` / `Close`。
- **进度**：`Excel.ReadWithProgress` 先扫描总行数再流式读取，适合长任务进度展示。
- **内存模型**：`Excel.Open` / `Excel.Create` 返回的 `Workbook` 是内存模型，整簿加载到内存。超大文件请用流式 API 而非 `Excel.Open`。
- **超链接数量**：流式写入器在超链接数量极大时内存不再恒定（内部缓冲全部超链接引用）。
- **追加**：`Excel.Append` 会读取整个既有文件再写出，适合中小文件增量追加。

---

## 附录 A 对象模型速查（类 / 成员索引表）

#### `Excel` 静态类（对象模型入口）

| 成员 | 说明 |
|---|---|
| `Open(path[, options])` | 按扩展名自动识别格式打开 |
| `Open(path, format[, options])` | 指定格式打开 |
| `Open(stream, format[, options])` | 从流打开（必须指定格式） |
| `Create()` / `Create(sheetName)` / `Create(string[])` / `Create(format)` | 新建空簿 |
| `Create<T>(data[, sheetName, format, configure])` | 新建并写 List\<T\> |
| `Create(DataTable[, sheetName, format])` | 新建并写 DataTable |
| `Write(path, Workbook[, options])` | 写出工作簿 |
| `Write(path, SheetData[, options])` | 写出单表（低层） |
| `Write(path, DataTable[, sheetName, options])` | 写出 DataTable |
| `Write<T>(path, data[, sheetName, configure])` | 写出 List\<T\> |
| `Read<T>(path[, sheetName, configure])` | 读为 List\<T\> |
| `ReadAsDataTable(path[, sheetName, firstRowIsHeader])` | 读为 DataTable |
| `GetSheetNames(path)` / `GetSheetNames(stream)` | 列出工作表名 |
| `StreamRows(path, sheetName, onRow)` | 流式逐行读取 |
| `CreateWriter(path)` / `CreateWriter(stream)` | 创建流式写入器 |
| `Append(path, SheetData[, properties])` | 追加数据 |
| `ReadWithProgress(path, sheetIndex, onProgress)` | 带进度读取 |
| `DetectFormat(path)` | 按扩展名识别格式 |

#### `Workbook`

| 成员 | 说明 |
|---|---|
| `Worksheets` | 工作表集合（`WorksheetCollection`） |
| `Properties` | 文档属性（`WorkbookProperties`） |
| `Format` | 当前格式（`ExcelFormat`） |
| `Security` | 文件级安全（`WorkbookSecurity`） |
| `Protection` | 工作簿保护（`WorkbookProtection`） |
| `Names` | 命名区域（`List<NamedRange>`） |
| `CurrentPath` | 当前目标路径 |
| `Save()` / `SaveAs(path[, format])` / `Save(stream, format)` | 保存 |

#### `Worksheet`

| 成员 | 说明 |
|---|---|
| `Name` | 工作表名 |
| `Cell(row, col)` / `Cell(address)` | 访问单元格 |
| `Range(address)` / `Range(r1, c1, r2, c2)` | 访问区域 |
| `Cells` | 整表单元格集合 |
| `SetValue(row, col, value)` / `SetValue(address, value)` | 设置值 |
| `Merge` / `Unmerge` / `MergedRanges` | 合并 |
| `RowHeights` / `ColumnWidths` | 行高 / 列宽 |
| `AutoColumnWidths()` | 列宽自适应 |
| `HeaderStyle` / `DefaultStyle` / `RowStyles` / `ColumnStyles` | 样式 |
| `Comments` | 批注 |
| `Validations` | 数据验证 |
| `Filter` | 自动筛选 |
| `ConditionalFormats` | 条件格式 |
| `Images` / `AddImage(...)` | 图片 |
| `Protection` | 工作表保护 |
| `Tables` / `AddTable` / `RemoveTable` | 超级表 |
| `FreezeRows` / `FreezeColumns` / `FreezeHeader` | 冻结窗格 |
| `ImportData<T>(data[, configure])` / `ImportData(DataTable[, includeHeader])` | 清空重建导入 |
| `ToSheetData()` | 导出为低层 SheetData 模型 |
| `RowCount` / `MaxColumn` | 尺寸信息 |

#### `Cells`

| 成员 | 说明 |
|---|---|
| `this[int row, int column]` | 按行列索引器 |
| `this[string address]` | 按 A1 地址索引器 |
| `Range(address)` / `Range(r1, c1, r2, c2)` | 提取区域 |
| `SetValue(...)` | 便捷写值 |
| `Clear()` | 清空整表值 |
| `GetEnumerator()` | 枚举已有单元格 |

#### `ExcelRange`

| 成员 | 说明 |
|---|---|
| `FirstRow` / `FirstCol` / `LastRow` / `LastCol` | 区域边界（1-based） |
| `Address` | A1 地址 |
| `RowCount` / `ColumnCount` | 尺寸 |
| `Cell(rowOffset, colOffset)` | 区域内相对偏移 |
| `Fill(value)` / `Fill(object?[,])` | 批量写入 |
| `ToValues()` / `ToCells()` | 读回 |
| `Style` | 整区统一样式 |
| `Merge()` / `Unmerge()` | 合并 |
| `Clear()` | 清空 |
| `GetEnumerator()` | 枚举（行优先） |

#### `Cell`

| 成员 | 说明 |
|---|---|
| `Type` / `Text` / `Number` / `Date` / `Boolean` | 值字段 |
| `Style` / `NumberFormat` | 样式 |
| `Formula` / `IsFormula` | 公式 |
| `Hyperlink` | 超链接 |
| `IsEmpty` | 是否空 |
| `FromText` / `FromNumber` / `FromDate` / `FromBoolean` / `FromFormula` / `Empty` | 工厂方法 |
| `SetValue(object?)` | 设置值 |
| `GetString` / `GetDouble` / `GetDateTime` / `GetBoolean` / `GetValue` | 强类型读取 |
| `TryGetString` / `TryGetDouble` / `TryGetDateTime` / `TryGetBoolean` | Try 读取 |

#### 模型类

- `CellStyle` / `BorderStyle` / `BorderEdge` / `HorizontalAlignment` / `VerticalAlignment`
- `CellRange`（0-based）
- `AutoFilter` / `FilterColumn` / `FilterType` / `FilterOperator`
- `DataValidation` / `DataValidationType`
- `ConditionalFormat` / `ConditionalFormatType` / `ConditionalOperator` / `ColorScaleInfo` / `DataBarInfo` / `IconSetInfo` / `IconSetStyle`
- `Hyperlink`
- `NamedRange`
- `XlTable` / `XlTableColumn` / `TableStyleStyle`
- `WorksheetImage` / `ImageAnchor` / `ImagePlacement` / `ImageMoveMode`
- `SheetProtection` / `WorkbookProtection`
- `WorkbookProperties` / `WorkbookSecurity`
- `DegradationInfo` / `DegradationCapability`
- `ExcelFormat` / `ExcelReadOptions` / `ExcelWriteOptions` / `WriteOptions<T>` / `ReadOptions<T>` / `LiteColumnAttribute`
- `CellRef`（A1 引用工具，静态类）
- `XlsxStreamWriter`
- `LiteExcelException` / `LiteXlsxException` / `InvalidSheetNameException`

---

## 附录 B 低层 API 参考（SheetData / XlsxReader / XlsxWriter / CsvBackend / 流式）

> **适用场景**：适合自定义 / 裸行数据 / 大文件场景。日常用法优先用对象模型 API（第 2-25 章）。低层 API 的坐标约定：`SheetData.Rows` 的 `Cell` 是 0-based 网格，`Headers` 为首行表头文本。

#### B.1 `SheetData`（一张工作表的完整数据）

↳ 正文：第 5 章 数据类型与转换（List<T> / DataTable 映射底层即 SheetData）、第 21 章 流式读取 / 进度回调 / 追加数据（流式 / 追加的数据载体）

```csharp
public sealed class SheetData
{
    public string SheetName { get; set; } = "Sheet1";
    public List<string> Headers { get; set; } = new();
    public List<IReadOnlyList<Cell>> Rows { get; set; } = new();
    public List<CellRange> MergedRanges { get; set; } = new();
    public AutoFilter? Filter { get; set; }
    public bool FreezeHeader { get; set; }
    public int FreezeRows { get; set; }
    public int FreezeColumns { get; set; }
    public List<double>? ColumnWidths { get; set; }
    public CellStyle? HeaderStyle { get; set; }
    public CellStyle? DefaultStyle { get; set; }
    public Dictionary<int, CellStyle>? RowStyles { get; set; }
    public Dictionary<int, CellStyle>? ColumnStyles { get; set; }
    public Dictionary<int, double>? RowHeights { get; set; }
    public Dictionary<string, string>? Comments { get; set; }
    public List<DataValidation>? Validations { get; set; }
    public List<ConditionalFormat> ConditionalFormats { get; set; } = new();
    public string? CodeName { get; set; }
    public List<WorksheetImage>? Images { get; set; }
    public SheetProtection? Protection { get; set; }
    public List<XlTable> Tables { get; set; } = new();
}
```

#### B.2 `XlsxReader`（静态类，零反射，AOT 安全）

↳ 正文：第 3 章 文件导航：打开 / 创建 / 保存 / 格式（打开 / 流读取的低层入口）、第 21 章 流式读取 / 进度回调 / 追加数据（21.4 流式读回）

| 成员 | 说明 |
|---|---|
| `Read(path, sheetIndex[, firstRowIsHeader])` | 按索引读单表 |
| `Read(path, sheetName[, firstRowIsHeader])` | 按名称读单表 |
| `Read(stream, sheetIndex/name[, firstRowIsHeader])` | 从流读单表 |
| `ReadAll(path)` / `ReadAll(stream)` | 读所有表 |
| `Read<T>(path, sheetIndex/name[, configure])` | 读为 List\<T\> |
| `ReadAsDataTable(path, sheetIndex/name[, firstRowIsHeader])` | 读为 DataTable |
| `GetSheetNames(path)` / `GetSheetNames(stream)` | 列出工作表名 |
| `StreamRows(path/stream, sheetName, onRow)` | 流式逐行读取 |
| `ReadWithProgress(path, sheetIndex, onProgress)` | 带进度读取 |
| `ReadProperties(path)` / `ReadProperties(stream)` | 读取文档属性 |

#### B.3 `XlsxWriter`（静态类，零反射，AOT 安全）

↳ 正文：第 3 章 文件导航：打开 / 创建 / 保存 / 格式（Excel.Write 写出）、第 21 章 流式读取 / 进度回调 / 追加数据（21.3 Append 的低层写出）

| 成员 | 说明 |
|---|---|
| `Write(path, SheetData[, properties])` | 写单表 |
| `Write(path, IReadOnlyList<SheetData>[, properties])` | 写多表 |
| `Write(stream, SheetData[, properties])` | 写单表到流 |
| `Write(stream, IReadOnlyList<SheetData>[, properties])` | 写多表到流 |
| `Write<T>(path, data[, configure])` | 写 List\<T\> |
| `Write(path, DataTable[, sheetName])` | 写 DataTable |
| `Append(path, SheetData[, properties])` | 追加数据 |
| `AutoColumnWidths(SheetData)` | 列宽自适应 |

注意：`XlsxWriter.Write` 对 `.xlsm` 扩展名自动写出 macroEnabled 主文档类型；`SheetData` 写出时校验表名（`InvalidSheetNameException`）与重复表名（`LiteExcelException`）。

#### B.4 `CsvBackend`（内部类，CSV 格式后端）

↳ 正文：第 20 章 多格式行为（20.3 CSV 行为）

> 低层 CSV 后端为 `internal`，日常 CSV 读写请走 `Excel.Open` / `Excel.Write`。此处列出行为要点供参考。

- 实现 RFC 4180 基础子集：双引号包裹含分隔符 / 换行 / 引号的字段。
- 读取分隔符自动探测（逗号 > 分号 > Tab，仅统计引号外）；`ExcelReadOptions.Separator` 可固定。
- 写出默认逗号，`ExcelWriteOptions.Separator` 可指定。
- 仅表格数据，不支持样式 / 合并 / 批注等 Excel 专有能力。

#### B.5 `XlsxStreamWriter`（流式写入器）

↳ 正文：第 21 章 流式读取 / 进度回调 / 追加数据（21.4 流式写入 CreateWriter）

> 适合大文件逐行写入。通过 `Excel.CreateWriter(path|stream)` 获取。

| 成员 | 说明 |
|---|---|
| `Create(path)` / `Create(stream)` | 创建写入器 |
| `WriteRow(IEnumerable<object?>)` | 写一行值 |
| `WriteRow(IEnumerable<Cell>)` | 写一行 Cell |
| `Close()` / `Dispose()` | 完成文件（必须调用） |

- 采用内联字符串（inlineStr），避免共享字符串表预扫描。
- 支持单工作表；样式 / 公式 / 超链接随行写入（styles.xml 与 sheet rels 在 Close 时统一写出）。
- 合并 / 筛选 / 图片等高级能力不支持。
- 超链接数量极大时内存不再恒定。

#### B.6 `CellRef`（A1 引用工具，静态类）

↳ 正文：第 4 章 单元格与取值（A1 地址访问单元格）

| 成员 | 说明 |
|---|---|
| `Parse(cellRef)` | `"A1"` -> `(row=0, col=0)` |
| `TryParse(cellRef, out pos)` | 尝试解析 |
| `ParseRange(range)` | 解析区域引用（0-based 含端点） |
| `ToString(row, col)` | `(0,0)` -> `"A1"` |
| `ColToLetter(col)` | `0` -> `"A"` |
| `LetterToCol(letters)` | `"A"` -> `0` |

#### B.7 新旧 API 对照

对象模型 API 与低层 API 等价关系速查（对象模型按扩展名自动路由格式；低层 API 仅处理 xlsx/xlsm）：

| 场景 | 对象模型 API | 低层 API |
|---|---|---|
| 打开文件 | `Excel.Open(path[, options])` | `XlsxReader.Read(path, 0)`（单表）/ `XlsxReader.ReadAll(path)`（全部表） |
| 新建并写出 | `Excel.Create(...)` + `wb.SaveAs(path)` | `XlsxWriter.Write(path, sheet)` |
| 读单表 | `wb.Worksheets[i]` / `wb.Worksheets["name"]` | `XlsxReader.Read(path, sheetIndex)` / `XlsxReader.Read(path, sheetName)` |
| 读 List\<T\> | `Excel.Read<T>(path[, sheetName, configure])` | `XlsxReader.Read<T>(path, 0[, configure])` |
| 读 DataTable | `Excel.ReadAsDataTable(path[, sheetName])` | `XlsxReader.ReadAsDataTable(path, 0)` |
| 写 List\<T\> | `Excel.Write(path, list[, sheetName, configure])` | `XlsxWriter.Write(path, list)` |
| 写 DataTable | `Excel.Write(path, table[, sheetName])` | `XlsxWriter.Write(path, table, sheetName)` |
| 流式读 | `Excel.StreamRows(path, sheetName, onRow)` | `XlsxReader.StreamRows(path, sheetName, onRow)` |
| 流式写 | `Excel.CreateWriter(path)` | `XlsxStreamWriter.Create(path)` |
| 追加数据 | `Excel.Append(path, sheetData[, properties])` | `XlsxWriter.Append(path, sheetData[, properties])` |
| 列出表名 | `Excel.GetSheetNames(path)` | `XlsxReader.GetSheetNames(path)` |

> ⚠️ xls / xlsb / csv 无低层 API，请一律使用对象模型 `Excel.Open` / `Excel.Write`（按扩展名路由）。
---

*本手册覆盖 LiteExcel 当前主线版本的全部公开能力。*