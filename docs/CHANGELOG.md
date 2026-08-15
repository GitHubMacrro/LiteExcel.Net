# Changelog

## [2.2.1] - 2026-08-15

### Added
- **保存部件保留（save fidelity）**：通过 `Excel.Open` 打开后修改再保存时，未映射的 OOXML 部件按原始字节保留，不再被静默删除。
  - `xlsm` 宏部件 `xl/vbaProject.bin` 及其工作簿关系、内容类型声明在保存时透传，宏不丢失。
  - 主题、绘图、图表、表格、外部链接等未映射部件与关系一并保留。
  - 结构变化保护：打开后新增/删除/重命名/移动工作表时，工作表级未映射关系不再复用到新文件，避免错位；部件字节仍保留为无害条目。
- **保留部件测试**：新增 `PreservationTests`（自定义部件 / 假宏 `vbaProject.bin` / 外部超链接 rels / 结构变化降级 / 新建无保留），共 4 个。

### Changed
- `Excel.OpenCore` 改为单次解压完成读表 / 读属性 / 捕获保留部件，保证三者在同一文件快照上完成。
- `XlsxWriter` 内部新增 rels 合并与 `[Content_Types].xml` 合并逻辑，重建部件与保留部件共存。

### Notes / 兼容性
- 既有 API 无任何破坏性变更；`xlsb` / `xls` 仍未实现，图片、图表、透视表等高级能力不在本版本范围。

## [2.2.0] - 2026-08-15

### Added
- **对象模型 API（Excel 门面）**：新增统一入口 `Excel`，提供 `Excel.Open(path)`、`Excel.Create(format)`、`Excel.Read<T>(path)`、`Excel.Write<T>(path, data)`、`Excel.ReadAsDataTable(path)`、`Excel.Write(path, DataTable)`、`Excel.GetSheetNames(path)`、`Excel.StreamRows(path, name, onRow)`、`Excel.CreateWriter(path/stream)` 等。
- **对象模型层级**：`Workbook -> Worksheet -> Cells/Cell/ExcelRange`，坐标统一为 1-based，支持 A1 地址。
  - `Workbook`：`Worksheets` 集合（新增/删除/移动/按名访问）、`Properties` 文档属性、`Save()` / `SaveAs(path[, format])` / `Save(stream, format)`。
  - `Worksheet`：`Cell("A1")` / `Cell(row, col)` / `Range("A1:D10")` / `Cells`、`SetValue`、`Merge` / `Unmerge`、冻结表头、样式、批注、验证、筛选。
  - `Cells`：索引器（1-based 坐标 / A1 地址）、`Range(...)`、枚举、批量清空。
  - `ExcelRange`：`Fill` / `Clear` / `Style` / `Merge` / `Unmerge` / `ToValues` / `ToCells` / 枚举。类名为 `ExcelRange`（非 `Range`），避免与 BCL `System.Range` 冲突。
- **`Cell` 便利方法**：`GetString` / `GetDouble` / `GetDateTime` / `GetBoolean` / `TryGet*` / `GetValue` / `SetValue`、`Style` / `NumberFormat`。
- **公式字符串支持**：读取解析 `<f>` 公式文本，写入输出 `<f>` + 缓存 `<v>`，不做公式计算引擎。`Cell.FromFormula` / `Cell.IsFormula`。
- **CSV 格式后端**：`ExcelFormat.Csv` 读写，RFC4180 子集（含分隔符/换行的字段用引号包裹，UTF-8 BOM）。CSV 仅覆盖表格数据，不支持样式/合并等 Excel 专有能力。
- **流式写入**：`XlsxStreamWriter`（`Excel.CreateWriter` 创建），逐行写入大文件不驻留内存，使用内联字符串；与流式读取 `StreamRows` 对应。
- **冻结表头读取**：`XlsxReader` 新增解析 `<pane state="frozen">`，`FreezeHeader` 可正确读回。
- **格式枚举占位**：`ExcelFormat.Xlsb` / `Xls` 已定义，读写抛 `NotSupportedException`。

### Changed
- **`Range` 更名 `ExcelRange`**：为避免与 BCL `System.Range` 的命名冲突，区域类型命名为 `ExcelRange`。
- **对象模型 API 读取首行语义**：`Worksheet` 采用原始网格模型，首行不强制拆分表头，表头识别归属映射层（`Excel.Read<T>` / `ReadAsDataTable` 仍按原语义处理）。

### Fixed
- **`Cell.SetValue` 写回遗漏**：`SetValue` 的 `CopyFrom` 分支未触发单元格变更通知，导致新值/公式/样式不落入网格，已修复。
- **`SheetToDataTable` 输入副作用**：无表头时不再修改传入 `SheetData` 的 `Headers`（在副本上补齐列名）。

### Notes / 兼容性
- `XlsxReader` / `XlsxWriter` / `SheetData` / `Cell` / `DataTable` / `List<T>` 等既有 API 全部保留，未做任何破坏性变更；新旧 API 混用，写出的文件互相兼容。
- 对象模型 API 中仅 `List<T>` 反射入口（`Excel.Read<T>` / `Excel.Write<T>`）不兼容 AOT，其余均为 AOT 安全。
- 已知限制：`xlsb` / `xls` 未实现；图片、图表、透视表、条件格式不在本版本范围。

## [2.1.4] - 2026-08-14

### Changed
- **发布元数据**：补充 `RepositoryUrl`、`PackageProjectUrl`，更新包描述与版权信息；无 API 变更。

### Fixed
- **Append 文档属性保留**：Append 现保留已有工作簿的作者、标题、主题和创建时间，并自动更新最后修改时间。
- **文档属性默认时间**：写入 `WorkbookProperties` 时，未显式指定的创建和修改时间自动填充当前时间。
- **app.xml 工作表元数据**：工作表数量和名称现按实际工作簿正确写入。
- **Excel 兼容性测试**：新增可公开提交、由 Microsoft Excel 创建的匿名 fixture，测试不再因私有真实文件缺失而静默通过。

### Changed
- **异常命名统一**：主异常类型更名为 `LiteExcelException`（旧名 `LiteXlsxException` 保留为兼容别名）。

## [2.1.1] - 2026-08-14

### Added
- **文档属性读写**：WorkbookProperties 模型（Creator 作者 / LastModifiedBy 最后保存者 / Created 创建时间 / Modified 修改时间 / Title 标题 / Subject 主题 / Application 应用名），XlsxWriter.Write(..., WorkbookProperties) 写出，XlsxReader.ReadProperties() 读取。Application 默认取宿主程序集名，可显式覆盖。

### Fixed
- **fills gray125 保留填充**：styles.xml 的 fills 列表前两个固定为 none + gray125（Excel OOXML 规范要求），用户填充色从索引 2 开始，修复了写入带填充色的表格在 Excel 中显示 12.5% 灰色图案的问题。

### Changed
- **命名统一**：包名、命名空间、测试/示例项目统一为 LiteExcel（原 CustomWin.Utils.LiteXlsx）。
- **批注作者名**：comments.xml 的 author 统一为 LiteExcel。

## [2.1.0] - 2026-08-01

### Added
- **Stream 读写支持**：XlsxWriter.Write(Stream, ...)、XlsxReader.Read(Stream, ...)、XlsxReader.StreamRows(Stream, ...)、XlsxReader.ReadAll(Stream)、XlsxReader.GetSheetNames(Stream)、DataTableApi.ReadAsDataTable(Stream, ...) 等所有读写 API 新增 Stream 重载。
- **进度回调**：XlsxReader.ReadWithProgress(string path, int sheetIndex, Action<int, int> onProgress) 支持带进度回调的逐行读取，onProgress(current, total) 从 1 递增到总数据行数。
- **行高读写**：SheetData.RowHeights（Dictionary<int, double>，key = 0-based 行索引，value = 磅值），写入时自动应用到对应行。
- **列宽自适应**：XlsxWriter.AutoColumnWidths(SheetData) 根据表头和数据内容估算每列最佳宽度（中文字符算 2，英文/数字算 1，范围 [8, 50]），自动设置 SheetData.ColumnWidths。
- **单元格批注读写**：SheetData.Comments（Dictionary<string, string>，key = A1 格式引用如 "A1"，value = 批注文本），写入时自动生成 comments.xml 和 sheet rels，读取时自动解析批注。
- **追加数据**：XlsxWriter.Append(string path, SheetData newData) 向已有文件追加数据，同名 sheet 合并列后追加行，不同名则作为新 sheet 加入；文件不存在时直接创建。
- **数据验证读写**：SheetData.Validations（List<DataValidation>），DataValidationType 枚举（List、WholeNumber、Decimal、Date），支持下拉列表（逗号分隔公式）、数值区间、空白允许、输入提示。
- **Sheet 名校验**：InvalidSheetNameException 在写出时校验 Sheet 名（非空、不超 31 字符、不含非法字符 \/?*[]:），非法时抛出。
- **错误提示优化**：LiteXlsxException 统一异常基类，所有读取/写入异常均使用此类型，错误信息包含具体原因和上下文。

### Changed
- **样式优先级**：行列级样式优先级规则明确为**覆盖式**（非合并式），即单元格 > 行 > 列 > 全表默认，高优先级完整覆盖低优先级属性。
- **真实文件兼容性**：改进读取逻辑，支持 Excel 直接创建的 xlsx 文件（含非标准命名空间、rId 引用、空 sheet 等边缘情况）。
