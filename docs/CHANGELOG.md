# Changelog

## [2.1.3] - 2026-08-14

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
