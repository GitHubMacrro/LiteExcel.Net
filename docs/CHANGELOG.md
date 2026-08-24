# Changelog

## [2.4.5] - 2026-08-24

### Added

- **命名区域读回**：`Workbook.Names` 读取 `workbook.xml` 中的 `definedNames`（名称/范围/作用域），兼容全局与 sheet-local 定义。打开即回填；写出未修改时原样保留。
- **`List<T>` 公式列**：`LiteColumnAttribute.IsFormula` / `WriteOptions.Column(..., isFormula:)`，字符串属性按公式写出（值可带或不带前导 `=`）。
- **高层 API 带数据建簿**：`Excel.Create<T>(data, sheetName, format, configure?)` / `Excel.Create(DataTable, sheetName?, format)` 一步建簿并写入首表；`Worksheet.ImportData<T>(...)` / `ImportData(DataTable, includeHeader)` 清空整表并从 A1 重建；`WorksheetCollection.Add<T>(name, data, configure?)` / `Add(name, DataTable)` 批量加表并写数据。泛型入口沿用 `[DynamicallyAccessedMembers]`，AOT 安全。`DataTableToSheet` 收敛为单一实现（统一 `CellFactory.FromObject`）。

### Changed

- **`List<T>` 映射转为 Native AOT 兼容**：`Excel.Read<T>` / `Excel.Write<T>` / `XlsxReader.Read<T>` / `XlsxWriter.Write<T>` 等反射入口由 `[RequiresUnreferencedCode]` 改为 `[DynamicallyAccessedMembers]` 标注，库以 `IsAotCompatible` 编译，新增 `tests/LiteExcel.AotSmoke` 原生 AOT 冒烟项目（`PublishAot` + `TrimmerRootAssembly`）。
  - 验证：`dotnet publish -r win-x64` 零 IL/CS 警告；原生可执行文件运行 15 项断言全部通过（含 `[LiteColumn]` 特性、公式列、Fluent 表达式配置、可空/decimal、DataTable 往返）。
  - 调用方以具体类型调用零警告；在未标注的开放泛型中转发会收到 IL2091 提示，按提示补标注即可。

### Notes

- 全量 **518 测试**，通过 net48 + net8.0；net48 以 internal polyfill 提供 `DynamicallyAccessedMembersAttribute`。

## [2.4.4] - 2026-08-22

### Added

- **真实 Excel 样本对拍基础设施**：打包集成 `tests/LiteExcel.Tests/Fixtures/` 两个新生成文件（含条件格式/图表/浮动图片），让库自读之外还有真实 Excel 对照。
- **InCell richData 图片读回**：`Worksheet.Images` 读取时回填 InCell 图片（Placement=InCell、Row/Column、Extension、Data）。
- **条件格式长尾类型**：支持 `ContainsText`/`BeginsWith`/`EndsWith`/`NotContainsText`/`Blanks`/`NoBlanks`/`Errors`/`NoErrors`/`Unique`/`Duplicate`/`TimePeriod`/`Top10`/`AboveAverage`/`BelowAverage`。
  - 参数：`Text.`、`TimePeriod`、`Rank`、`Percent`。
  - xls/xlsb/csv 一律走降级上报。
  - xlsx/xlsm 启读回双向支持。

### Fixed

- **条件格式长尾类型 cfRule type 值错误**：`unique`/`duplicate`/`blanks`/`noBlanks`/`errors`/`noErrors` 写出非标准 `ST_CfType` 值，Excel 打开报「已修复的部件 / XML 错误」；`top10` `rank="3%"` 非法；`belowAverage` 并不存在。改为标准枚举（uniqueValues/duplicateValues/containsBlanks/notContainsBlanks/containsErrors/notContainsErrors，排名 rank=int 与 percent=bool，aboveAverage 用 aboveAverage 属性并入）。

### Notes

- 验证：497/497 单元测试通过，net48/net8.0 构建干净；真实 Excel COM 打开含条件格式工作簿无修复提示，规则保留完整（14 条）。

## [2.4.3] - 2026-08-22

### Added

- **CSV 分隔符选项**：`ExcelReadOptions.Separator` / `ExcelWriteOptions.Separator`（char?）。逗号 / 分号 / Tab 均可读写。
  - 读取默认 **null → 自动探测**：统计首段内容引号外的 `,`、`;`、Tab 频率取最多；三者均未出现回退逗号。显式指定则始终使用。
  - 写出默认 **null → 逗号**（与历史一致，零破坏）；含分隔符的字段自动引号包裹。
- **浮动图片读回**（xlsx/xlsm）：打开含 `oneCellAnchor`/`twoCellAnchor` 图片的工作簿，`Worksheet.Images` 自动回填 `WorksheetImage`（Row/Column/Placement=Floating/Name/AltText/Anchor/MoveMode/Data）。
  - 读回保持：img 字节往返、锚点位置、图片描述；写回不清除既有图片。
  - InCell richData 图片读回将随 2.4.4 一起提供。
- **条件格式**（xlsx/xlsm）：支持 `cellIs` / `expression` / `colorScale` / `dataBar` 四类规则的读写；`SheetData.ConditionalFormats` / `Worksheet.ConditionalFormats` 承载；xls/xlsb/csv 按降级回调上报。
  - `ConditionalFormat`（Sqref/Type/Formula/Formula2/Operator/Style/ColorScale/DataBar/Priority）
  - `ColorScaleInfo`（低/中高三色）
  - `DataBarInfo`（颜色/是否显示值/长度范围）

### Fixed

- **条件格式 cfvo 类型非法导致 Excel 修复提示**：dataBar 的 `<cfvo type="auto">` 不是 `ST_CfvoType` 合法值，colorScale 三色误用 `num 0/1/2` 阈值，Excel 打开会提示「已修复的部件 / XML 错误 / sheet1.xml」并丢弃规则。现改为 schema 合法的 `min` / `max`（dataBar）与 `min` / `percent 50` / `max`（三色色阶）。

### Notes

- 本版为 P1 全部并入（CSV 分隔符 + 图片读回 + 条件格式四类）。
- 全量 **487 测试通过**，net48 + net8.0 构建干净。

## [2.4.2] - 2026-08-21

### Added

- **统一降级报告机制（新公开 API）**：`ExcelWriteOptions.OnDegradation`（可选回调，默认 null，即历史行为不变）；新增 `DegradationInfo`（Capability / SheetName / TargetFormat / Message）与 `DegradationCapability` 枚举（16 项能力：批注/数据验证/筛选/图片/文档属性/命名区域/样式/合并/冻结/超链接/行高/列宽/公式/图表/透视表/InCell richData）。
  - 写出目标格式不支持某能力时，逐项回调上报，不再静默丢弃。
  - xls / xlsb：批注、数据验证、自动筛选、图片，以及完整单元格样式（仅保留数字格式，其余上报）。
  - CSV：样式/行高/列宽/合并/冻结/筛选/批注/验证/超链接/公式/图片逐项上报。
- **xlsb 保真透传**：打开时捕获保留部件（图表/主题/未知关系/内容类型声明），保存时原样透传；xlsb 打开-改-保存不再丢未知部件。
- **xlsb 文档属性**：读取 docProps（core/app）并写回；`WorkbookProperties` 在 xlsb 读/写往返闭环。
- **流式写入器补齐**（`Excel.CreateWriter`）：
  - 支持单元格样式与数字格式（styles.xml 延迟到 Close 生成，`s` 属性正确引用）。
  - 支持公式（`<f>` 元素 + 缓存值）。
  - 支持超链接（外部 r:id rels + 内部 location，Close 时落超链接区与 sheet rels）。**注意**：超链接需缓冲到 Close，数量极大时内存不再恒定。
  - 扩展名校验：仅允许 .xlsx/.xlsm；`.csv/.xls/.xlsb` 路径明确报错（不再静默写出内容）。
- **CSV 解析器重写**（字符级状态机）：
  - 引号字段内的换行不再错误拆行（读写对称）。
  - 空行保留（行号不再错位）。
  - LF / CRLF 均可。
  - 转义双引号 `""` 正确还原。
  - 未闭合引号抛出明确的 `FormatException`。
- **Excel.StreamRows / XlsxWriter.Append 格式门禁**：对 xls / xlsb / csv 显式报错（"该格式不支持流式读取/追加"），不再误报"不是有效的 xlsx 文件"。

### Fixed（打开-保存保真）

- **改工作表名不再丢图表/图片关联**：sheet rels 按数量判断结构变化，与表名解耦。
- **`XlsxWriter.Append` 保留全部部件**：xlsm 追加不再丢 VBA 宏与图表；xlsx 追加不再丢主题/表格/drawing。
- **InCell 图片 + 已有 richData 的文件再保存不再 ZIP 重名**：跳过逻辑纳入 InCell richData 全部条目。
- **命名区域与工作簿窗口视图保留**：`definedNames` / `bookViews` 原样回写 workbook.xml（schema 位序正确）。
- **陈旧 `calcChain.xml` 不透传**，并写 `<calcPr fullCalcOnLoad="1"/>`，告别 Excel 修复提示。
- **文本公式缓存值不再被覆盖**：`Cell.Formula` 独立承载公式串，`Text/Number/Date/Boolean` 恒为缓存值；旧写法（`IsFormula=true` + Text）由写入器兼容垫片继续支持。
- **列宽打开-保存双向**：`<cols>` 读取回填 + 稀疏列宽索引错位修复，xlsx/xlsm/xlsb/xls 四格式往返一致。
- **重复工作表名在低层 API 也显式报错**（与高层 `WorksheetCollection.Add` 行为一致）。

### Known limitations（范围决策）

- xls 的保真透传、xls 文档属性、xls/xlsb 图表读入读回：不实现。
- xls / xlsb 完整单元格样式（字体/填充/边框/对齐）：本期仅数字格式保留，其余经 `OnDegradation` 显式上报；不构造可能破坏文件的二进制样式记录。
- CSV 编码仅支持 UTF-8（含 BOM）；非 UTF-8 需调用方先转码（零依赖约束）。
- 图片读取仍仅写回（打开文件不回填 `Images`）。

### Notes

- 无破坏性 API 变更。`OnDegradation` 默认 null，老调用方零感知。
- 密码 / 加密 / 超链接 / 冻结窗格 / 图片行为与 2.4.1 一致。
- 全量 **469 测试通过**，net48 + net8.0 构建干净。

## [2.4.1] - 2026-08-19

### Added
- **图片细化锚点能力**（Floating 图片）：
  - `ImageMoveMode` 枚举：`MoveAndSizeWithCells`（随格移动+缩放，twoCellAnchor）/ `MoveButDontSizeWithCells`（随格移动不缩放，默认）/ `FixedPosition`（固定位置，editAs="absolute"）
  - `ImageAnchor` 类：`TopLeftCell`（A1 引用）+ `TopLeftOffsetX/Y`（EMU 偏移）+ `WidthPixels/HeightPixels` + `MoveMode`
  - `WorksheetImage.Anchor`（可选，设置后优先于 Row/Column）+ `AltText`（cNvPr descr 无障碍文本）+ 只读 `CellAddress`
  - `Worksheet.AddImage(byte[], ImageAnchor, extension?, name?, altText?)` 新重载
  - 写入侧 `BuildDrawingXml` / `MergeDrawingXml` 两处统一锚点渲染（oneCellAnchor/twoCellAnchor + editAs + 偏移 + descr）
  - `twoCellAnchor` 的 to 按默认列宽≈64px/行高≈20px 估算（随格缩放特性下初始尺寸≈设定像素）
  - Excel COM 验证三种模式 + AltText 正确、无修复提示

### Notes
- 向后兼容：现有 `Row/Column` `AddImage` 行为不变（默认 MoveButDontSizeWithCells，无 editAs）
- InCell 图片忽略 Anchor（richData 无锚点概念）
- 全量 **429 测试通过**，net48+net8.0 构建干净

## [2.4.0] - 2026-08-18

### Added
- **文件级安全（打开密码 / 修改密码）**：xlsx/xlsm/xlsb 三格式读写闭环。
  - **打开密码读取**：`Excel.Open(path, new ExcelReadOptions { OpenPassword = "..." })` 现可读取 Agile Encryption（AES-256-CBC/SHA512/spinCount=100000）加密工作簿。内部 `Internal/Encryption/AgileDecryptor.cs` 实现与 Excel 兼容的迭代哈希 + blockKey 派生（net48 兼容，无 PBKDF2 依赖）。
  - **修改密码识别**：识别 `<fileSharing>`（写保护），读取时**提供 ModifyPassword 即授权**（乐观授权；因 SHA-512 哈希跨 Excel 版本不稳定，不校验样本值）。读取后 `Workbook.Security.HasModifyPassword` 可判断写保护状态。
  - **密码保存与加密写出**：`ExcelReadOptions.OpenPassword` 打开后 `SaveAs` 默认继承打开密码；`wb.Security.SetOpenPassword("...")` / `wb.Security.SetModifyPassword("...")` 显式设/移除密码后写出。内部 `Internal/Encryption/OoxmlEncryptor.cs`（与解密对称）+ `Internal/Cfb/EncryptedCfbWriter.cs`（绝对扇区 FAT）。
  - `Workbook.Security`（`WorkbookSecurity`）：`HasOpenPassword` / `HasModifyPassword` / `HasModifyAccess` / `IsReadOnly` / `CanSave` / `ReadOnlyRecommended`。
  - 修改密码 = `<fileSharing>`（写保护，非 zip 加密）；`ModifyPasswordTouched` 时不透传原 fileSharing。
  - 密码**绝不**出现在异常/日志/测试输出。
  - 验证：真实 Excel COM 打开（A1=Hello Encrypted / B2=123.45 / C3=中文测试）、msoffcrypto 独立工具交叉解密成功。43 个密码测试（SecurityState 18 + OpenPasswordRead 22 + ModifyPassword 14 - 重叠）。
- **超链接（xlsx/xlsm/xlsb/xls 四格式）**：`Cell.Hyperlink`（`Hyperlink { Target, Tooltip, IsInternal }`）。
  - xlsx/xlsm：写出 `<hyperlinks>` + sheet rels，读取解析回填；内部跳转用 `location`（不走 External rel），外部 URL/文件/mailto/UNC 均可读写。
  - xlsb：BIFF12 `BrtHLink`（0x01EE）读写 + sheet `.bin.rels`（外部走 relId，内部走 location）。
  - xls：BIFF8 `HLINK`（0x01B8）+ `HLinkTooltip`（0x0800）读写，支持 URL Moniker 与内部跳转。
  - Excel COM 验证四格式超链接可点击、tooltip 正确、内部跳转 `SubAddress` 正确。
- **冻结窗格增强**：`Worksheet.FreezeRows` / `FreezeColumns`（`SheetData.FreezeRows/FreezeColumns` 承载），xlsx/xlsb/xls 三格式一致支持任意行列冻结。`FreezeHeader` 兼容为 `FreezeRows=1`。写出 `pane`（ySplit/xSplit/topLeftCell/activePane）+ 读回。Excel COM 验证三格式 `FreezePanes=True`、`SplitRow/SplitColumn` 正确。
- **图片写回（xlsx/xlsm）**：双模式——
  - **Floating 浮动图片**：`ws.AddImage(byte[] data, row, col, widthPx, heightPx, ImagePlacement.Floating)`，生成 `xl/drawings/drawingN.xml`（oneCellAnchor）+ media + drawing rels。打开已有图片的工作簿再 AddImage 会**合并**进既有 drawing（追加锚点 + rel），不产生 zip 重名或重复 drawing rel。Excel COM 识别 1 shape。
  - **InCell 嵌入图片**：`ws.AddImage(data, row, col, ImagePlacement.InCell)`，生成 richData 体系（metadata.xml + richValueRel + rdrichvalue + rdrichvaluestructure + rdRichValueTypes），单元格输出 `<c t="e" vm="n">`。Excel 无修复打开、值 = #VALUE!（与真实样本一致）。
  - 自动探测图片扩展名（PNG/JPEG/GIF/BMP）与像素尺寸（`Internal/ImageHeaders.cs`）；支持多 sheet、多图片、混合模式。

### Fixed（本轮 code review 清理）
- **xlsb 修改密码读写闭环**：读取侧解析 `BrtFileSharingIso`/`BrtFileSharing`（0x02A4/0x0224）→ 识别写保护与只读状态；写出侧生成 `BrtFileSharingIso` 记录。此前 xlsb 修改密码读写缺失（双密码 xlsb 仅给打开密码即可写）。
- **只读绕过修复**：`WorkbookSecurity.SetModifyPassword` / `ClearAll` 在未获得修改权限（`HasModifyAccess=false`）时抛异常，防止未授权剥离/替换写保护。
- **内部超链接 OOXML 修正**：`IsInternal` 链接改为写 `location` 属性（不写 External rel），读取按 `location` 判定内部、scheme 判定外部（修复 `mailto:`/`file://`/UNC 误判）。
- **解密正确性**：verifier 哈希按 `hashSize` 截断比对（支持 SHA-1/256/384 Agile）；解密结果按 `dataSize` 截断（去掉 AES 零填充）；校验 `dataIntegrity` HMAC（EncryptedPackage 被篡改时明确报错）。
- **非加密文件误传 OpenPassword**：`Excel.Open` 提供 OpenPassword 时先判定是否加密工作簿，非加密文件给明确异常而非晦涩 CFB 错误。
- **图片 zip 重名**：AddImage 到含既有 drawing/media 的文件时跳过保留序号、合并 drawing，避免 `ZipArchive` 重名异常与重复 drawing rel。
- **Cell.CopyFrom 共享 Hyperlink 引用**：改为深拷贝（`Clone()`）。

### Changed
- `Worksheet.FreezeHeader` 语义：现为 `FreezeRows = 1` 的便捷别名，读取时若 ySplit=1 仍回填 `FreezeHeader=true`（向后兼容）。

### Notes / 兼容性
- 既有 API 无破坏性变更（新增均为增量属性/重载）。
- **图片仅写回（xlsx/xlsm）**：打开文件不会回填 `Images`；图片读取不在 2.4.0 范围。
- **已知限制**：xls/xlsb 图片不在 2.4.0 范围；xls 老格式密码（RC4 XOR）不支持。
- 真实文件验证：加密样本（`files/打开修改都需要密码.xlsx` 等）解密/写出经 Excel COM + msoffcrypto 双验证；图片样本（`files/图片.xlsx`）结构对齐；超链接/冻结窗格四格式经 Excel COM 验证。
- 全量 **423 测试通过**，net48+net8.0 干净。

## [2.3.0] - 2026-08-17

### Added
- **`Excel.Open(Stream, format)` 对象模型 Stream 打开**：新增 `Excel.Open(Stream stream, ExcelFormat format, ExcelReadOptions? options)` 重载，支持五格式从流读取。必须显式指定格式（流无扩展名）；输入流不关闭（由调用方管理）；支持不可定位流（内部复制到内存）；打开后 `CurrentPath` 为 null，需 `SaveAs` 指定保存路径。与 `Workbook.Save(Stream, format)` 配对，五格式 Stream 读写闭环。
  - 底层后端新增 Stream 重载：`XlsbBackend.ReadVbaProject/ReadWorkbookCodeName/ReadDate1904(Stream)`、`XlsBackend.ReadDate1904(Stream)`、`CsvBackend.Read(Stream, sheetName)`。
  - 新增 `StreamOpenTests` 15 个（五格式 Stream 往返、不可定位流、流不关闭、CurrentPath=null、SaveAs 可用、1904 保留、加密识别含 `<stream>` 显示名回归、参数校验）。
- **加密文件识别**：带打开密码的 xlsx/xlsm/xlsb（OLE CFB 容器，含 `EncryptionInfo`/`EncryptedPackage` 流）与加密 `.xls`（BIFF8 `FILEPASS` 记录）打开时现可识别并抛 `LiteExcelException`（"文件已加密（带打开密码）"），不再误报为 zip 损坏或解析出乱数据。完整密码读写规划在后续版本。
  - 新增 `Internal/EncryptionDetector.cs`（CFB 魔数嗅探 + `EncryptionInfo` 流检测，复用 `CfbFile`）；`Excel.Open` 的 xlsx/xlsm/xlsb 路径在进 zip 前先识别。
  - 加密识别现已覆盖**所有公开 path 读取入口**：`Excel.Open`、`Excel.Read<T>`、`Excel.ReadAsDataTable`、`Excel.GetSheetNames`、`Excel.StreamRows`、`XlsxReader.Read/ReadAll/GetSheetNames/StreamRows/ReadWithProgress/ReadProperties`。
  - 新增 `EncryptedWorkbookTests` 17 个（真实 Excel 生成的 4 个加密 fixture + 公开读取入口覆盖）。
- **1904 日期系统写出**：`Workbook.Date1904`（打开时捕获）现可在 xlsx/xlsb/xls 写出侧写回标志并保持日期序列一致，修复 1904 工作簿往返偏移 4 年的缺陷。
  - `XlsxWriter` 写 `<workbookPr date1904="1"/>`；`XlsbWriter` 写 `BrtWbProp` flags bit0；`XlsWriter` 写 `DATE1904` 记录。
  - 日期序列换算统一到 `FormatDetector.DateToSerial(date, date1904)`（1904 基准 = OADate - 1462）。
  - Excel COM 验证：LiteExcel 生成的 1904 xlsx/xlsb 无修复提示，序列值正确（2024-03-15=43904、1904-01-01=0）。
  - 新增 `Date1904Tests` 6 个（含真实 Excel fixture `excel-authored-date1904.xlsb`）。
- **`Excel.Write` 扩展名推断一致化**：现与 `DetectFormat` 完全一致——`.xls` 扩展名转 Xls、`.xlsb` 转 Xlsb（此前忽略这两个扩展名），规则简单可预测。
- 新增 `DegradationBehaviorTests` 8 个（扩展名推断、宏保护 xlsx/xls、Stream 宏保护、xlsm/xlsb 宏仍可用、CSV 多表报错等）。

### Changed
- **宏保护扩展到 `.xlsx`**：含 VBA 宏的工作簿 `SaveAs` 到 `.xlsx` 或 `.xls`（不支持宏）现抛 `LiteExcelException`（在创建文件前拦截），防止宏被静默丢弃或生成不一致文件。含宏工作簿请保存为 `.xlsm` 或 `.xlsb`。无宏工作簿不受影响。

### Notes / 兼容性
- 既有 API 无破坏性变更（`Date1904` 为 internal 属性，不暴露公开 API；`XlsWriter.Write`/`XlsbWriter.Write`/`XlsxWriter.Write` 的 `date1904` 均为带默认值的可选参数）。
- 加密文件此前会误报为 zip 损坏；2.3.0 起抛明确 `LiteExcelException`。这是错误信息改善，非行为破坏。
- **net48 兼容性修复**：Stream 打开加密文件时，错误信息中的显示名 `"<stream>"` 在 net48 下会被 `Path.GetFileName` 判定为非法路径字符而抛 `ArgumentException`（net8.0 不抛）。现改用 `SafeDisplayName` 兜底，net48 下正常抛 `LiteExcelException`（由外部 net48 验证程序发现并修复）。
- 含宏工作簿保存为 `.xlsx` 此前可能生成包含 `vbaProject.bin` 但主文档类型为普通 xlsx 的不一致文件；2.3.0 起明确抛错。这是保护性变更。
- 真实文件验证：Excel COM 打开 1904 xlsx/xlsb 无修复、日期正确；SheetJS 交叉验证 xlsb 大文件数据一致（10k/50k 行、中文、emoji、特殊字符、合并、冻结）。
- 全量 **302 测试通过**（256 + 15 Stream Open + 17 加密 + 6 个 1904 + 8 个降级行为），net48+net8.0 干净。

## [2.2.6] - 2026-08-17

### Fixed
- **`xlsm` 保存后 Excel 打不开（issue #1）**：写 `[Content_Types].xml` 时 `/xl/workbook.xml` 的主文档类型写死为 `sheet.main+xml`，保存 `.xlsm` 未切换为 `macroEnabled.main+xml`，Excel 校验扩展名与内容类型不一致拒绝打开。现按格式/扩展名正确写出 `application/vnd.ms-excel.sheet.macroEnabled.main+xml`（`Workbook.SaveAs`、`XlsxWriter.Write(path)`、`XlsxStreamWriter.Create(path)` 三个入口均覆盖）。PR #3 贡献。
- **带宏 `xlsm` 经保存后 VBA 模块错位失效（issue #4）**：2.2.1 的宏保留只透传 `vbaProject.bin` 字节，重建 workbook.xml/sheet XML 时丢失 `workbookPr@codeName` 与 `sheetPr@codeName`，宿主失去绑定后被 Excel 重命名（`ThisWorkbook1`、事件宏静默失效）。现于打开时捕获、保存时按 schema 位置写回这两个 codeName（`SheetData.CodeName` 新公开属性承载工作表级）。PR #5 贡献。
- **`XlsxStreamWriter` 写出的文件 Excel 打不开**：两处问题——
  1. styles.xml 的 fills 仅含 1 个 `none` 填充，缺少规范要求的前置 `gray125` 项（与 2.1.1 主写入器同款修复），现改为 `none` + `gray125` 两项；
  2. 单元格引用 `r` 写死为第 1 行（`CellRef.ToString(0, ...)`），所有行都写成 A1/B1/... 且 `<row>` 缺 `r` 属性，Excel 严格校验即拒开。现按实际行号写出 `<row r="n">` 与 `r="An"`。
- **PR #3 合并遗漏**：`XlsxStreamWriter` 构造函数引用了未声明的 `_macroEnabled` 字段，补上字段声明。
- **`xls`→`xlsb` 行高错误（数据"丢失"）**：`XlsBackend.ParseRowHeight` 读 BIFF8 `ROW` 记录的 `miyRw` 时偏移错误（读了 `colMac` 位置，应为 offset 6），导致源 `xls` 行高被误读为 `colMac/20`（用户文件读出 0.65pt 而非 15pt），写出 `.xlsb` 后行高塌缩、Excel 打开疑似"只剩空表"。现按 `rw(0)+colMic(2)+colMac(4)+miyRw(6)` 正确解析。Excel COM 验证 `xls`→`xlsb` 后行高 30pt 正确保留、行不隐藏。
- **`SaveAs` 扩展名与格式不匹配时静默产出错误文件**：`Workbook.SaveAs(path, format)` 现校验扩展名与格式一致，不匹配抛 `LiteExcelException`（明确失败优于静默写错格式）。
- **`xlsm`→`xlsb` 宏丢失**：xlsb 写入接入 `vbaProject.bin` 保留（Content_Types Override + workbook.bin.rels 关系）与 workbook/sheet codeName 写回；`Excel.Open` 的 xlsb 路径同步捕获 vbaProject 字节与 codeName。Excel COM 验证转换后 VBA 工程组件数与源一致。

### Added
- **`Excel.Create(string[] sheetNames, format)` 批量建表重载**（PR #2 贡献）：传 null 或空数组保留默认 Sheet1，重名抛 `LiteExcelException`；README 中英 API 表同步。
- **`xlsb` 写入后端**：`wb.SaveAs("file.xlsb", ExcelFormat.Xlsb)` 现可写出 BIFF12 工作簿（至此 xlsx/xlsm/csv/xls/xlsb 五格式读写闭环完成）。
  - 多工作表（中文名）、文本/数字/日期/布尔单元格、共享字符串表、数字格式、合并单元格、列宽、行高、冻结表头。
  - 公式单元格按缓存结果值静态写出（公式文本不保留，与 xls 写入一致）。
  - 新增 `Internal/XlsbWriter.cs`（记录级写序列对照 Excel 原生输出实证：`BrtWbProp` 必须含 codeName 字段、`BrtWsProp` 为工作表首个必选记录、`BrtPane` 冻结 topLeftCell 行=1、Short 单元格仅在同一行连续列复用）。
  - `Excel.Create(ExcelFormat.Xlsb)` / `Excel.Create(ExcelFormat.Xls)` 现均可用。

### Notes / 兼容性
- 既有 API 无破坏性变更（`SheetData.CodeName` 为纯增量，普通文件为 null 不影响现有行为）。
- **真实文件验证**：
  - 本机 Excel COM 打开修复后的 `.xlsm`（`SaveAs(ExcelFormat.Xlsm)` 输出）与流式写入器输出的 `.xlsx`，均无修复提示、单元格值正确（中文、数字、日期 OADate）。
  - `xlsb` 写入：Excel COM 打开含中文/数字/日期/布尔/合并/冻结/列宽/行高/多表（中文名）的 `.xlsb` 输出无修复提示且值正确；Excel 打开后另存为 `.xlsb`，LiteExcel 再读回逐值一致；SheetJS 独立交叉验证一致（表名、值、合并范围）。
  - `xls`→`xlsb`：Excel COM 打开输出文件确认行高与行可见性正确（`ParseRowHeight` 偏移修复）、行列值与源一致；`xlsm`→`xlsb`：Excel COM 确认 VBA 工程组件数与源一致。
- 全量 256 测试通过（249 + `XlsbWriteTests` 7，并将 3 个"xlsb 写入不支持"旧测试改为往返/跨格式断言），net48+net8.0 干净。

## [2.2.5] - 2026-08-15

### Added
- **公式文本解析**：`xls`（BIFF8）与 `xlsb`（BIFF12）读取时，公式单元格的 RPN 现可解析为 A1 文本。
  - 支持单元格引用（A1/$A$1）、区域（A1:B2）、数字/字符串/布尔/错误常量、算术与比较运算符、括号、常见内置函数（`SUM`/`IF`/`ROUND`/`MAX` 等，含 `PtgAttrSum` 快捷写法）。
  - 公式通过 `Cell.IsFormula` 与 `Cell.Text` 暴露（缓存结果值仍保留在数值字段）。
  - 不支持的公式（数组公式、3D 引用、命名区域等）安全降级为仅缓存结果值。
- 新增 `Internal/Biff/FormulaParser.cs`（RPN→A1 解析器）与 `Internal/Biff/FormulaFtab.cs`（BIFF8 内置函数表）。
- 新增 `Fixtures/excel-formulas.xls`（Excel 生成，含 10 条常见公式）与 `FormulaTests`（3 个）。

### Notes / 兼容性
- 既有 API 无破坏性变更；xls 写入仍按计划将公式降级为静态缓存值。
- 真实文件验证：Excel 生成的 10 条公式（含 `=IF(A1>5,1,0)`、`=CONCATENATE(A1,B1)` 等）全部正确解析；真实 xls/xlsb fixture 的 `=B2*2` 均正确返回。
- 全量 230 测试通过，net48+net8.0 干净。

## [2.2.4] - 2026-08-15

### Added
- **`xls` 写入后端**：`wb.SaveAs("file.xls", ExcelFormat.Xls)` 现可写出 BIFF8 工作簿。
  - 多工作表（中文名）、文本/数字/日期/布尔单元格、合并单元格、列宽、行高、冻结表头、自定义数字格式。
  - 公式单元格按缓存结果值静态写出（公式文本不保留，已知限制）。
  - 输出为 OLE2/CFB 容器（`Internal/Cfb/CfbWriter.cs`）+ BIFF8 记录（`Internal/Biff/XlsWriter.cs`），零依赖、net48 兼容。

### Notes / 兼容性
- 既有 API 无任何破坏性变更；`xlsb` 写入仍未实现，图片、图表、透视表等高级能力不在本版本范围。
- **真实文件验证**：用 Excel COM 打开 LiteExcel 写出的 .xls（含 3000 行×3 列中文数据、日期、布尔、合并、冻结、列宽、公式结果），值全部正确。
- 修复过程中以 SheetJS/真实文件为基准实证了多项 BIFF8 关键细节：BOF 最低兼容版本字段、BIFF8 DIMENSIONS 行宽为 4 字节、必须写出 FORMAT 记录与 16 个内置样式 XF、COLINFO 为 12 字节、SST 续接段不得切裂 UTF-16 字符、WINDOW2 标志等。

## [2.2.3] - 2026-08-15

### Added
- **`xlsb` 读取后端**：`Excel.Open("file.xlsb")` 现可读取二进制 OOXML 变体 `.xlsb` 文件。
  - 数据单元格：文本 / 数字 / 日期（1900/1904 系统、内置与自定义格式识别）/ 布尔 / 错误。
  - 共享字符串表（SST）、合并单元格、列宽（BrtColInfo）、行高（BrtRowHdr）、冻结表头（BrtPane）。
  - 公式单元格返回缓存结果值；公式文本暂不解析（已知限制）。
  - `xlsb` 写入暂不支持，`SaveAs` 到 `.xlsb` 抛 `NotSupportedException`。

### Added（内部实现）
- 新增 BIFF12 记录读取器 `Internal/Biff12/Biff12Records.cs`（LEB128 变长记录头，无 Instance 字段）与读取后端 `Internal/XlsbBackend.cs`。
- 抽取共享逻辑：`Internal/FormatDetector.cs`（数字格式→日期识别）、`Internal/BiffShared.cs`（RK 数值、错误码），供 xls / xlsb 后端复用。

### Notes / 兼容性
- 既有 API 无任何破坏性变更；图片、图表、透视表等高级能力不在本版本范围。
- `xlsb` 读取保持 AOT 友好（无反射）与 net48 兼容。
- **真实文件验证**：新增由 Microsoft Excel 生成的 fixture `excel-authored.xlsb`（中文表名、合并、冻结、列宽、日期格式、公式、3000 行唯一字符串），读取结果与期望值逐单元格一致（9011 行全匹配）。
- 真实文件验证还暴露并修复了一个仅真实文件才可见的问题：日期单元格以 RK 压缩数值存储时未做日期识别，以及 `BrtCellXfs` 索引基线错位导致样式指向偏移。

## [2.2.2] - 2026-08-15

### Added
- **`xls` 读取后端**：`Excel.Open("file.xls")` 现可读取传统 `.xls` 文件（OLE2/CFB 复合文档 + BIFF8，Excel 97+）。
  - 数据单元格：文本 / 数字 / 日期（1900/1904 系统、内置与自定义格式识别）/ 布尔 / 错误。
  - 共享字符串表（SST）：支持压缩与 UTF-16 字符串，以及跨 `CONTINUE` 记录续接（含续接段重新声明编码）。
  - 合并单元格、列宽（COLINFO）、行高（ROW）、冻结表头（Pane）。
  - 公式单元格返回缓存结果值；公式文本暂不解析（已知限制）。
  - `xls` 写入暂不支持，`SaveAs` 到 `.xls` 抛 `NotSupportedException`。

### Added（内部实现）
- 新增 OLE2/CFB 容器解析器 `Internal/Cfb/CfbFile.cs`（FAT / DIFAT / 目录 / mini stream）。
- 新增 BIFF8 记录读取器 `Internal/Biff/BiffRecords.cs`、Unicode 字符串续接读取器 `Internal/Biff/BiffStringReader.cs`、读取后端 `Internal/Biff/XlsBackend.cs`。

### Notes / 兼容性
- 既有 API 无任何破坏性变更；`xlsb` 仍未实现，图片、图表、透视表等高级能力不在本版本范围。
- `xls` 读取保持 AOT 友好（无反射）与 net48 兼容。
- **真实文件验证**：新增由 Microsoft Excel 生成的 fixture `excel-authored.xls`（中文表名、合并、冻结、列宽、公式、3000 行唯一字符串强制 SST 跨 CONTINUE），读取结果与期望值逐单元格一致（9011 行全匹配）。

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
