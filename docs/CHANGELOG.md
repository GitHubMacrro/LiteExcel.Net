# Changelog

## [2.4.78] - 2026-09-18

### Added

- **xlsb → xlsx/xlsm 跨格式转换（阶段 A：格式无关部件直通）**：打开 xlsb 后另存为 xlsx/xlsm 时，不再丢弃全部保留部件，而是把两种容器中同构的「格式无关」部件原样保留：工作簿主题 `theme1.xml`、`customXml/*`（含 Power Query M 代码）、`docProps/custom.xml`、VBA 宏工程（xlsm）、媒体 `xl/media/*`、ActiveX 控件、打印设置 `printerSettings/*`、控件属性 `ctrlProps/*`、VML 绘图 `vmlDrawing*`。各部件的关系（`_rels/.rels`、`workbook.xml.rels`、工作表 rels）与 `[Content_Types].xml` 按目标容器重写，并剔除指向未写出部件的悬空引用。
- **xlsb → xlsx/xlsm 跨格式转换（阶段 C：Power Query / Power Pivot 连接）**：`xl/connections.bin` / `xl/queryTables/queryTableN.bin` 转码为 OOXML XML（`connections.xml` / `queryTableN.xml`），`xl/model/item.data` 直通，并合成数据模型链接表所需的 `_xlcn.LinkedTable_*` 定义名。转码输出与 Excel 自身转换逐元素一致。
- **xlsb → xlsx/xlsm 跨格式转换（阶段 B2：基础样式）**：解析 `styles.bin` 的 `BrtFont`/`BrtFill`/`BrtBorder`/`BrtXF` 并按 `ixfe` 应用到单元格。
- **xlsb → xlsx/xlsm 跨格式转换（阶段 B3：绘图 / ActiveX 形状）**：`xl/drawings/drawing*.xml` 的 xlsb 专有表达（`xdr:graphicFrame` + `com14:compatSp`）转码为 OOXML `xdr:sp`（`a14:compatExt spid` + `a14:hiddenLine`），使含 ActiveX 形状的工作簿能被 Excel 打开。
- **xlsb → xlsx/xlsm 跨格式转换（阶段 D：透视表）**：`xl/pivotCache/pivotCacheDefinitionN.bin` 与 `xl/pivotTables/pivotTableN.bin` 转码为 OOXML（头 / location / pivotFields / items / rowFields / rowItems / colFields / colItems / dataFields / pivotHierarchies / 样式，以及 OLAP 数据模型的 dimensions / measureGroups / maps / cacheHierarchies / 数值·日期 sharedItems / slicerData 缓存）。合成 workbook `<pivotCaches>`（cacheId 按 workbook.bin 缓存列表位置映射）与 x14 extLst。
- **xlsb → xlsx/xlsm 跨格式转换（阶段 E：切片器）**：`xl/slicerCaches/slicerCacheN.bin` 与 `xl/slicers/slicerN.bin` 转码为 OOXML，接线 CT / workbook rels / workbook extLst（x14 slicerCaches）/ 工作表 rels / 工作表级 `x14:slicerList`；保留原始 xlsb `sheetId` 以保证切片器 `tabId` 链接。
- **XLSB 手术式编辑（修改内容 / 改表名 / 添加表，保真保留高级部件）**：源为含透视表/切片器/连接/Power Query/数据模型/宏的 xlsb 时，改动后另存不再走「整本重建」（会丢 sheet 记录、产生孤儿/重复部件、致 Excel 拒开），而是逐字节保留源包全部部件，只补丁改动处：
  - **修改单元格内容**：对受影响的 `sheetN.bin` 做字节级单元格记录补丁（保留行头、表格部件引用、what-if、打印设置、`binaryIndex` 引用等）；改动落在超级表表头行时同步更新该表列名（`BrtBeginListCol.stCaption`）。
  - **改表名**：改写 `workbook.bin` 的 `BrtBundleSh` 名称。
  - **添加表（仅追加末尾）**：生成新 `sheetN.bin`，追加 `BrtBundleSh` + worksheet rel + Content-Types override。
  - 通过 `Worksheet.ModifiedCells` 跟踪被改单元格；仅当打开时的表按原位置保留且只发生「内容修改 / 改名 / 末尾追加」时启用（可见性、标签色、导入等其它修改仍回退重建）。经 Excel COM 正常打开验证：任意表内容修改、改名、添加表均无修复/闪退。

### Fixed

- **跨格式静默丢弃高级部件**（违反保真契约）：源为 xlsb、目标为非 xlsb 时，含透视表/切片器/连接/Power Query/数据模型等高级部件此前会被静默丢弃。阶段 A 改为显式上报 `DegradationCapability.PivotTables`（记入 `Workbook.SaveDegradations`），`AllowFeatureLossOnSave=false` 时阻止保存；阶段 C/D/E 落地后这些部件已可转码保留，对 xlsx/xlsm 目标不再上报/阻止（可用 `LITEXCEL_DISABLE_PIVOT_WIRING=1` 回退到丢弃+上报）。
- **高级部件检测扩展**：`SourceHasAdvancedXlsbParts` 现同时覆盖工作簿级部件（`pivotCache` / `pivotTables` / `slicerCaches` / `slicers` / `queryTables` / `connections` / 数据模型），而非仅工作表 rels 中的透视/切片引用。

### Notes

- **`xl/drawings/drawing*.xml` 暂不直通**：该部件是 xlsb 专有的 ActiveX 图形表达（`xdr:graphicFrame` + `com14:compatSp`），混入 xlsx 会被 Excel 拒绝打开（`0x800A03EC`）；阶段 B3 已做 drawing 转码。`vmlDrawing*` 保留。
- **透视表/切片器接线默认启用**：`raw_repro.xlsb` → xlsx/xlsm 经 Excel 打开验证为 9 表 + 4 透视表、无修复。剩余保真差异（不影响打开）：pivotFields 翻倍怪癖、pivotTable `<formats>`/`<extLst>`、definedNames 中的结构化表引用（`ptgElfLel`）暂未解码。
- **⚠️ 已知限制：删除含 Power Query 连接的 xlsb 工作表，部分场景输出会被 Excel 拒开**。删除某工作表时，若该表是某个 Power Query 连接的数据源，Excel 会对 `workbook.bin` 做**整套「另存为」规范化**：重建定义名表（`_xlcn.LinkedTable_*` 去尾缀、为被删表新建全局占位名如 `CriteriaValue`/`DataSelectionCriteria`、把失效名的 rgce 改写成 `PtgName` 重指向、重排/重编号），并同步重建 `BrtExternSheet`(XTI) 表、重映射所有 3D 引用的 `ixti`、重编号 pivot/sheet 部件、重编码 `styles`/`sharedStrings`。库当前仅做**手术式**改动（标记失效化、局部重编号），**无法完全复刻**该规范化，因此：
  - **可用**：删除 `DayList`、以及 `DayList` + `Data Selection Criteria`（已在 Excel 验证正常打开）；
  - **不可用（会被 Excel 报修复或闪退）**：单独删除其余工作表（如 `BE Production Yield`、`FTLRGroupName`、`ConnectionInfo`、`BE Lot Loss Details`、`BE Lot Loss Details Data`、`Data Selection Criteria`、`ProductGroupData`、`BUSP_VW_BEPRODUCTIONYIELD Data`）。
  - **安全网**：删除此类文件的工作表时，库会经 `Workbook.SaveDegradations` 上报 `DegradationCapability.PivotTables` 警告（"输出可能在 Excel 触发修复/闪退"），调用方可据此提示用户。**未做静默丢弃，但也无法保证输出可打开。**
  - **建议**：删除含 PQ 数据模型的工作表属于高风险操作；如需任意删表，请改用 Excel COM 直接执行删除，或删除后用 Excel 验证输出。

### Tests

- 新增 `XlsbToXlsxConversionTests`（7 项）：格式无关部件直通 + 关系重写；drawing XML 排除 / VML 保留；透视表部件转码保留；VBA 按目标格式取舍。含包完整性断言（无悬空关系、部件均有内容类型声明）。
- 新增 `XlsbConnectionTranscoderTests`、`XlsbPivotCacheTranscoderTests`、`XlsbPivotTableTranscoderTests`、`XlsbSlicerTranscoderTests` 等阶段 C/D/E 转码测试（含真实文件断言，CI 无样本时跳过）。
- 全量 **725** 测试通过（net8.0）；net48 构建通过。

## [2.4.77] - 2026-09-18

### Added

- **XLSB 条件格式写入（全类型）**：`Worksheet.ConditionalFormats` 现对 xlsb 重建写出生效，覆盖**全部 18 种 OOXML 规则类型**：cellIs / expression / colorScale / dataBar / iconSet / top10 / aboveAverage / belowAverage / containsText / beginsWith / endsWith / notContainsText / uniqueValues / duplicateValues / containsBlanks / notContainsBlanks / containsErrors / notContainsErrors / timePeriod / textLength，含 `Style` 的 dxf 填充/字体/边框。
  - 记录层级（[MS-XLSB]，经真实 Excel 样本逐字节标定）：`BrtBeginCF`(0x01CD) → `BrtCFRule`(0x01CF) → [类型子记录] → `BrtEndCFRule`(0x01D0) → `BrtEndCF`(0x01CE)。
  - 类型子记录：iconSet = `BrtBeginIconSet`(0x01D1) + `BrtCFVO`(0x01D7)×N + `BrtEndIconSet`(0x01D2)；colorScale = `BrtBeginColorScale`(0x01D5) + `BrtCFVO`×N + `BrtColor`(0x0234)×N + `BrtEndColorScale`(0x01D6)；dataBar = `BrtBeginDataBar`(0x01D3) + `BrtCFVO`×2 + `BrtColor`(0x0234) + `BrtEndDataBar`(0x01D4)。
  - 条件格式样式（dxf）写入 `styles.bin` 的 `BrtDXF`(0x01FB)：填充色/字体色/字重/斜体/下划线/删除线/边框，经全局去重（与规则 `dxfId` 一致）。
  - 关键标定点：**BIFF12 记录头为 LEB128 变长编码**（非固定 2+2 或 2+4）；**CF 公式的单元格/区域引用相对 sqref 左上角锚点**存储，且 Ptg 字节与单元格公式不同（`Ref`=0x4C、`Area`=0x2D、`Func`=0x41、`FuncVar`=0x42）；`BrtCFRule` 的公式段展开长度 = 长度 + 2×引用数；colorScale/dataBar 的颜色记录为 `BrtColor`(0x0234)；`BrtBeginDataBar` = `minLength(u8) + maxLength(u8) + flags(u8)`（bit0 = showValue）；timePeriod 各时间段的 rgce 为 Excel 固定模板（与锚点无关）。
  - 经 Excel COM 打开验证：无文件级修复提示，16 条规则（覆盖 12 种类型）往返（xlsb→xlsx 另存）XML 与原语义一致。
- **工作表可见性 `Worksheet.Visible`**：新增 `SheetVisibility` 枚举（`Visible` / `Hidden` / `VeryHidden`）。支持 xlsx / xlsm / xlsb / xls 四格式读写（xlsx 走 `<sheet state>`、xlsb 走 `BrtBundleSh.hsState`、xls 走 `BOUNDSHEET.grbit`）。将最后一张可见表设为隐藏会抛 `LiteExcelException`（Excel 要求至少保留一张可见表）。
- **工作表标签颜色 `Worksheet.TabColor`**：`#RRGGBB` / `RRGGBB`，null = 无。xlsx / xlsm 读写（`sheetPr/tabColor`）；xlsb / xls 写出时经 `DegradationCapability.SheetVisibility` 上报后丢弃。
- **XLSB 命名区域读回**：打开 .xlsb 时解析 `BrtDefinedName` + rgce 令牌（`PtgRef`/`PtgArea`/`PtgRef3d`/`PtgArea3d` + 常量），填充 `Workbook.Names`（含 sheet-local `LocalSheetId`）。复合表达式（函数调用、运算符、名称引用等）跳过，不产出错误引用。
- **XLSB 超级表（Table/ListObject）读写**：读 `BrtBeginList`/`BrtBeginListCol`/`BrtTableStyleClient` + sheet `BrtTablePart` + rels；写 `xl/tables/tableN.bin` + `BrtBeginTableParts` + sheet rels + Content-Types。`Worksheet.AddTable` 现对 xlsb 生效。
- **XLSB 数据验证读写**：读 `BrtDVal`（列表/整数/小数/日期）；写 `BrtBeginDVs`/`BrtDVal`/`BrtEndDVs`。xlsb 写出数据验证不再降级上报。
- **XLSB 浮动图片写入**：`Worksheet.AddImage(..., ImagePlacement.Floating)` 现对 xlsb 生效。xlsb 的 `xl/drawings/drawingN.xml`、`drawingN.xml.rels`、`xl/media/imageN.*` 与 xlsx 同为标准 XML/二进制部件，复用 `XlsxWriter.ImagePlan` 生成；工作表经 `BrtDrawing`(0x0226) 引用 sheet rels 的 drawing 关系，`[Content_Types].xml` 声明 media Default + drawing Override。打开-保存（verbatim/手术式）时既有图片原样透传不翻倍。**InCell 图片**（richData 体系，需 BIFF12 `metadata.bin` + 单元格 vm）在 xlsb 维持 C 级降级上报。此变更经 Excel COM 打开验证（无修复提示，识别 1 个图片形状）。

### Fixed

- **XLSB 删表后另存导致 Excel 崩溃（0xc0000005）**：修复了手术式删除在含 Power Pivot / Power Query 数据模型的工作簿上产出损坏文件、Excel 打开即崩溃的问题；现已能**完整保真**复刻 Excel「删除工作表后另存」的行为。
  - **根因**（以 Excel COM 基准产出逐字节对照定位）：Excel 删除被数据模型引用的工作表时，会同步完成一整套改动，缺任一步都会使 Excel 打开时在 `EXCEL.EXE` 内部访问违例（`0xc0000005`）崩溃或拒绝打开：
    1. 剩余 sheet/table/binaryIndex 部件**重编号**（消除编号空洞，如 `sheet9→sheet8`）；
    2. `workbook.bin.rels` 的 rId **全量重编**（`rId9+` 递减，无空洞）；
    3. `BrtExternSheet`（XTI）引用被删表的条目 `itab` 置 `-1`、其余递减；
    4. 定义名：挂被删表的 sheet-local 名移除、其余 `itab` 递减；引用被删表的名称 **rgce 失效化**（`18 19` + ixti 指向失效条目 → body 改写为 `10 FF FF FF FF`）；
    5. `_xlcn.LinkedTable_*` 连接名/定义名**原样保留**（尾部数字为表名后缀，非冗余，不可剥离）；
    6. `[Content_Types].xml` 移除被删表 override 并同步重编号；
    7. 透视表 `BrtBeginPivotTable` 名称长度字段越界时规范化（Excel 修复源文件异常值）。
  - **修复**：重写 `WriteSurgicalXlsb` 完整实现上述 7 项同步改动。删表后**数据模型、PQ 连接、透视表/缓存、超级表、customXml、ActiveX、绘图、VBA 宏、工作表标签颜色、样式等全部保留**，输出与 Excel 自身产出等价（`workbook.bin` 仅剩连接 GUID / activeTab 两处无差异），Excel 打开正常。
- **XLSB 手术式删表遗漏 rId 重编号（`0x046D` / `0x0430`）**：删除工作表后 `xl/workbook.bin.rels` 会整体重编号，但 `workbook.bin` 内嵌的透视缓存引用 `0x046D`（→ `pivotCacheDefinitionN.bin`）与切片缓存引用 `0x0430`（→ `slicerCacheN.bin`）此前未同步重编号，致其指向错部件（实测指向 `slicerCacheN.bin` / `theme1.xml`）；Excel 打开判为断链并删除 `pivotTable1/2.bin`、`slicerCache2/3.bin`、`slicer1.bin` 及一条工作簿属性记录（"已删除的部件/记录"修复提示）。现按 `newRel` 映射重写这两类记录内嵌的 rId（布局 `flags(u32) + cch(u16) + rId(UTF16) [+ 尾部 u32]`，`cch` 为 **u16**，区别于 `0x0182` 的 u32）。
- **XLSB 透视表 `cacheId` 归位错误**：`BrtBeginPivotTable`(0x0118) 的 `cacheId`(off28) 此前在越界时按透视表名的数字后缀猜测；其真实语义是「该透视表 rels 指向的 `pivotCacheDefinitionN` 在 workbook 缓存引用序列（`0x0182` / `0x046D` 出现顺序）中的 **0 基索引**」。脏值会被 Excel 判为断链并删除该透视表（实测 `pivotTable1/2.bin`）。现由透视表 rels 解析缓存目标、查得索引后写回该字段。
- **XLSB 手术式删表改写 `_xlcn.LinkedTable_*` 名**：此前对连接名/定义名去掉尾部 `"1"`（旧称"冗余"）。实测该 `1` 是表名后缀（如 `Table1`），剥离会截断合法名（`Table1`→`Table`），且源名本身可能已是 Excel 形态（`Table11`），多次另存会逐次降级（`Table11`→`Table1`→`Table`）。经全文件扫描确认该名为纯元数据、无任何部件按字符串引用（`item.data` 逐字节不变、`queryTableN.bin` 不含该名、公式引用走索引 `PtgName`），Excel 另存时自行追加 `1`。现改为**原样保留源名**，不再改写。
- **`lengthIs` 非法 OOXML 类型（xlsx / xlsb）**：`ConditionalFormatType.TextLength` 此前在 xlsx 写出为 `<cfRule type="lengthIs">`——`lengthIs` **不在 OOXML `ST_CfType` 枚举**中，Excel 打开含该规则的文件会直接拒绝（经真实 Excel COM 逐字节对照确证：仅把 Excel 原生文件的 `cellIs` 改为 `lengthIs` 即拒开）。现改为 Excel 原生形式 `cellIs` + `LEN(ref) <op> <value>`（between/notBetween 输出两条公式），xlsx 与 xlsb 一致。
- **`FormulaEncoder` 多字符运算符丢失**：词法分析中 `>=` / `<=` / `<>` 匹配后未加入 token 即 `continue`，导致这些比较运算符被静默丢弃、公式编码错误（影响所有 BIFF8/BIFF12 公式与条件格式）。现正确产出 `PtgGe`/`PtgLe`/`PtgNe`。
- **`FormulaFtab` 变参函数表补全**：`SEARCH`(82) / `LEFT`(115) / `RIGHT`(116) 补入 `VarArgFuncs`，使文本类条件格式公式（`SEARCH`/`LEFT`/`RIGHT`）编码为 `PtgFuncVar` 而非 `PtgFunc`（后者会导致 Excel 拒开文件）。
- **XLSB 手术式删除的 XTI 递减**：`ModifyExternSheet` 现对「itab > 被删表索引」的 XTI 条目正确递减（此前遗漏），与 Excel 自身行为一致。
- **XLSB 手术式删除的依赖部件路径**：修正被删表 rels 依赖（如 `../tables/tableN.bin`）的路径规范化，避免残留孤儿部件（如 `table11.bin`）。
- **XLSB 手术式删除的记录保真**：修正 `ModifyWorkbookBin` 对 `0x0817`（BrtFileVersion）等记录的解析，确保未修改的记录逐字节保留（此前旧实现会吞并后续记录）。
- **XLSB 超级表降级误报**：`XlsbWriter.ReportDegradations` 此前对含超级表的工作表**无条件**上报 `DegradationCapability.Tables`（"超级表已丢弃"），但三条写出路径（重建 / verbatim / 手术式）实际都保留超级表——重建经 `BuildTableBin` 写出，verbatim/手术式原样透传 `xl/tables/*.bin`。该上报恒为假，已移除（xlsb 不存在丢表路径）。
- **数据模型部件 `xl/model/item.data` 按 STORED 写出**：与 Excel 自身产出约定一致（该部件是内存映射数据库，Excel 存为不压缩）。此前用 `CompressionLevel.Optimal` 使输出体积显著小于 Excel；改为 STORED 后输出大小与 Excel 接近（如 `raw_repro` 删表：1.33MB → 2.31MB）。不影响内容与打开。

### Changed

- **XLSB 条件格式降级上报取消**：全部 18 种 OOXML 规则类型均已支持写出，仅在遇到未知类型枚举值时经 `DegradationCapability.ConditionalFormatting` 上报。
- **`IconSetInfo` 默认阈值改为整数**：未显式给出 `Thresholds` 时，均分阈值取整（如 3 图标 → 0/33/67），与 Excel 原生输出一致（此前为 33.333/66.667）。
- **`IconSetInfo.Reverse`**：新增 `Reverse` 属性（对应 Excel `reverse`），xlsb 经 `BrtBeginIconSet` 标志位写出。

### Notes

- **XLSB 条件格式**：全类型已支持写出。verbatim / 手术式路径下既有条件格式仍原样透传。
- **XLSB 命名区域写出**仍为 C 级降级（本轮仅实现读回）。
- **输出体积与 Excel 的差异（均非信息丢失）**：删表另存时输出与 Excel 的差异来自——① `model/item.data` 已按 STORED 对齐；② 其余部件我们压缩率更高（更小无害）；③ 省略陈旧 `calcChain.bin`（Excel 打开时重建）；④ Excel 重编码 `styles`/`sharedStrings`/`docProps`/`vbaProject` 的副作用（我们保留更完整内容）。经未压缩内容逐目录比对，**零信息遗漏**。

### Tests

- 新增 `XlsbConditionalFormatTests`（cellIs / between 双公式 / iconSet 模板+子记录 / expression 锚点相对引用 / top10 标志 / colorScale 2 色与 3 色 / dataBar 头+CFVO+颜色 / textLength→cellIs+LEN / duplicate / timePeriod / dxf 填充 / dxf 边框边类型 / 未知类型降级 / 全类型不降级，共 14 项）。
- 新增 `SheetVisibilityTests`（可见性/tabColor 四格式往返 + 最后可见表守卫 + verbatim 失效验证）、`XlsbNamedRangeReadTests`（真实样本命名区域读回）、`XlsbTableTests`（表读写往返 + 真实 12 表 + **真实样本打开-保存不误报 Tables 降级**）、`XlsbDataValidationTests`、`XlsbImageTests`（xlsb 浮动图片 media/drawing/BrtDrawing/CT + 打开-保存不翻倍）；`DeleteSheetTests` 新增数据模型工作簿删表**保真成功**（含 `AllowFeatureLossOnSave=true/false`）+ 未被引用表可删。
- 新增 `ConditionalFormatTests.TextLength_WritesCellIsWithLen_NotInvalidLengthIs`（xlsx 回归）。
- 新增 `FormulaTests.Xlsb_ComparisonOperators_RoundTrip`（`>=` `<=` `<>` `>` `<` `=` 编码回归，6 例）。
- 新增 `DeleteSheetTests.Delete_XlsbSurgical_RenumbersPivotAndSlicerCacheRefs`（`0x046D`/`0x0430` 内嵌 rId 随 rels 递减的回归）。
- 新增 `DeleteSheetTests.Delete_XlsbSurgical_RepairsPivotTableCacheId`（透视表 `cacheId` 按 rels 归位的回归）。
- 新增 `DeleteSheetTests.Delete_XlsbSurgical_PreservesLinkedTableNames`（`_xlcn.LinkedTable_*` 名删表前后原样不变的回归）。
- 新增 fixture `excel-authored-namedranges.xlsb`（Excel 生成，含全局 + sheet-local 命名区域）、`excel-authored-table.xlsb`（Excel 生成，含超级表，用于验证降级误报修复）。
- 全量 **710** 测试通过（net8.0）；net48 构建通过。
- Excel COM 验证：四格式可见性、xlsx/xlsm tabColor、xlsb 命名区域/超级表/数据验证均与 Excel 视角一致；数据模型工作簿（`raw_repro.xlsb` 9 表）删 `DayList` → 输出 8 表打开正常，透视表 8/连接 12/VBA/数据模型/标签颜色 4 个全部保留，`workbook.bin` 与 Excel 基准仅差 2 条无语义记录（连接 GUID / activeTab）。

## [2.4.76] - 2026-09-11

### Fixed

- **严重保真问题修复**：打开包含 Excel Table 的 xlsx/xlsm 文件后保存，不再重复生成表部件；原始表 XML（包括计算列公式、原始 dxf/uid 等）按原内容保留。
- **删除工作表的手术式保存**：仅删除工作表且其余工作表未修改时，xlsx/xlsm 保存改为只摘除目标工作表及其关系，其他工作表、透视表、图表、切片器、ActiveX、连接、查询表和自定义 XML 原样保留。
- **工作表关系清理**：删除工作表后同步移除孤儿 worksheet 关系、内容类型声明和过期计算链，避免 Excel 打开时执行文件级修复。
- **XLSB 跨格式安全边界**：xlsx/xlsm 的 XML 高级部件不再透传到 xlsb；包含无法安全重建的高级部件的 XLSB 在非 verbatim 编辑/结构变化时默认阻止保存，避免静默丢失透视表、连接等内容。
- **XLSB 手术式删除**：删除 XLSB 工作表后保存为 xlsb，改为只摘除目标工作表及其引用，其他工作表、透视表、透视缓存、超级表、切片器、连接、查询表、数据模型、ActiveX、图表和 VBA 宏原样保留。同步调整 `workbook.bin` 的 BundleSh、命名区域 itab、外部表引用 XTI 及活动表索引。
- **XLSB 跨格式降级**：xlsb 转换为 xlsm 时保留透视表等二进制部件（超表转为重建），避免默认阻止。
- **能力降级默认放行（行为变更）**：`Workbook.AllowFeatureLossOnSave` 默认值由 `false` 改为 **`true`**。目标格式不支持的能力（VBA 宏、BIFF8 透视表、XLSB 高级部件）不再默认抛异常阻止保存，而是放行并记录到新增的只读清单 `Workbook.SaveDegradations`（非静默，可查询）。设 `AllowFeatureLossOnSave = false` 恢复严格模式（抛异常阻止）。
- **宏转换**：`xlsm`/`xlsb` 转换为 `xlsx`/`xls` 时剥离 VBA 宏工程（含 `vbaProject.bin`、关系、Content-Type 声明、工作簿 codeName），并上报 `DegradationCapability.Macros`。注意：调用了 VBA 自定义函数（UDF）的公式在 Excel 中会显示 `#NAME?`；内置函数与普通引用不受影响。
- **XLS 透视表检测接入**：打开 `.xls` 时接入 `SourceHasPivotTables`（此前从未赋值，守卫失效），使透视表降级能正确上报。
- **XLS 工作表尺寸上限裁剪**：写出 xls（BIFF8）时将列/行裁剪到格式上限（256 列 / 65536 行）。此前超出上限的列宽声明（如声明到第 386 列）或越界单元格会产生越界的 DIMENSIONS/COLINFO 记录，导致 Excel 打开时执行文件级修复。超界数据经 `DegradationCapability.SheetSize` 上报。修复 `xlsx`/`xlsm` → `xls` 的转换。
- **XLSB 手术式删除保留 VBA（回归修复）**：删除 XLSB 工作表后另存为 xlsb 时曾因 `WriteSurgicalXlsb` 中 vbaProject 被双重跳过而丢失 `xl/vbaProject.bin`（Content-Type/关系却仍声明，产生不一致包）。现已让 Parts 循环正常写出该部件，并同步移除过期 `calcChain.bin` 的孤儿关系、修复 `WriteVerbatim` 的 vbaProject 重复条目。
- **BIFF8 日期健壮性**：超出 OLE Automation 日期范围的日期格式数值回退为数字，不再导致整个 xls 文件读取失败。
- **工作表枚举删除**：`foreach (var ws in wb.Worksheets) { ws.Delete(); }` 使用快照枚举，可安全删除；新增 `WorksheetCollection.RemoveAll(Predicate<Worksheet>)` 批量删除入口。

### Tests

- 新增表重复、直接 foreach 删除、`RemoveAll`、跨格式部件过滤和 BIFF8 日期回归测试。
- 使用真实复杂 Excel 样本验证：xlsx/xlsm 删除工作表后 Excel COM 正常打开，表格与透视表保持；XLSB 高级部件结构修改保存被安全拦截。

## [2.4.75] - 2026-09-10

> **补充发布说明**：NuGet/tag 的 2.4.74 基于较早提交打 tag（漏掉了 `Worksheet.Delete()` 与其文档同步），故将该项补入本版本重新发布，使各发布版本的 CHANGELOG 与实际包内容严格对应。

### Added

- **Worksheet.Delete()**：新增工作表级删除方法；`Worksheets.Remove`/`RemoveAt` 保持原语义。删除最后一张表时 `Delete()` 抛 `LiteExcelException`。删除会同步清理引用该表的命名区域，避免写出后 Excel 报 `#REF!`。见 §7.6。

## [2.4.74] - 2026-09-10

### Added

- **统一读取门面扩展到 xls/xlsb**：`Excel.Read<T>`、`Excel.ReadSheet`、`Excel.ReadAsDataTable` 现支持按路径自动路由 xlsx/xlsm/xlsb/xls；从流读取时新增 `ExcelFormat` 显式重载。`Excel.StreamRows` 和 `Excel.EnumerateRows` 同样扩展到 xls/xlsb（path 和显式格式 Stream 重载）。CSV 仍不支持流式读取。
- **XLSB BIFF12 记录级流式读取**：`XlsbRowStreamReader` 逐条读取 BIFF12 记录，遇到 `BrtRowHdr` 时 yield 前一行，支持所有单元格记录类型（Blank/Rk/Error/Bool/Real/St/Isst + Short 变体 + 公式缓存值）。工作表数据不再驻留内存。
- **XLS BIFF8 记录级流式读取**：`XlsRowStreamReader` 预扫描全局段后从目标表 BOF 逐条读取 BIFF8 记录按行 yield，支持 Number/Rk/MulRk/LabelSst/Label/BoolErr/Formula 全部单元格记录。
- **XLSB 自动筛选范围读写**：读取 `BrtBeginAFilter` 提取筛选范围到 `SheetData.Filter.Range`；写出时在 `EndSheetData` 后写出范围记录。复杂筛选条件降级上报。
- **XLSB 高级部件工作表级保护**：含透视表/切片器/时间线的 XLSB，仅当修改了高级部件所在工作表时默认阻止保存；修改无关工作表允许保存。此前为工作簿级判断，会误阻无关工作表。
- **XLSB 批注读写闭环**：根据真实 Excel 样本校准，批注存储在独立 `commentsN.bin` 部件（BIFF12 记录 `0x0274`-`0x027D`）而非 `sheetN.bin` 内；VML 与 XLSX 相同。读取解析 `BrtCommentText` + VML `<x:Row>/<x:Column>` 定位；写出生成完整记录链。此前推测的记录号（`0x003E` 等）已被证伪并修正。
- **XLS 批注读写闭环**：BIFF8 批注记录组（`MSODRAWING` + `OBJ` + `TXO` + `CONTINUE` + `NOTE`）完整读写。读取从 NOTE 记录解析坐标、从 TXO 后续 CONTINUE 解析文本；写出构建最小化 Office Drawing 形状容器。
- **插入/删除行列**：`Worksheet.InsertRows/DeleteRows/InsertColumns/DeleteColumns`（1-based 坐标），同步偏移合并区域、行高列宽、行/列样式、批注、数据验证、自动筛选、图片锚点；删除区域内对象删除、跨越对象收缩。
- **xls/xlsb 公式写回**：新增 `FormulaEncoder`（A1 → RPN），支持常量、单元格引用（含绝对）、区域引用、基础运算符和内置函数；`XlsbWriter` 写 `BrtFmla*` 记录、`XlsWriter` 写 `FORMULA` 记录。不支持的公式降级为缓存值写出。

### Changed

- **XLSB 降级消息更新**：批注降级消息明确说明读取已支持、写出未实现、未修改时 verbatim 保留；自动筛选降级仅对复杂条件触发，范围已正常写出。
- **批注/公式不再触发降级上报**：xls/xlsb 的批注与基础公式现已支持读写，写出时不再上报 `Comments` / `Formulas` 降级。

## [2.4.73]

### Fixed

- **切片器在打开再保存后消失**：`workbook.xml` 的 `<extLst>`（含 `x14:slicerCaches`）与工作表 `<extLst>`（含 `x14:slicerList`）被丢弃，切片器缓存与切片器部件成孤儿，Excel COM 报告 `SlicerCaches` 为空。改为原样捕获并按 OOXML 全序回写，关系编号重排时同步重映射 `r:id`。
- **工作表原始起始行号在打开再保存后丢失**：数据从非首行起的文件（如第 3 行起，配合会话表/透视表的绝对行号引用），保存后整体上移到第 1 行，透视表 `<location>` 与合并区域等绝对引用随之错位。改为记录首个实际行的原始行号，写出时补齐前导空行、合并区域与隐藏行按绝对行号换算。
- **工作表 `sheetId` 被重排为位置序号**：切片器缓存等扩展以 `tabId` 引用工作表 `sheetId`，重排为 1-based 位置序号会导致切片器缓存无法链接到工作表。改为原样回写 `<sheet>` 元素的 `sheetId` 与 `state` 属性。
- **图表在打开再保存后消失**：保留的绘图部件（图表 / 形状）没有配套新增图片时，`sheet{N}.xml` 不写 `<drawing>` 元素，工作表级关系随之悬空，Excel 判定该表无绘图，图表连同图形一并不再显示。改为「本次有新图片 **或** 保留的工作表关系已有绘图关联」时都写出 `<drawing>`；关系编号被重排时同步改写该元素引用的 Id。
- **含透视表的文件打开再保存后 Excel 无法打开**：`workbook.xml` 丢失 `<pivotCaches>`，透视表缓存定义部件虽在包内但无从引用。改为原样回写并按 OOXML `CT_Workbook` 全序置于 `calcPr` 之后；关系编号重排时同步重映射 `r:id`。
- **跨工作簿引用在打开再保存后失效**：`workbook.xml` 丢失 `<externalReferences>`，外部链接部件成孤儿，跨工作簿公式的缓存值与链接一起失效。改为原样回写并置于 `sheets` 之后、`definedNames` 之前，同步重映射 `r:id`。
- **浮动图片在打开再保存后翻倍**：读取回填到 `Worksheet.Images` 的图片被当作新增图片再写一遍，而保真透传的绘图部件已含同一张图，导致每次往返 `xdr:pic` 与 `xl/media` 条目各增一份。读取回填的图片改为标记来源并在写出时跳过。调用方新增的图片仍正常并入既有绘图部件。
- **`Excel.Append` 丢工作簿级引用**：追加路径捕获保留部件时只带工作簿代码名，`<bookViews>` / `<definedNames>` / `<pivotCaches>` / `<externalReferences>` 未随之回写，追加一次即丢命名区域与透视表缓存引用。

### Changed

- **读回的浮动图片属只保真层级**：`Worksheet.Images` 中经打开回填的项仅供查看（尺寸、锚点、字节），修改其属性不会影响保存结果，图片本身随绘图部件原样透传。调用方通过 `AddImage` 新增的图片不受影响。

### Added

- **`Workbook.AllowFeatureLossOnSave`**：源 XLS 文件包含透视表时默认阻止保存（抛 `LiteExcelException`），因为当前模型无法保真写回或转换 BIFF8 透视表，保存会永久删除透视表。设为 `true` 后允许保存，透视表等不可保真能力经降级回调（`OnDegradation`）上报。
- **XLSB 原样保留**：打开含透视表/切片器的 .xlsb 并直接保存（未修改、结构未变）时，`workbook.bin` / `styles.bin` / `sheet{N}.bin` 原样写出而非重建，保留 BIFF12 透视表/切片器宿主记录与扩展样式。此前重建会丢失宿主记录导致 Excel 提示删除 `/xl/pivotTables/pivotTable1.bin`。
- **XLSX/XLSM 扩展样式原样保留**：打开含 `slicerStyles` / `timelineStyles` / `pivotButton` XF 的 .xlsx/.xlsm 并直接保存时，`styles.xml` / `sharedStrings.xml` / `sheet{N}.xml` 原样写出而非重建，保留切片器/透视表视觉样式与稀疏单元格布局。此前重建会丢失 `extLst` 扩展样式并膨胀 `sheet{N}.xml`（空单元格全量写出）。

## [2.4.72]

### Added

- **CSV 编码选项**：`ExcelReadOptions.Encoding` / `ExcelWriteOptions.Encoding` 接受任意 `System.Text.Encoding` 实例。读取时显式指定优先于文件 BOM，未指定时保持 BOM 探测并回退 UTF-8；写出时 BOM 由所给编码的 preamble 决定（默认仍为 UTF-8 带 BOM）。编码实例由调用方提供，**本库不引用任何编码包，零依赖不变**；net48 的 BCL 自带 GBK 等代码页，net8.0 需调用方自行注册 `CodePagesEncodingProvider`。

### Fixed

- **CSV 写出未真正写入 UTF-8 BOM**：`Encoding.GetBytes()` 从不产出 preamble，而 `new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)` 仅设置标志位不影响 `GetBytes`，导致写出的 CSV 实际无 BOM，Windows 版 Excel 打开时按本地代码页解读、中文乱码（文档此前声称「写带 BOM」，与实现不一致）。改为显式写出 `Encoding.GetPreamble()`。**注意：此修复改变了 CSV 写出的字节输出**（默认路径下文件开头新增 3 字节 `EF BB BF`）。

## [2.4.71] - 2026-09-03

### Added

- **`EnumerateRows` 拉取式流式读取**：`Excel.EnumerateRows(path/stream, sheetName?)` 返回 `IEnumerable<IReadOnlyList<Cell>>`，逐行 yield、支持 LINQ 与提前中断、不驻留内存。与 `StreamRows` 的区别：拉取模型、不跳首行、支持 `First()`/`Take(n)`/`Skip(n)`。`sheetName` 为 null 时取第一张表。
- **`Value` 便捷属性**：`Cell.Value` 与 `ExcelRange.Value`（等价于 `SetValue`/`GetValue` 的属性写法，贴近 Excel interop 习惯）。`Cell.Value` 单格读写标量；`ExcelRange.Value` 单格读写标量、多格读返回 `object?[,]`、写标量铺满 / 写二维数组按位填（等价 `Fill`）。`SetValue` / `GetString()` / `GetDouble()` 等原 API 全部保留。
- **流式写入行数上限处理**：`XlsxStreamWriter` / `Excel.CreateWriter` 新增 `RowLimitExceededMode` 参数（默认 `Throw`）与可选 `spillHeader` 表头参数。
  - `Throw`：单表达到 1,048,576 行上限时抛 `RowLimitExceededException`（继承 `LiteExcelException`，含 `RowNumber` / `MaxRows`），避免产出超出 Excel 行数上限的损坏文件。
  - `SpillToNewSheet`：达到上限自动新建工作表（`Sheet1` / `Sheet2` / ...）继续写入，支持超过百万行的数据集一次写出。
  - `Truncate`：达到上限停止写入，文件以满行为止；通过 `writer.Truncated` 属性查询是否截断。
  - `spillHeader`：仅在 `SpillToNewSheet` 下生效，作为每张表（含 Sheet1）的首行表头，调用方只写数据行。
  - 相应把 `workbook.xml` / `workbook.xml.rels` / `[Content_Types].xml` 延迟到 `Close()` 按实际表数写出。

### Docs

- **使用手册重构（中英双语）**：`docs/USAGE.zh-CN.md` 与 `docs/USAGE.en.md` 统一为 24 章 + 附录 A/B 结构（全文目录 + 章内目录），修正 3 处事实错误（图表/透视表在 xlsx / xlsm / xlsb 打开再保存时透传保留、xls/xlsb 公式读写行为拆分、能力矩阵重建为 26 项）。
- **README 三份**：新增「效果预览」截图（LiteExcel 产出在 Excel 中打开的实际效果）、能力矩阵与已知边界小节；去除硬编码版本号与「2.4.x+」引入标注。
- **截图入库**：`docs/screenshots/` 新增 6 张真实截图，并在 README 与使用手册对应章节内嵌。
- **流式写入文档**：§21.4 补充单表行数上限说明、三种 `RowLimitExceededMode` 行为（Throw/SpillToNewSheet/Truncate）与 `spillHeader` 表头示例。

## [2.4.7] - 2026-08-27

### Added

- **对象API补齐**：`Excel.Append`（对象模型单行追加同 `SheetData` 语义）、`Excel.ReadWithProgress`（进度读取）、`Excel.GetSheetNames(Stream)`（仅 xlsx/xlsm）。
- **`Worksheet.AutoColumnWidths()`**：实例方法，低层估算后回填 `ColumnWidths`。
- **`GetSheetNames(path)` 路由修复**：xlsb / xls / csv 不再走 zip 元数据路径（会误读），改经 `Excel.Open` 按格式解析，正确返回表名。
- **Excel 2021 兼容性修复（两处）**：
  - `sheet1.xml` 子元素顺序重排为 OOXML schema 全序（autoFilter → mergeCells → conditionalFormatting → dataValidations → hyperlinks → drawing → legacyDrawing → tableParts），消除「XML 错误 / 丢弃整表」。
  - 批注写回补 VML legacyDrawing（`xl/drawings/vmlDrawing{N}.vml` + `<legacyDrawing r:id="rIdC1"/>`，rels ContentType 用 `.../vmlDrawing`，`[Content_Types].xml` 补 `Default vml`），Excel 2021 下批注可见。
- `FacadeGapsTests`：新增 4 个测试（Append 委托、AutoColumnWidths 回填、ReadWithProgress 回调、GetSheetNames(Stream)）。

## [2.4.6] - 2026-08-25

### Added

- **工作表 / 工作簿保护**：`Worksheet.Protection`（`sheetProtection`，锁编辑 + 可选密码）与 `Workbook.Protection`（`workbookProtection`，锁结构/窗口 + 可选密码）。
  - 密码以 SHA-512 + salt 哈希写出（与 `fileSharing` 同机制，`SetPassword` / `VerifyPassword`），不含明文。
  - xlsx/xlsm 写出/读回；xls/xlsb 走既有降级上报。
  - 验证：7 个单测 + Excel COM 真实打开（`ProtectContents` / `ProtectStructure` 为 True）+ AOT 编译零 IL 警告。
  - 修复：`sheetProtection` 必须位于 `<sheetData>` 之后且含 `sheet="1"`（否则 Excel 拒绝打开或保护不生效）。
- **超级表（Table / ListObject）**：`Worksheet.AddTable(ref, name, style?)` / `RemoveTable` / `Tables` 读回；60 种 Excel 内置条纹样式枚举（`TableStyleStyle`）或任意样式名字符串；列级格式（`Column(name).NumberFormat` / `.Style` → `dataDxfId`）；表头样式（`HeaderStyle` → `headerRowDxfId`）。
  - 写出：`xl/tables/table{N}.xml` + `<tableParts>` + sheet rels + ContentTypes + `styles.xml` `<tableStyles>`。
  - 读回：`dataDxfId` 反查 `styles.xml` dxfs 的数字格式与样式。
  - 未知样式名经 `OnDegradation` 上报（Excel 打开时静默退化为无样式）；xls/xlsb/csv 表降级上报。
  - 保真修复：打开含表 Excel 原稿再保存，`xl/tables/*` 与 `<tableParts>` 透传，不再丢表。
  - 验证：15 个单测 + Excel COM（表名/行数/列数/加行扩展）+ AOT 运行期断言。
- **图标集条件格式（iconSet）**：`ConditionalFormatType.IconSet` + `IconSetInfo`（17 种内置集合枚举 `IconSetStyle`，`Percent`/`ShowValue`/`Thresholds`）。
  - 写出 `<cfRule type="iconSet">` + `<iconSet iconSet="…">` + cfvo 阈值（默认按图标数均分，`Percent=false` 写 `type="num"`）。
  - 读回解析集合名/percent/showValue/阈值；`iconSet` 属性缺失时按默认（3TrafficLights1 族）处理。
  - 验证：4 个单测 + 真实 Excel 样本读回（5 条规则）+ Excel COM（库写出文件打开无修复、规则识别）+ AOT 运行期断言。

### Notes

- 全量 **541 测试**，net48 + net8.0 通过。

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
