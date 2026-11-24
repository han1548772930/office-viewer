# SHEET-22 / SHEET-24 / SHEET-29 原生图表读写

更新：2026-11-24。开发中，未验收。用户要求先实现、后统一测试；本批只阅读源码、修改代码和编写测试草稿，没有执行类型检查、测试、构建、WASM 重建、浏览器或基准。

## 来源与范围

- 固定 OSS：E:\rust_test\univer，v1.6.0，eaf4cfa04fabdbcfc922c692f3450459f8a29546。目录、设置和入口依据另见 [SHEET-22 审计](SHEET-22-chart-type-audit.md)。
- 固定发布类型 .research.local/univer-1.6.0/packages/engine-chart/lib/types/types.d.ts 的 IChartContext.plotVisibleOnly 明确声明缺省保留历史 true；这是公开声明，尚未核验上游运行效果。
- 本项目独立使用传统 SpreadsheetML/DrawingML 图表和 OPC 关系读写，没有使用 Univer Server、Exchange 或商业 SDK。
- 共享目录覆盖 barChart/bar3DChart/lineChart/line3DChart/areaChart/area3DChart/pieChart/pie3DChart/doughnutChart/ofPieChart/scatterChart/radarChart/bubbleChart/stockChart 共 14 种 XML plot。方向、堆积、百分比、3D、饼图拆分和组合系列分别映射；14 是 XML 节点口径，不表示 45 个目录入口已完成。

## 本批落盘

- crates/office-io-common/src/chart_{kind,options,color,xml}.rs：共享分类、强类型配置、颜色和有限预算的 XML/signature。
- crates/office-writer/src/xlsx_chart{,_series,_style,_bindings,_source,_warnings}.rs、xlsx.rs：原生系列/轴/图例/标签/趋势线/误差线和组合图；source manifest 保留原始嵌套设置，不建立第二份权威计算引擎。
- crates/office-parser/src/native_chart.rs、pivot_sidecar.rs、parser.rs：独立 chart part 读取、关系及 pivot 绑定、原生 drawing 几何保留、HSL lightness 和按 XML 顺序执行并逐步 clamp 的 alpha/color transforms。twoCellAnchor.editAs 进入 drawing behavior。
- packages/office-studio/src/parser/workbookCharts.ts、workbookNormalizer.ts：源工作表重映射，native ordinal → field identity，custom-error 源范围、bubble size/OHLC、完整设置和稀疏 anchor/EMU 几何。twoCellAnchor editAs=absolute 仍从单元格计算位置。
- 无法映射的对象登记 metadata.unresolvedOfficeCharts，原始 DTO/XML 保留在 sheet source；writer 拒绝静默遗漏图表。

## 原始名称系列设置

- 固定发布类型 `types.d.ts` 的 `seriesStyleMap` 和 `chart-builder/configuration/cartesian-series.d.ts` 的 `ChartPublicSeriesSelector`、本项目设置入口及字段映射依据见 [SHEET-22 审计 P13](SHEET-22-chart-type-audit.md)。本项目以 `options.namedSeries: { name, style }[]` 持久化；名称按原始源匹配，保留大小写、空白及空字符串，显示别名不改变 selector。精确名称、重名和未活动源的边界是本地规则，尚未核验固定版上游运行效果。
- 原生写出按全系列 → 数值字段 → 原始名称覆盖深合并；组合类型和主/次轴在覆盖与当前可见性筛选后校验。辅助表规划只读取角色，隐藏系列的绘图约束在重新变为活动源时生效。
- 分组 Bubble 使用真实 group 名；普通空 header 使用 `Series ${field + 1}`，未分组 Bubble 保留空 header。HLC/OHLC 的逻辑名为 `最高 / 最低 / 收盘`，字段设置共用 `closeIndex`，原生三个/四个向量仍分别引用其实际源字段。
- 重复配置名称返回 `CHART_NAMED_SERIES_DUPLICATE`；一个 selector 命中多个活动源返回 `CHART_NAMED_SERIES_AMBIGUOUS`。暂未活动的名称设置继续保留，且不套用到显示别名。
- 原始 options 已由 `AuthoredChartOptions { typed, source }` 包装层完整保存。typed surface 未映射的嵌套路径返回 `CHART_NATIVE_UNMAPPED_SETTING`，包括 Series/Point/Mapping 扩展；manifest 的保留不能作为 native Office 渲染等价的证明。
- parser 对相同 plot 的多个主/次轴节点保留原类型，例如两组 `bubbleChart`；只有不同 plot kind 才判组合。独立多源 Bubble 的完整领域导入适配仍是剩余项。

## 真实 helper 绑定

分组气泡、多字段分类、当前隐藏数据、按零值以及缺省气泡大小使用隐藏辅助工作表的真实公式引用；cache 不是唯一数据来源。

- 每个图表单独 helper，工作表名/ID 避让。A1 marker、隐藏状态和经核对的 manifest 共同恢复 ownership；原始 key 与重开后的 chart ID 分开，再次导出重建并复用经证明的 owner。marker 不匹配的用户数据不会被覆盖。
- 当前样本按手动隐藏、像素尺寸/零尺寸、worksheet filter 的 filteredRows、table filter 的存储 filteredRows 和折叠 outline 取舍。字段保留原始 identity；分类或 bubble/stock 主角色隐藏时输出零系列的原生空图。
- 分组用 EXACT 保持大小写；数字和对应 ECMAScript 字符串按同一组。ryu-js 负责数字文本身份；formula 引用组头。
- Y 的按零值与 X、size、自定义 error 的缺失状态有独立 vector key；组外样本保持 #N/A 掩码。缺省 size=1 也有公式引用，Y 编辑后不只依赖过期 literal cache。
- helper 已排除当前不可见数据，native plotVisOnly=0 允许 Office 读取隐藏辅助表；原始 authored plotVisibleOnly 保留。直接绑定继续使用 authored 缺省 true。
- helper 整本预算：1,000,000 cells / 64 MiB 文本；公式至多 8192 bytes，单元格文本至多 32767 UTF-16 units。预算不等于端到端内存或性能验收。

## P12 日期轴、轴标题与悬停设置

- 固定 v1.6.0 `chart-ui/lib/es/index.js` 的轴设置组件分别写 `axes[axis].titleStyle` 和 `titleAlignment`；`engine-chart/lib/types/enum.d.ts` 声明 `start/center/end`。`axis-title.operator.d.ts::applyAxisTitleStyle` 没有公开运行时，端点位置、反向轴与最终文字盒对齐仍未核实；本项目采用沿逻辑轴的位置语义，不将原生 `a:pPr@algn` 当作等价定位。
- `chart_options.rs` 增加 Axis/DepthAxis 的 `titleStyle/titleAlignment`、Mapping 的 `useDateAxis` 和 AxisPointer 四个样式字段。原生标题写入实际字体/颜色；位置字段留在经 signature 核对的作者 manifest，并返回 `CHART_NATIVE_AXIS_TITLE_POSITION`。外部原生标题恢复字体，不从段落文字对齐编造端点位置。
- `excel_date.rs` 和 `xlsx_chart_dates.rs` 使用工作簿日期系统；有效 numeric serial、严格 ISO 日历/时间/时区文字、serial 0 和 fictitious serial 60 有明确规则。数值/ISO 日期走现有公式 helper 的独立日期 vector；Y 的 zero 策略不会把缺失日期变为 epoch 0，分组外 X 样本仍是缺失。日期 helper 按当前文字语法生成真实源引用，语法形状变化需在工作台重新准备，numeric text 使用 Office locale 的 VALUE 转换并记录 `CHART_NATIVE_DATE_BINDINGS`。
- 普通 2D 分类日期写 `dateAx`、`baseTimeUnit=days` 和 `cat/numRef/numCache`；XY 写格式化 `valAx` 与独立 X 公式 vector；parser 从外部 `dateAx` 恢复 `useDateAxis`。3D/pivot 保留 date numeric labels 和作者配置，当前仍用分类轴，elapsed spacing 返回 `CHART_NATIVE_DATE_CATEGORY_SPACING`，宿主原生几何适配仍待实现。子日粒度在 Office 日期轴上的最终间隔、自动刻度和零日期标签未经过独立样本验证。
- 悬停 AxisPointer 是工作台渲染交互配置，写出保留完整作者字段并返回 `CHART_NATIVE_AXIS_POINTER`；没有将保留字段声称为原生 Office 悬停行为。
- `chart-input-preparation.ts` 在聚合/Top N 后同步投影 X/size；合并行坐标冲突或 synthetic Others 返回缺失坐标，Others 即使只归并一行也不带独立原始 record。日期血缘和日期坐标已计入 `chart-compute-memory.ts` admission estimate，此估算不是实测 heap/复制数据。
- 草稿增加 `tests/fixtures/charts/date-serials.json`，IO/TS 共用日期语义夹具；`chart-date-preparation.contract.test.mjs`、writer helper 日期测试、parser 原生日期轴/真实 XLSX 读回测试覆盖相应协议。均未运行。图表浏览器编辑/撤销/保存、Office 重开及真实内存基准仍未验收。

## 待统一验证的草稿

### P14 渐变、纹理、原生差异与缓冲所有权

- 固定 v1.6.0 `engine-chart/lib/types/types.d.ts`、`chart-ui/lib/es/index.js` 的 paint 配置与工作台入口依据见 [SHEET-22 审计](SHEET-22-chart-type-audit.md)。本批原生适配参考 DrawingML 的 `a:gradFill/a:lin/a:gsLst`、`a:pattFill` 和 `c:plotArea/c:layout/c:manualLayout`；没有将原始配置保存等同于 Office 视觉等价。
- `chart_color.rs` 的 solid/linear-gradient 类型与 `chart_options.rs` 的 Series/Point paint 校验保留 2–256 个 stops、有限坐标、0–1 offset/opacity、CSS 颜色和未知字段拒绝。54 种 pattern preset 由共享枚举校验；`fillPattern` 缺省继承，显式 `null` 取消继承，两者分开读写。
- `xlsx_chart_paint.rs` / `xlsx_chart_paint_tests.rs` 写线性渐变角度、按 offset 排序的 stops、逐 stop alpha、纹理及 plot manual layout；作者原顺序留在经 signature 校验的 manifest。automatic gradient 的派生 stops 不写回作者 options；负柱与主/次反向轴按实际 series paint 单独处理。
- `native_chart.rs` / `native_chart_paint_tests.rs` 按 DrawingML namespace 识别渐变与布局，拒绝非法角度、非法 stops 和 namespace 碰撞。path/radial 渐变、无法等价的 fill bounds 和 patterned line stroke 保留源 XML并给出具体差异。当前诊断包括 `CHART_NATIVE_GRADIENT_GEOMETRY`、`CHART_NATIVE_GRADIENT_FILL_BOUNDS`、`CHART_NATIVE_AUTOMATIC_GRADIENT_UNMAPPED`、`CHART_NATIVE_PATTERNED_LINE_STROKE`。
- `workbookCharts.ts` 每次按宿主重新生成 `metadata.officeChartDifferences[chartId] = { index, reasons }`，原因去重；无差异时移除旧记录。非数组/null、空洞/继承项、非字符串或空原因不会生成半条成功记录，而是进入 unresolved；原始 DTO/XML 仍保存在 `sheet.source.charts`。
- `xlsx_chart.rs` 的 prepared numeric/date/category vectors 用 `Cow::Borrowed` 引用同一 ChartInput owner；直接数值读取只分配数值 vector，不再同时生成无用途的字符串数组。系列、统计和 string/number caches 直接追加到同一 chart part；默认 bubble size literal 用迭代器生成。XML escape 按字符串区间追加，避免全量 escaped cache/manifest 副本。所有 cache/series 追加检查 16 MiB 上界及 `try_reserve`，拒绝非有限 numeric cache；没有声称整个 writer 已零拷贝。
- signature 校验临时追加 closing 后解析再 truncate，不复制整份 XML；`chart_xml.rs::parse_with_reserved_nodes` 预留 manifest 的 3 个节点，使 writer 与最终 parser 使用同一 300000 nodes / 128 depth 上界。最终 part 仍有 16 MiB 限额。
- 增加 `xlsx_chart_buffer_tests.rs` 的 owner 指针、zero 策略源不变、Unicode/entities、共享 part 字节上界和 NaN/Inf 草稿；`native_chart_roundtrip_tests.rs` 增加 writer → 独立 parser 的显式/自动渐变、stop alpha、plot layout、pattern/null 继承草稿；`workbook-chart-io.contract.test.mjs` 增加 differences 成功/清理/拒绝草稿。全部未运行。
- 当前仍存在作者 manifest 的 JSON Value/字符串及样式局部字符串、signature XML 树等分配；整条写出峰值、复制字节、真实 Office 重开及图内视觉/几何等价均未验收，不据静态借用路径宣称性能达标。

### P15 瀑布图计算读取与权限边界

- `chart-tasks.ts` / `chart-results.ts` 使用协议 2 的多系列 `waterfallLayout`。密集控制数组校验逐 own index 拒绝空洞/继承项；额外 series identity / total-column 控制字节计入预算，结果 tuples、buffers、counts 与输入身份均验证。累计布局由 Rust `office-compute/src/chart/waterfall.rs` 产生，TS 无第二份生产累计引擎。
- 本地 `FChart.getComputedData/getComputedLayout/getComputedRenderCommands` 共用 Studio 的 `ChartStatisticsService`，多读复用一个任务，返回 clone，读取不写内容历史；这些是本地扩展，不增加固定上游 Facade 数量。
- 异步 continuation 再检查 signal 取消、unit/chart 身份、user、workbook revision、chart signature，并重新判断宿主/源 worksheet 与 source range 查看权限；不改内容 revision 的 ACL 更新同样受约束。非空 waterfall 没有合法 computed cache 时同步 layout/render 返回 `COMPUTE_NOT_READY`。
- `sheet.command.add-chart/update-chart/remove-chart` 在候选 mutator 中按真实图表宿主 `sheetId` 检查编辑权限，避免直接命令绕过 Facade 或只比较 worksheet 内容的保护路径。历史重放及其他批量对象入口继续按统一发布保护核对。
- Root 草稿为 `chart-waterfall-compute-protocol.contract.test.mjs` 和 `chart-computed-facade.contract.test.mjs`：硬编码独立 tuple、transfer 后入站 detached、完整输出拒绝、共享任务/cache clone、取消/版本/ACL 的 resolve→continuation 窗口。图表侧草稿和 Nn 默认行为详见 [SHEET-22 审计](SHEET-22-chart-type-audit.md)。均未执行。

- crates/office-writer/src/xlsx_chart_bindings.rs：grouped zero/raw errors/default size、当前行隐藏原因、多字段分类、主角色隐藏空图与刷新、ownership/reuse 和稀疏大源预算。
- crates/office-writer/src/xlsx_chart_series.rs：名称覆盖后的组合/次轴、隐藏后重新生效的约束、空名、精确分组名称、配置重名和活动源歧义、显示别名及 stock 逻辑字段继承。
- crates/office-writer/src/xlsx_chart_warnings.rs：未知顶层/嵌套控件保留、路径诊断和真实 native manifest 写出。
- crates/office-parser/src/native_chart.rs：饱和色 luminance、逐 transform clamp/alpha 顺序。
- crates/office-parser/src/native_chart_roundtrip_tests.rs：真实 XLSX writer → 独立 parser、组合系列/统计/轴/未知嵌套设置、namedSeries 组合覆盖和双轴 grouped bubble helper。
- packages/office-studio/tests/workbook-chart-io.contract.test.mjs：field remap、EMU/sparse geometry、ownership proof、unresolved reason、两种源方向和可见性缺省。

全部草稿未运行，Office 重开、浏览器编辑/历史和性能闭环未验证。

## 剩余差异与未核实项

- literal/unbound、headerless/disjoint、多个独立 native bubble series、重复 Y field 的独立系列、超出五类本地组合系列的 native imports 仍需适配；保留具体 unresolved reason。
- table filter 的 stored filteredRows 可能陈旧，尚需将本项目唯一 filter evaluator 的当前结果接入 export preparation；不能宣称表筛选源即时导出已经验收。
- aggregation/top-N、非饼图百分比标签仍缺 native formula adapter；独立 tick 样式/长度、circle radar、smooth area、rose/half pie 等有明确诊断。
- marker、hole、bubble scale、3D 尺寸和 stock gap 的 native 限制会 clamp。manifest 保存原值不等于 Office 表现完全相同。
- helper 的新增组、源可见性变化和数字分类的 Office 文本转换存在刷新/格式边界，返回 warning；仍需 Office 样本核验。
- drawing behavior 已解析，领域图表仍以像素 position/size 编辑；轴结构变换和再次 native 写出 behavior 闭环尚未完整。
- per-chart color map override、其他 DrawingML 色彩变换及原生类型专属完整参数还须补齐。
- path/radial 渐变、非标准方向/部分图型的 fill bounds、patterned line stroke 的完整原生适配仍需补齐；配置/XML 保留和差异记录不能替代原生表现等价。
- 无实测性能、复制/峰值内存和独立 Office 验证，保持未验收。
