# SHEET-13 验证列表缓存、Rust 查询与 Ant 虚拟弹层

续接日期：**2026-11-27**，文件名沿用前批日期。固定 Univer **v1.6.0 / eaf4cfa04fabdbcfc922c692f3450459f8a29546**。功能 ID 为 SHEET-13，计算/内存属于 COMMON-11；承接[标签与尺寸](SHEET-13-validation-tag-dimensions-2026-11-28.md)、[日期与列表输入](SHEET-13-validation-cell-input-2026-11-28.md)和[格式化来源](SHEET-13-list-source-index-numfmt-2026-11-28.md)。已有真实生产接线和未运行草稿，状态仍为**开发中、未验收**。

## 1. 固定来源与入口

| 来源 | 已核实的契约 | 本地实现及差异 |
| --- | --- | --- |
| [ListDropDown.tsx][list-ui] 的 SelectList/ListDropDown | 搜索、候选颜色、选择与 Edit；以 toLowerCase 过滤并强调首个匹配片段；普通列表逐项 map | Ant Form/Input、Checkbox/Tag 和 Listy 虚拟列表；Rust 查询完整候选后仅可见行生成控件/回调；匹配片段后续已接，见[续接](#match-highlights) |
| [dv-list-cache.service.ts][list-cache] 的 getOrCompute/markRuleDirty | unit→sheet→rule 候选/颜色缓存；规则更新、移除与公式结果事件失效 | 仅内联列表持续缓存并隔离用户/地域；动态引用继续唯一公式读取链，没有完整依赖 epoch |
| [dropdown-manager.service.ts][dropdown-manager] 的 handleSave | 候选 label/value 关系及真实提交 | 选择继续既有 typed 命令、候选顺序和内容历史；CSV、多选保持打开等差异沿前批记录 |
| [shape/layout.ts][tag-layout] 的 layoutDropdowns | 当前单元格宽度下 item 换行，自动宽取最大布局行宽 | 本批未改变标签自动尺寸算法，不把全部标签宽度相加 |
| 本地 [Listy 类型](../../node_modules/antdv-next/dist/listy/interface.d.ts)/[实现](../../node_modules/antdv-next/dist/listy/index.js) | antdv-next 1.8.2 的 items/itemRender/rowKey/virtual/scrollTo 与动态测量 | 本地已安装源码为准；Context7 当前文档用于交叉参照，该版本没有独立索引文档 |

本批没有顶部工作区修改。常驻、扩展、More 溢出、上下文根的**入口增量分别为 0**；没有商业 SDK。缓存、索引、控件与草稿数均不作为功能完成数。

## 2. 缓存与来源所有权

[validation-list-cache.ts](../../packages/office-studio/src/features/sheet-validation/validation-list-cache.ts) 提供 SheetValidationListReadCache。默认最多 **64 条目 / 8MiB 估算保留量 / 每规则 65,536 候选**，按 unit/sheet/user/region 隔离；签名包含 rule ID/type/formula1/formula2。内联候选可跨无关 revision 复用，颜色/文字/模式变化使用新签名。超出保留预算跳过缓存，**不截断候选**。候选对象和数组就地冻结，没有再复制完整候选对象集合。

LRU 估算包含签名、文字/颜色及对象/索引开销；它限制保留量，不是已测浏览器 heap 峰值。缓存不持有 workbook、源单元格、规则或权限对象。签名 getter、提取期间 clear/切换 scope 后，旧结果不能写入新 scope。

[validation-cell-options.ts](../../packages/office-studio/src/features/sheet-validation/validation-cell-options.ts)、[反馈宿主](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-feedback.ts)和[单格宿主](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts)已实际消费缓存；读取先检查目标、合并与来源 View，owner/revision/user/region 更换后丢弃旧结果。引用和动态来源每次通过共享公式读取链；完整数组/spill、依赖版本及 volatile epoch 仍未实现。

Popup current/queryContext 检查目标 View 与最终 generation。旧 owner 的 options/selected 返回空数组；View 撤销结束原 lease，Edit 拒绝仍保留合法草稿。入口 token 包含区域与 sequence；关闭后宿主、投影视图替换或区域切换同样更新 generation，切回原值不会授权旧入口/延迟回焦。inputSession 仍位于原 watcher index 6，成功结束输入的例外保持。

## 3. 计算域与协议 2

分层为 Vue 控件→[查询宿主](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-list-query.ts)→[查询服务](../../packages/office-studio/src/features/sheet-validation/validation-list-query.ts)→既有 OfficeComputeClient/Worker→Rust。查询只产生派生索引，不修改权威单元格、规则或内容历史。

| 接口或数据 | 契约 |
| --- | --- |
| kind | validationListQuery；所有端仍使用协议 **2** |
| 新必需 WASM 导出 | create_validation_list_query(offsets, textBytes, queryBytes, control)；缺失接口/旧产物明确失败，不回退旧 API 或替代 matcher |
| 输入 | Uint32Array offsets 为偶数**字节偏移**，长度 rows+1；textBytes/queryBytes 为保留孤立 surrogate 的精确 UTF-16LE |
| 参数/窗口 | rowCount/textByteLength/queryByteLength/start/count；正窗口 count、字节/rows/安全整数上界在 TS/Rust 验证 |
| 大小写 | JS 无参数 toLocaleLowerCase 后编码；Rust 不重做大小写或修复 surrogate |
| 算法 | offset 验证、KMP prefix/搜索分块；不跨行匹配；total 与指定命中窗口 indices 分离 |
| 取消 | 每 16,384 工作单位使用既有 compute_yield 宏任务让出并检查取消；同步创建/边界复制沿客户端 terminate/recreate 策略 |
| 输出 | 独立 Uint32Array/total；验证完整窗口、所有索引边界与严格递增后原子发布，无部分成功 |

TS 域代码是 [protocol](../../packages/office-studio/src/compute/validation-list-query-protocol.ts)、[tasks](../../packages/office-studio/src/compute/validation-list-query-tasks.ts)、[results](../../packages/office-studio/src/compute/validation-list-query-results.ts)。Rust 算法是 [validation_list_query.rs](../../crates/office-compute/src/validation_list_query.rs)，浏览器适配是 [validation_list_query_wasm.rs](../../crates/office-compute/src/validation_list_query_wasm.rs)。compute protocol/tasks/guards/client/worker/loader/index、TaskKind 和 Rust lib 已接线。

两端相同预算公式（字节）：

```text
input  = (rows + 1) * 4 + textBytes + queryBytes
output = min(max(0, rows - start), count) * 4
peak   = input * 2 + queryBytes * 2 + output * 2 + 8MiB
```

peak 包含 Worker JS/linear-memory 输入、KMP prefix、保留与独立窗口输出；源模型、候选对象、宿主 retained source 和 GPU 内存另计。准备阶段另有候选/折叠文本的保守 admission，不把该公式当作全应用峰值或已测达标。

WASM 适配复用 office-wasm-bridge 的 serde-wasm-bindgen 控制 DTO、js-sys TypedArray、wasm-bindgen-futures/web-sys 宏任务让出与 buffer 边界。普通 JS 输入必须复制到 linear memory；输出同步创建独立有效区间 copy 后释放 Rust owner，再 transferable。无跨 await lease、整块 WASM memory transfer 或额外 Rust Vec clone。元数据保留 controlDtoConversions、输入/输出次数和字节；空 buffer 的有效复制与 transport/准备次数分别计量，不称全链路零拷贝。

## 4. 查询宿主与实际虚拟 UI

查询服务保留一份 packed source，按 **1024 labels / 32,768 UTF-16 units** 分块准备，映射每 **2048 indices** 让出并检查 owner/取消。默认 Worker 阈值暂为 **1024 options / 65,536 text units**；小查询保留原 case-fold/substring/order 契约，阈值未经基准。compute client 使用 copy ownership，准备复制、Worker transfer 和 WASM 必要复制分开计量。

宿主保留一个当前 source、一个完成结果缓存和惰性有界 client；以 immutable options+query 复用结果，无关 revision 保持 matches 身份，避免多选后拆掉虚拟列表。单个 shallowRef view 原子发布结果/忙态/错误/进度。state/capability getter、成功/错误发布、finally 重查 token/pending，旧任务重入不能清空后继 source/结果/忙态。关闭、View 丢失、source/owner 替换与卸载取消并释放派生数据/自有 client。错误保留 errorCode，重试使用当前合法上下文。

[SheetValidationCellControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetValidationCellControl.vue)已使用 Ant **Listy，显式 virtual=true**，动态行高与主题尺寸。仅可见 itemRender 创建 Checkbox/Tag/事件闭包，没有全量 UI 回调数组。搜索草稿/稳定 query 分离，compositionEnd 后更新查询；Ant Spin/Progress/Alert/Retry 展示真实准备/查询进度及错误。Form/Input/Checkbox/Tag/Popover/Button 均为 antdv-next 控件，沿统一 ConfigProvider/token，无新增原生可替换 UI。

[导航控制器](../../packages/office-studio/src/components/spreadsheet/sheet-validation-list-navigation.ts)使用完整结果索引处理 ArrowUp/Down/Home/End；Ant Checkbox 使用 roving tabIndex。scrollTo 后等待有限动态测量帧，各 scroll/await/focus 检查 owner/query/revision/IME。卸载释放 RAF，nextTick 后检查组件活性，过期请求不能聚焦后继行。

选择继续 **sheet.command.commit-validation-list-input** 与原候选排序/类型/验证、文档隔离内容历史、内部保存；查询/搜索/导航不新增内容历史。Office 保真及真实 UI/IME 尚未验收。

## 5. 未运行草稿与静态证据

| 草稿 | 声明 / 计划 | 目标场景 |
| --- | ---: | --- |
| [列表缓存](../../packages/office-studio/tests/sheet-validation-list-cache.contract.test.mjs) | 9 / 20 | 内联身份/签名、LRU/预算、scope/权限与 clear/提取重入；引用不伪造持续缓存 |
| [查询协议/服务](../../packages/office-studio/tests/sheet-validation-list-query.contract.test.mjs) | 14 / 49 | UTF-16/locale、字节/窗口、独立 buffer、预算、实际 client ownership/计量、取消与所有索引边界/顺序 |
| [查询宿主/导航](../../packages/office-studio/tests/sheet-validation-list-query-host.contract.test.mjs) | 17 / 39 | source 重用、权限/卸载/惰性 client、成功/失败/finally 重入、虚拟索引/动态测量 |
| [Popup 宿主](../../packages/office-studio/tests/sheet-validation-cell-host.contract.test.mjs) | 20 / 31 | 原输入；内联身份/revision、View/region、关闭后投影视图替换再恢复的旧入口/延迟焦点；Studio/standalone×assignment/receiveView |
| [真实浏览器](../../packages/office-studio/tests/sheet-validation-cell.browser.test.mjs) | 23 / 45 | 原日期/列表；classic/grid 明暗、真实 Rust 长列表 End/Home/颜色/历史/内部重开、IME 稳定搜索和 Worker progress 切表取消 |
| [Rust 纯内核](../../crates/office-compute/src/validation_list_query/tests.rs) | 6 声明，运行 0 | 独立小窗口 oracle、KMP/步进、surrogate/不跨行、offset/形状/预算和消费状态 |
| [真实 WASM 绑定](../../crates/office-compute/tests/wasm_validation_list_query.rs) | 3 声明，运行 0 | Worker DTO、linear-memory 别名拒绝、复制/变更/grow/release、宏任务让出/重新创建 |

JS oracle/异步 adapter 仅在草稿中控制比较/时机，不是生产第二权威。[浏览器 fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html)只观测真实 OfficeComputeClient；缺少新查询 WASM 产物时长列表草稿失败，不 skip。既有 wasm-protocol-v2/slicer-member mocks 已补新必需接口，并独立覆盖遗漏查询绑定。

限定代码静态检查已通过：累计 **86 文件、86 scripts、7 Vue templates、7 styles、715 本地引用、28 Node 语法解析**；**27 JS 草稿、366 声明 / 663 计划，运行 0**。四个新增 Rust 文件和共享 lib/control 单独登记，累计 **10 Rust 文件**只读核对模块/接口路径与声明，不执行 Rust 编译或测试。

文档/来源/locale 静态检查已通过：**11 文档、278 本地链接（含 51 固定来源链接）、2 anchors、12 唯一固定 GitHub 源路径；19 locales×47 keys**，包含新增的 19×5 query keys。全部 **107 修改文件**的冲突标记、尾随空白和 replacement character 检查通过；限定 git diff --check 通过。固定参考 HEAD 正确且工作树无修改；这些结果只核对来源、语法、路径、文本与计数，不执行产品验证。

旧 76 文件/664 引用/319 声明/541 计划/19×42 keys 属于尺寸批历史快照，不覆盖本次缓存/查询或旧文件追加。当前没有运行测试、类型检查、构建、Rust/WASM 检查、dev、浏览器、Office 往返或性能基准，全部行为保持未验收。

## 6. 剩余实施与统一验收

1. 完整动态依赖/数组/spill、volatile epoch 与大来源求值仍属唯一共享公式引擎。当前查询在 Rust/Worker 分块，来源公式求值仍同步；查询迁移不表示来源引擎迁移。
2. 单个长字符串 toLocaleLowerCase 仍整串同步；编码/查询分块不能消除它。固定 Univer 使用 toLowerCase，本地保留原无参数 toLocaleLowerCase，默认 locale 差异仍需核实。匹配片段已在下述续接实现，运行验收仍后置。
3. 统一验证先生成匹配协议 2/新必需接口 WASM，再执行合同、Worker 绑定、真实 Listy/IME、两后端/主题/焦点/权限、编辑→历史→内部保存重开，以及 XLSX 导出/独立重开。
4. 用同一普通/大样本比较原查询、TS 小路径和 Worker/Rust 端到端路径，记录设备/浏览器/后端、规模/冷热、准备/查询/滚动/编辑帧、长任务、每类复制/transfer、JS/WASM/GPU 内存。未实测，不声称优化、全链路零拷贝或大样本达标。
5. Studio 的 refreshWorkbookAutoRowHeights currentUserId、detached sheetViewCellAt tables、Canvas 空白格/继承范围主题 table filter context 仍由主任务协调；本批未覆盖其在途文件，沿尺寸证据保留缺口。

<a id="match-highlights"></a>

## 7. 搜索匹配片段续接

沿固定 SelectList 的首个匹配片段契约，[validation-list-highlight.ts](../../packages/office-studio/src/features/sheet-validation/validation-list-highlight.ts)将与查询相同的无参数 locale 折叠索引映射回原文字。在 Intl.Segmenter 可用时强调完整 grapheme，处理 İ 扩展、土耳其组合点收缩、立陶宛附加点、组合字符和 emoji ZWJ，保留原始 label/value。低层搜索仍可匹配 surrogate 单元，但 UI 不把成对 surrogate/完整 grapheme 拆为强调片段。没有 Segmenter 时沿已有文本布局的 code-point 回退，完整 grapheme 能力仍需补齐。

仅可见 itemRender 创建强调 span，颜色仍由原 Ant Tag 管理，字重使用统一主题 fontWeightStrong；没有新交互控件、候选/命令/历史或协议变化。助手只遍历命中前缀，不建立全量索引映射数组；整串折叠仍同步，未测性能，不称优化达标。query 草稿在 IME 中变化不改稳定搜索的强调片段，compositionEnd 后同步新结果。

新增[Unicode 强调合同](../../packages/office-studio/tests/sheet-validation-list-highlight.contract.test.mjs)包含固定边界与原文字重组、空/无匹配、土耳其/立陶宛及无 Segmenter 场景；既有真实浏览器草稿追加 IME 前后文字与主题字重检查。草稿未运行，上述 86 文件/366 声明/663 计划等为查询批静态快照，不覆盖此续接。

[list-ui]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/views/dropdown/ListDropDown.tsx
[list-cache]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation/src/services/dv-list-cache.service.ts
[dropdown-manager]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts
[tag-layout]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-data-validation-ui/src/views/widgets/shape/layout.ts
