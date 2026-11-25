# SHEET-04 范围主题九区段编辑与原子应用

日期：2026-11-26。状态：**实现及 45 条新增测试草稿已落盘，未验收**。本批续接[范围主题目录与规则](SHEET-04-range-themes.md)、[局部结构](SHEET-04-range-theme-structure.md)和[移位共享事务](SHEET-02-cell-shift-transactions.md)。按用户安排先实现、后统一测试，没有执行测试、类型检查、构建、dev、浏览器、Cargo/Rust/WASM、基准或 `git diff --check`。

## 固定依据与入口

固定参考为 `E:\rust_test\univer`，Univer v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。先用 `rg --files` 定位真实集中命令文件，再读取模型、服务和命令；没有引用不存在的按完整命令 ID 命名的源码路径。

| 已读取固定来源 | 可观察契约与实现差异 |
| --- | --- |
| `packages/sheets/src/models/range-theme-util.ts` | 九区段、稀疏作者样式、显式 false/null、逐边合并、JSON 与行列/首尾组合。固定 `IRangeThemeStyleItem` 主要是 bg/ol/bd/cl/ht/vt/bl；本地增加既有字体、换行、数字格式和布局属性的编辑。未知 ol 等字段保留，但不据此声称其绘制完整。 |
| `packages/sheets/src/models/range-theme-model.ts`、`services/range-theme.service.ts` | 内置/自定义目录、已使用规则、目录更新及资源生命周期。内置名字优先；本地 UI 编辑内置主题创建新名字的副本。 |
| `packages/sheets/src/commands/commands/register-range-theme.command.ts`、`unregister-range-theme.command.ts` | 注册、取消注册与工作簿内容历史。原四个命令继续可用；本地同内容注册是 no-op，旧配置由同一历史快照恢复。 |
| `packages/sheets/src/commands/commands/add-worksheet-range-theme.command.ts`、`delete-worksheet-range-theme.command.ts` | 设置/移除规则及撤销重做。本地给既有 set DTO 增加可选主题配置，将保存目录与应用规则作为一次事务；没有先注册再应用造成部分成功。 |
| `packages/sheets/src/models/range-themes/default.ts`、`build-in-theme.factory.ts` | 32 个内置主题。原目录、隐藏行条纹、首尾行为和来源继承继续使用已有 resolver。 |

固定 OSS 的 `sheets-ui/src` / `design/src` 未检索到对应九区段配置 UI；该公开能力以 API 为主，本地表单是工作台扩展。未将本地布局声称为固定版商业样式弹窗的逐像素复刻。

本批**常驻根增量 0、插件扩展根增量 0、溢出根增量 0、独立上下文根增量 0**。复用现有开始 → 工具第 3 项 `sheet.operation.open-range-theme`，classic/grid 位置不变；当前 Sheets 常规根仍为 69（30/6/11/17/5）。弹窗内的操作、区段和属性值不计作 Ribbon 根或已完成功能数量。

| 弹窗入口 | 分类与状态 | 实际操作 |
| --- | --- | --- |
| 主题 Select | 原 32 内置项和自定义目录项；编辑草稿时暂时禁用 | 原有主题应用走固定 set 命令。 |
| 编辑副本 / 编辑主题、新建主题 | 只改变临时草稿；组合输入期间禁用 | 内置创建唯一副本名；自定义默认编辑同名配置；新建从空配置开始。 |
| 区段 Select | 1 个选择器、9 个区段，顺序来自领域常量 | whole、first/second column、first/second row、header/last column、header/last row。条纹相位沿用 resolver，名称不代表被隐藏后重新编号的物理首行/列。 |
| 属性表单 | 每区段 16 项共用字段；颜色/字体/对齐/换行/数字格式/旋转/缩进/缩小适应 | 继承删除该作者属性；false、null 和无边框明确覆盖较低区段。wrap 与 shrink 以及角度与竖排保持既有布局互斥语义。 |
| 边框编辑 | 10 个已建模边/斜线位置；线型包含 14 个实际值与继承选项 | 修改当前边，保留其他边和未知扩展；颜色继承清除对应原始颜色别名。 |
| 清空区段、恢复初始配置、放弃编辑 | 仅草稿变更；不增加内容历史 | 清空区段保存 null，恢复复制打开时配置，放弃回到主题选择。 |
| 保存并应用 | 名字校验、内容/历史版本、宿主与全部影响范围权限通过后执行 | 在一次固定 set 命令中发布目录及规则，创建一次内容历史。 |

## 分层、模型与事务

归属按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：九区段编辑 helper 和 workbook planner 属于 Sheets feature；Vue 只管理草稿、表单及命令调用，未增加解析、公式或文件写出内核。

[sheet-range-theme-editor-state.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-editor-state.ts) 提供可独立测试的字段、区段、名字及边框编辑。编辑产生脱离原数据的候选，拒绝无效字体/枚举/配置，不修改正在打开的主题。颜色从继承或 theme/tint 来源改为具体值时清理对应旧来源，防止公共读回恢复已清除颜色；其余样式/主题 metadata 保留。顶部未知主题扩展的归一化也修正了没有 metadata 时对 undefined 的 JSON 克隆。

[sheet-range-theme-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-command.ts) 新增共享 workbook planner 与发布 helper，Studio 和独立工作台共用。set 可携带 `rangeThemeStyleJson` 或原 builder：先归一化作者配置、准备目录、规则和权限范围，再发布。修改同名主题收集所有引用它的工作表/整条规则范围；移除/替换仍包含被移除规则的完整范围。相同目录和相同规则为 no-op；空选区不能经 set 偷先注册，remove/unregister 不能夹带保存配置。目录替换保留已有顺序与 own key，独立视图刷新所有已有目录投影，领域 worksheet 不增加视图目录副本。

[sheet-range-theme-commands.ts](../../packages/office-studio/src/api/sheet-range-theme-commands.ts) 捕获 workbook 模型身份、revision、history generation 与用户；执行前、候选准备后及 commit 更新回调检查当前性。commit 显式传入两个版本，预先完成 view 投影，沿用 sparse style inheritance，关闭无关公式重算。权限失败、序列化重入、历史重置或无效配置不能留下已注册但未应用的目录项。四个固定命令及现有 Facade 方法继续使用同一内容事务。

持久数据仍是 workbook `rangeThemes`、worksheet `rangeThemeRules` 和作者扩展；应用后的模型、dirty、撤销重做及内部 snapshot 重开使用既有路径。草稿、选择区段、边框位置、输入法状态、错误和焦点 generation 不进入模型或内容历史。

## 宿主与 Ant UI

[use-sheet-range-theme.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-range-theme.ts) 的请求增加脱离模型的 `themeStyles`；确认捕获模型身份、内容/历史版本及 activation generation。独立模式还核对原 view 身份和 signature。切工作簿、工作表、Studio、活动状态或用户取消旧弹窗；准备时离开并返回同 ID 的新激活也不能发布旧候选。

Studio 回传只接受执行前 revision + 1、原历史 generation 及实际提交模型；applyView 后继续比对候选 signature，避免 observer 未提交的 view 修改污染模型。独立发布和延迟父回调也核对候选、身份及 generation。迟到关闭回调只向有效的当前宿主回焦点。

新增可选 DTO `expectedHistoryGeneration` 和 `host:{unitId,sheetId}`；host 约束只用于捕获的 UI，显式 Facade/API 后台工作表调用仍可执行。`confirmRangeTheme` / Ant confirm 增加可选第三参数作者草稿；旧仅选择主题及移除调用保持可用。新 workbook planner 是内部共享实现，本批未增加新的公开 Facade 方法。

[SheetRangeThemeDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetRangeThemeDialog.vue) 和新增[SheetRangeThemeStyleEditor.vue](../../packages/office-studio/src/components/spreadsheet/SheetRangeThemeStyleEditor.vue) 使用 `antdv-next` Modal、Form/FormItem、Input、InputNumber、Select、ColorPicker、Button、Alert 与现有 ConfigProvider / theme token。没有新增原生表单、contenteditable 或第二个画布输入桥。编辑已有主题后改成新名字只创建并应用新主题，说明文案同步改变，不误称更新所有旧引用。中英资源位于[range-themes.ts](../../packages/office-studio/src/i18n/sheets/locales/range-themes.ts)。

本批查询了 Context7 的 `antdv-next` 文档并沿用已有 `value-format="rgb"`、受控控件和 ColorPicker clear 用法。实际主题、弹层、键盘与中文输入法仍待统一验证；composition 只阻止未结束组合输入中的提交，不声称真实 IME 已验收。

## 修改文件与草稿

源文件共 **8**：新增 editor-state、StyleEditor；修改共享 command、API command、use-sheet-range-theme、Dialog、style normalizer、locale。测试/fixture 共 **4**：三个新增文件及既有范围主题 fixture 的增量方法。未改 package/lock、Rust、后端服务或文件写出。

| 新草稿文件 | 静态 test 声明 | 验证意图 |
| --- | ---: | --- |
| [sheet-range-theme-editor.contract.test.mjs](../../packages/office-studio/tests/sheet-range-theme-editor.contract.test.mjs) | 26 = 17 纯契约 + 9 API/历史 | 九区段/继承/原始字段/边框、名字/无效输入、完整候选、全工作簿权限、一份历史及内部重开、重入/history reset/no-op。 |
| [sheet-range-theme-editor-host.contract.test.mjs](../../packages/office-studio/tests/sheet-range-theme-editor-host.contract.test.mjs) | 12 = 7 独立 + 5 Studio | 原子发布、请求目录隔离、同 ID 身份替换、用户/激活变化、跨表保护、observer/revision/history 和父回传保护。 |
| [sheet-range-theme-editor.browser.test.mjs](../../packages/office-studio/tests/sheet-range-theme-editor.browser.test.mjs)及[fixture](../../packages/office-studio/tests/fixtures/sheet-range-theme.html) | 7 | 真实工作台九区段/Ant 表单、修改所有引用、保护/版本、名字/合成 composition、grid/暗色/取消回焦、草稿恢复与保存。 |
| 合计 | **45，全部未运行** | 循环内参数不另计声明，也不计作功能完成数。 |

旧基础 46、局部结构 32、共享移位 56 是不同批次，本批不重复计算。只进行了源文件/文档读取、`rg` 检索、手工静态审阅、文本声明计数与文档写入；没有解析或执行本项目测试/组件/算法。

## 剩余缺口

- 45 个新增及相关旧草稿、前端类型/构建、真实明暗主题、焦点恢复、权限变化、中文 IME 和 Canvas2D/WebGPU 均未运行/验收。浏览器草稿强制 Canvas2D，合成 composition 不代替真实系统输入法。
- 原样式来源 `ol`、完整 theme/tint、复杂 shaping 及全部混合态仍需补齐或验证；本地新增字号、缩进等扩展在内部模型保存，不冒称固定 `IRangeThemeStyleItem` 的所有公共 builder/writer 已覆盖这些扩展。
- 全套公开 snapshot 输出、XLSX 表/范围主题映射及本项目 parser/writer 的 Office 保存重开未验收。内部 JSON 不等同实际 Office 文件闭环。
- 大量规则索引、样式投影和布局的 Worker/Rust/WASM、取消、复制/内存计量及同数据集基准仍待实施/测量。本批只是有界交互和既有主线程 planner，不声称优化或达标；[WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)不变。
- 复杂复制/格式刷/填充/轴与局部移位组合、协作权限和服务持久化继续保留完整目标，见[后端需求](../BACKEND_REQUIREMENTS.md)。

SHEET-04、Sheets/Docs 与所有 Pro 的完整范围保持不变，本批状态为“未验收”。
