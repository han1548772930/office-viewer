# SHEET-26 迷你图设置、分组与数据区域

本批先实现，统一验证后置；所有测试均为未运行草稿，SHEET-26 保持“未验收”。

## 固定来源和入口清点

参考 Univer v1.6.0，OSS 提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据是项目缓存 `.research.local/univer-1.6.0/packages/sheets-sparkline` 与 `sheets-sparkline-ui` 的公开 ES 和类型；仅静态研究，没有安装或运行商业 SDK。

| 类别、顺序与位置 | 入口、条件 | 命令与具体来源 |
| --- | --- | --- |
| Sheets 扩展，INSERT / MEDIA，order 5 | 插入迷你图；Sheets 可见，工作簿、工作表和目标区域权限共同决定禁用 | `sheet.operation.open-sparkline-selector`；UI `lib/es/index.js` 的 `Ft`、`Nt`、`Bt` |
| 主画布 DATA，order 5 | 迷你图子菜单，活动格包含迷你图且选择可配置时可见；编辑、删除两个子项 | `sheet.sparkline.context-menu-id` → `sheet.operation.open-sparkline-panel` / `sheet.command.remove-sparkline`；`It`、`Rt`、`Lt`、`zt`、`Bt` |
| 列头 / 行头 LAYOUT | 同一迷你图子菜单和编辑、删除子项，各是一处上下文挂载 | UI `Bt` 的 `COL_HEADER`、`ROW_HEADER` 分支 |
| 移动浮层 | 编辑、删除当前；按命中单元格与全组权限判断 | UI `Kt`、`on`；不是桌面 Ribbon 根项 |
| 数据区域选择器，从上到下 | 源数据区域、放置区域；确认前验证 item/group/none 放置关系 | `Pt`、selector 类型与 core `N` / `Se` |
| 基本设置，从上到下 | 类型（折线、柱形、胜负、饼图）；主题；系列颜色；线宽（折线）；渐变（折线）或圆角（柱形/胜负） | UI `Xt`；core `ISparklineGroupConfig`、`SparklineTypeEnum` |
| 标记，三行两列 | 第 1 行高值/首值，第 2 行低值/末值，第 3 行负值/所有标记；每项有显隐和颜色，颜色随启用显示；`markersPoint`（所有标记）仅折线适用，其余五类也支持柱形/胜负 | UI `Xt`、`Yt`；core `ISparklinePoints` |
| 基本管理 | 删除当前、删除整组；Pie 保留类型和管理，隐藏无意义风格 | UI `Xt` 的 `C` |
| 高级，第 1 组 | 合并组、拆分组；基于选中成员和现有成员数判定能力 | UI `Jt` 的 `H`、`W`；core `Ne` combine/unCombine 分支 |
| 高级，第 2 组 | 显示零轴、轴颜色（零轴启用时显示）、反向 | UI `Jt` 的 `F`、`I`、`L` |
| 高级，第 3 组 | 下限 individual/group/custom 与自定义值；再上限 individual/group/custom 与自定义值，两端独立 | UI `Jt` 的 `ce`、`z`、`B`；core `ISparklineExtremumConfig` |
| 高级，第 4 组 | 包含隐藏源、空白处理、非数字处理；两种缺失数据各有 spacing/zero/connect | UI `Jt` 的 `ne`、`V`；core `updateSparklineExtraNum` |
| 数据修改 API | 单项改源/移动、整组多范围改源/移动、全组设置、删除单项/整组 | core `facade/f-sparkline.d.ts`、`f-sparkline-group.d.ts`、`commands/commands/set-sheet-sparkline.command.d.ts` |

分类手工复核：本插件桌面扩展 Ribbon 根声明 1；主区、列头、行头上下文子菜单声明 3 处、子项挂载 6 处（编辑/删除 2 个唯一动作）；移动浮层 2 个动作，单独记录；没有发现本插件的 Docs、独立上下文 Ribbon Tab 或 More 根声明，不能据此推断完整工作台无此类入口。基本面板 20 个条件控件声明（标记显隐/颜色分别计，渐变/圆角为两个互斥声明），高级 12，选择器 2，共 34 个控件声明。选项值和布局变体不另算控件，不作为完成功能数量。

## 实现契约

以现有 `cell.sparkline` 为唯一持久数据；显式 `groupId` 的成员由工作表单元格派生，不另建分组注册表。设置更新命中组的全部成员；combine 只抽取选中成员、保留原组未选成员；uncombine 将选中成员独立成组；单项与整组改源/移动、当前与整组删除均产生一个原子 patch 计划。

上游放置关系：多个目标区域必须每个为单格；一个目标区域必须单行或单列；多个源区域必须各为单行/单列并一一配对；一个矩形源按目标总数等于源行数时逐行，否则逐列切分。来源 `N` / `Se`，不根据目标方向猜测切分。单项 reset 需要一个源和一个目标格。

本项目按 A1 字符串表达可序列化范围；源允许当前工作簿中的带引号工作表名和绝对引用，目标归属当前工作表。精确列/行边界、重复目标、分组一致性、整组权限和 revision/history/owner 应在提交前拒绝，失败不得部分修改。

已用有界静态提取确认上游核心命令是 `sheet.command.add-sparkline`、`sheet.command.remove-sparkline`、`sheet.command.set-sparkline`，UI 打开命令是 `sheet.operation.open-sparkline-selector` 与 `sheet.operation.open-sparkline-panel`。本项目统一内容命令保留 `sheet.command.set-sparklines`，操作种类在类型化 request 中表达；Root 负责公开 Facade、资源与事件兼容接线。

## 当前实现与验证

本文件在开始编辑前记录清点。本子任务共涉及 13 个文件：本证据、6 个生产文件、4 个新增测试草稿和 2 个独立夹具。Root 的领域类型、正式 API/命令、资源/事件、工作台/scene/大纲接线和旧合同迁移不计入本子任务文件数。

| 文件 | 已落盘行为 |
| --- | --- |
| [sparkline-reference.ts](../../packages/office-studio/src/features/sparkline/sparkline-reference.ts) | A1、多区域、引号与双单引号转义、绝对/相对引用、当前工作簿工作表限定、Excel 行列边界及移动越界拒绝 |
| [sparkline-model.ts](../../packages/office-studio/src/features/sparkline/sparkline-model.ts) | 四图型、五预设及 custom、六标记、独立极值、隐藏/两种缺失数据、轴/线宽/渐变/圆角；纯插入/设置/合并/拆分/改源/移动/删除原子计划与分组一致性 |
| [sparkline-render.ts](../../packages/office-studio/src/features/sparkline/sparkline-render.ts) | 四类真实绘图命令、点颜色与并列极值、gap/connect/zero、跨表读取、隐藏源、大纲数组与新区间字段、独立组极值、源读权限和帧缓存 |
| [sparkline-panel-state.ts](../../packages/office-studio/src/features/sparkline/sparkline-panel-state.ts) | 快照/草稿、只发实际修改的 nested 字段、inactive custom 数值保留、临时空数字与严格发布校验、全组目标/新源权限与 owner context |
| [SparklineSettingsFields.vue](../../packages/office-studio/src/features/sparkline/SparklineSettingsFields.vue) | 按上游顺序的 Ant Form.Item/Select/Checkbox/ColorPicker/InputNumber/Button/Tooltip；基本、高级、三行两列标记、Pie 显隐及真实管理事件 |
| [SparklinePanel.vue](../../packages/office-studio/src/components/spreadsheet/SparklinePanel.vue) | Ant Modal/Form、多范围输入、编辑范围/主成员/数据操作、Facade 提交、false 留面板、拆并后重新读取；stale 保留草稿、generation/owner/IME 与弹窗/焦点保护 |

模型/服务与 UI 分层：单元格 `SparklineSpec` 持有唯一配置，分组信息每次从成员派生；planner 不写工作表，Panel 不直接写模型。设置与数据移动合并为同一请求，由 Root 的统一命令执行权限和版本检查并产生内容历史。`SparklineSettingsInput.extremumMin/extremumMax` 支持仅传 `value` 的 partial 输入，保留未修改 mode；持久化 Spec 的极值 mode 仍必填。混合组的 axis、point 与极值 patch 只包含实际改变字段。

画布输入使用 Ant Input/InputNumber，替换控件没有新增原生输入桥；主题来自统一 ConfigProvider 与 theme token，弹层保持注入主题。Panel 的内容历史按工作簿隔离；提交使用 `host: 'workbench'`、expectedRevision、historyGeneration 和 expectedUserId。迟到 Input、设置、管理、Modal cancel/reload/afterClose 回调带原会话；更换目标/Studio 后拒绝旧回调。成功才发 applied/close；真实 no-op false 保留草稿和面板。

渲染计算当前在 feature 的纯 TypeScript 层执行，输出共享 DrawCommand 供 Canvas2D/WebGPU 使用；帧内 sample/group 缓存避免重复读取，权限在 sample cache 命中之前重查，组极值缓存区分源权限。源隐藏读取覆盖 manual rowHiddenSpans/columnHiddenSpans、旧 outlineHiddenRows/Columns、新 outlineHiddenRowSpans/outlineHiddenColumnSpans 以及 filteredRows。大样本 Worker/Rust/WASM 路径、取消/lease/复制次数和内存指标尚未实现/验收，不能把帧缓存称为性能达标或零拷贝。

### 测试草稿与状态

仅静态读取、编辑和编写草稿。本批未执行测试、类型检查、构建、Rust/WASM/Cargo、dev、服务、浏览器、基准、计数脚本、链接检查或 `git diff --check`。下面数量为手工按 test 声明与注册参数展开复核，表示未运行草稿，不表示通过数量或功能覆盖。

| 未运行草稿 | 声明 / 展开场景 | 对应行为 |
| --- | --- | --- |
| [sparkline-settings-and-groups.contract.test.mjs](../../packages/office-studio/tests/sparkline-settings-and-groups.contract.test.mjs) | 22 / 62 | 源维度配对、square 行优先、引号/绝对/边界、嵌套配置、全组设置、选中子集合拆并、删除、整组增减成员、移动+设置、冲突/失源/no-op/一致性 |
| [sparkline-render-settings.contract.test.mjs](../../packages/office-studio/tests/sparkline-render-settings.contract.test.mjs) | 18 / 30 | 空白/typed text、公式缓存/布尔、数组与 spans 隐藏、跨表/权限缓存、组极值权限变化、六标记/ties/反向、轴/正负柱/胜负/Pie/gradient/radius/finite geometry |
| [sparkline-panel-state.contract.test.mjs](../../packages/office-studio/tests/sparkline-panel-state.contract.test.mjs) | 15 / 22 | partial dirty patch、混合 mode、inactive value、空 custom、owner revision/history/user/host/dispose、完整成员/源/目标权限、一个历史事务与 false |
| [sheet-sparkline-settings.browser.test.mjs](../../packages/office-studio/tests/sheet-sparkline-settings.browser.test.mjs) | 21 / 35 | 真实 classic/grid Ribbon、三处 Canvas 右键、Ant 基本/高级、拆并/改源/删除/undo/native save-reopen、false/stale/reload/旧回调、合成 IME、主题/焦点及 Canvas2D 像素 |
| [sparkline-workbench-fixture.mjs](../../packages/office-studio/tests/fixtures/sparkline-workbench-fixture.mjs)、[sheet-sparkline-workbench.html](../../packages/office-studio/tests/fixtures/sheet-sparkline-workbench.html) | 夹具，不另计测试 | 独立原生模型；真实 Workbench/Pinia/Ant/Canvas 与受控生产 Panel；不修改共用 drawing 夹具，不启动浏览器 |

合同合计 55 个声明 / 114 个展开场景；浏览器 21 / 35；总计 76 / 149，全部未运行。旧 sparkline/API/parser 合同的迁移由 Root 负责，不重复计入。

### 已知差异与剩余缺口

- 上游 `wn` Pie 使用 signed sum；本项目使用数值绝对值分配扇区，避免负角度。这是明确实现差异，Microsoft Office 对应样本行为尚未核实。Pie 采用上游固定八色，不复用无意义的线样式。
- 上游线下 `.20 → .05` 渐变，本项目用 12 个带透明度的 polygon 区间近似；圆形标记与圆角柱也用 polygon。双渲染后端的一致性、裁剪、缩放与真实视觉仍未验证，不能宣称像素对齐。
- 移动浮层编辑/删除虽已清点，本子任务未实现或验证移动浮层；触摸、多选/移动 UI、所有上下文组合及复杂保护仍未验收。
- 本子任务不验收 Root 的正式枚举、Range 公共类型、Facade 返回结构、资源名称/序列化、事件参数和三处工作台接线；正式 API/资源/事件合同仍需统一阶段验证。
- native save-reopen 草稿只覆盖本项目原生快照。独立 XLSX/XLS 导入导出、扩展 OOXML 样式/分组/源引用、复杂历史文件和外部打开样本尚未验证；需要时扩展现有解析/写出链路，不能丢字段后声称完整支持。
- Windows 真实 IME、焦点/弹层竞态、主题、Canvas2D/WebGPU、复杂跨表/隐藏源与真实权限操作均未运行。合成 composition 草稿不替代真实输入法验收。
- 大样本解析/数据读取/组极值/渲染计算的 Worker/WASM 任务及性能/内存/取消指标仍缺；首屏、滚动、编辑、缩放、长任务和传输必须同数据集实测，当前没有达标结论。

SHEET-26 保持“未验收”；本批是实现与草稿落盘，不缩减完整 Pro 目标。
