# SHEET-21：绘图组合、仿射变换与混合层级

记录日期：2026-11-26。状态：**实现及契约草稿已落盘，未验收**。

本批属于 SHEET-21「形状及绘图（含 Pro 目标）」中的组合、取消组合、混合层级、选区变换和删除子项；不将这些子项计为整个 SHEET-21 已完成。完整形状、连接器、文本及 Pro 目录仍在项目目标内。全工作台入口清点继续以 [顶部完整清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)、[Sheets 固定版源码清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md) 和 [功能对齐表](../UNIVER_FEATURE_PARITY.md) 为准。

原子任务新增 12 个实现文件、4 个契约草稿文件（**45 个 `test` 声明**）、1 个模型夹具和本文，共 18 个文件。中央模型、共享仿射数学、Canvas、ExcelWorkbench、Facade、包导出及工作表复制由主任务负责，不纳入这 18 个文件。后续主任务的共享选区/八项对齐/作者 affine/最终轴生命周期接线，以及固定 move-drawing/实际 skeleton 边界/1000ms 连续历史，初批 53 条、当前共 82 条新草稿及实际工作台夹具见[绘图工作台增量证据](SHEET-21-drawing-workbench-affine.md)，均未运行。

本阶段遵从已有执行约束，仅进行了源码阅读、搜索、编辑和人工静态审阅。**没有运行测试、类型检查、构建、Vite/dev、浏览器、Rust/Cargo/WASM、服务、性能基准或 `git diff --check`**。测试声明、入口登记和 JSON 草稿均不能作为浏览器或 Office 文件验收证据。

## 固定版来源与入口清点

只读对照根目录：`E:\rust_test\univer`；固定版本 **v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**。

| 依据 | 固定源码路径与关键位置 | 本批追踪结论 |
| --- | --- | --- |
| 上下文标签、显隐、禁用 | `packages/drawing-ui/src/menu/group.menu.ts:30`、`arrange.menu.ts:27`、`locale/zh-CN.ts:112` | 组合父菜单根据 focus 和可组类型隐藏；组合不足两项禁用；取消组合无组选中禁用；排列无 focus 隐藏 |
| 分组与顺序 | `packages/drawing-ui/src/menu/schema.ts:63` | `DRAWING / OTHERS` 内组合父项 order 1、排列父项 order 2；子项顺序见下表 |
| operation 到业务命令 | `packages/drawing-ui/src/commands/operations/drawing-group.operation.ts:37`、`drawing-arrange.operation.ts:29` | 读取 focus、构造 group 更新或 order 通知，由 Sheets 插件接内容命令 |
| 可组类型与几何 | `packages/drawing/src/utils/drawing-group.ts`、`packages/core/src/types/interfaces/i-drawing.ts`、`packages/engine-render/src/basics/transform.ts` | 实际 helper 覆盖 image、shape、chart、slicer、timeline、group 六类；不能依据 operation 中旧注释的四类描述缩减目标 |
| 内容事务和撤销重做 | `packages/sheets-drawing-ui/src/commands/commands/group-sheet-drawing.command.ts:45`、`packages/sheets-drawing-ui/src/commands/commands/ungroup-sheet-drawing.command.ts:42`；`packages/sheets-drawing/src/commands/commands/set-drawing-arrange.command.ts:35`、`packages/sheets-drawing/src/commands/mutations/set-drawing-apply.mutation.ts` | 分组/取消分组接受 `parent / children` 更新；排列接受 `drawingIds` 加 action 或绝对 zOrder；一次内容命令对应一次历史 |
| 组模型、顺序和逆操作 | `packages/drawing/src/services/drawing-manager-impl.service.ts:609`、`:709`、`:739`、`:869`、`:968` | group 数据与 order 分离、按最高选中层登记组；分组和取消分组均可逆；绝对顺序 floor/clamp |
| 绝对层级边界 | `packages/core/src/common/drawing-order.ts:21` | 保留 NaN → 0、正 Infinity → 最末项、负 Infinity → 0 的公开归一化行为 |
| 保存/载入 | `packages/sheets-drawing/src/controllers/sheet-drawing.controller.ts:36`、`:92`、`:123` | 上游通过 `SHEET_DRAWING_PLUGIN` 保存 data + order；本批使用本项目领域状态，不复用 JSON1/上游私有服务 |

以下计数仅覆盖本批分组/排列子项。展示位置分类为：常驻 0、顶部扩展 0、溢出 0、**上下文父入口 2、上下文叶动作 6**。这些入口属于 drawing-ui 插件，但不再重复计入顶部扩展。画布浮条复用同一命令时应登记为另一个展示位置，不能再计为新的产品能力。

| 位置 / 顺序 | 标签 | 上游 operation ID | 本地内容 command ID | 状态契约 |
| --- | --- | --- | --- | --- |
| DRAWING / OTHERS / 1 | 组合 | `contextMenu.drawing-group` | 父入口 | focus 为空或有不可组类型时隐藏 |
| 组合 / 0 | 组合 | `drawing.operation.set-drawing-group` | `sheet.command.group-sheet-image` | 至少两个同父可编辑、可用、可组对象；拒绝祖先/后代重叠 |
| 组合 / 1 | 取消组合 | `drawing.operation.cancel-drawing-group` | `sheet.command.ungroup-sheet-image` | 至少一个组选中；保留混合选区中的非组对象 |
| DRAWING / OTHERS / 2 | 排列 | `contextMenu.drawing-arrange` | 父入口 | focus 为空时隐藏 |
| 排列 / 0 | 置于顶层 | `drawing.operation.set-drawing-arrange-front` | `sheet.command.set-drawing-arrange` | front；选中成员的相对顺序保持 |
| 排列 / 1 | 上移一层 | `drawing.operation.set-drawing-arrange-forward` | 同上 | forward；连续选中块移动一层 |
| 排列 / 2 | 置于底层 | `drawing.operation.set-drawing-arrange-back` | 同上 | back；选中成员的相对顺序保持 |
| 排列 / 3 | 下移一层 | `drawing.operation.set-drawing-arrange-backward` | 同上 | backward；边界无内容写入 |

本地另有 `sheet.command.transform-drawing-group` 和 `sheet.command.remove-sheet-drawings`，分别承担保留完整矩阵的组/多选变换和混合 Delete，**是本项目新增 ID**。固定上游 `SetSheetDrawingCommand` 的 ID 为 **`sheet.command.set-sheet-image`**，已属于图片宿主；没有虚构或重复注册 `sheet.command.set-sheet-drawing`。

本地菜单目录保存上游 operation ID，实际菜单、键盘和 API 进入共享的内容 command。本文没有将全部上游 operation 注册、商业形状菜单或 Pro 全目录声明为已兼容。

## 领域状态、几何与真实宿主

领域和算法文件：

- [sheet-drawings.ts](../../packages/office-studio/src/domain/sheet-drawings.ts)：typed ref、组、每工作表 drawingState 和 shape DTO。
- [drawing-group-types.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-types.ts)、[drawing-group-geometry.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-geometry.ts)：能力、请求、事务、完整仿射几何和组选框计算。
- [drawing-group-model.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-model.ts)：单父树、身份校验、paint/hit 顺序及组选中目标。
- [drawing-group-planner.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-planner.ts)、[drawing-group-transaction.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-transaction.ts)：组合、取消组合、排列、组/混合选区变换与前后状态 guard。
- [drawing-group-reconcile.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-reconcile.ts)：显式删除、空组清理和完整 subtree copy 映射。
- [drawing-group-commands.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-commands.ts)、[index.ts](../../packages/office-studio/src/features/drawing-groups/index.ts)：固定菜单目录、命令 ID 和模块导出。

矩阵 `{a,b,c,d,e,f}` 将对象局部 CSS 像素映射到**不含表头的工作表世界坐标**。复用主任务的 `domain/drawing-affine.ts`，保留旋转、反射、错切和非等比缩放，不通过角度/宽高的有损分解覆盖权威矩阵。组只保存世界选框和成员顺序，叶几何由真实 image/chart/shape 宿主提供；不缓存另一份组内子坐标。

root/member 均从底到顶，每个引用恰好一个父容器。typed kind 区分同 ID 的 image/chart；未知 kind、缺 owner、重复父、循环、过期状态/几何/revision 都拒绝。创建或复制组时额外拒绝与叶 ID 冲突。组合至少两成员，取消组合在原容器位置展开；组变换后取消组合不再次变换叶。显式删除可保留单成员组，并递归移除空组。

非连续成员 `[a,c]` 在 root `[a,b,c,d]` 中组合后得到 `[b,g,d]`，组成员保持 `[a,c]`。混合多选也可直接变换，无需先创建组；嵌套组每片叶最多一个 geometry patch。独立移动子对象时按子先父顺序重新拟合未直接变换的祖先选框。

[sheet-drawing-host.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts) 已使用本项目实际 `WorkbookModel`：

| Owner | 读取与发布规则 |
| --- | --- |
| image | 用 `sheetImageDrawingGeometry` 和真实稀疏轴投影；用 `sheetImageWithDrawingGeometry` 发布；保留资源、源内 flip、替代文本和归一化 crop，更新兼容锚点 |
| chart | legacy position 减表头后读取；发布完整 drawingGeometry，兼容 position 写 AABB + 表头；保留 source/options/local size；native 正向变换 revision + 1，反向恢复原 guard revision并拒绝后续新内容 |
| shape | 读取/写回 `SheetDrawingShapeDTO.geometry`；source/style/text/metadata 仍属于 shape 作者 |
| slicer/timeline | 后续已接完整 author affine、真实 final axes、native command/history/事件、Ant 控件和混合组；`groupable/transformable` 已开放，lockedPosition 仍阻止变换，折叠 Both 保留作者但临时不可用。来源与 98/182 未运行草稿见[切片器增量](SHEET-25-slicer-timeline-affine.md) |

隐藏行列可能把 legacy Both 图片或无绝对布局的控制压为零尺寸。宿主保留原资源/锚点，提供仅供 state 的 1px 可逆 placeholder 并标记 `available:false`；调用方的 available callback 不能覆盖原生不可用状态。**绘制仍必须读取原生对象，不能画此 placeholder**。恢复显示后重新投影原锚点；选中不可用叶时组合/变换明确拒绝，排列或删除可以保留正确引用生命周期。

宿主先在 detached candidate 上创建/验证新 scene，再统一发布各 owner 和 state。权限范围使用完整 affine AABB，覆盖受影响组、后代与位移前后 footprint，加一次表头后进入本项目实际轴范围计算。Delete 展开组并只删除对应 owners，保留 Table/Pivot 源；主图表不可经组删除绕过。图表媒体清理由既有 `pruneChartImageResources` 在同一 candidate 中完成，其他 owner 仍引用的 mask 保留，最后 owner 删除时清注册表。

## 命令、历史与接线边界

[sheet-drawing-commands.ts](../../packages/office-studio/src/api/sheet-drawing-commands.ts) 接受 typed selection、公开 drawing refs、groupId、frame/delta、action/zOrder、内容/历史版本、UI host、expectedSelection 和取消信号。公开 numeric kind 为 image 0、shape 1、chart 2、group 6、slicer 10、timeline 11；仅 ID 有歧义时拒绝。

公开 `drawings` 从首项推导未指定宿主；`parent / children` 组合批次从首 parent 或首 child 推导，并检查所有已声明 unit/sheet 一致。组合批次的初始选区来自实际 children，成功后选中全部新建外层组；不足两成员的 no-op 不继续变换不存在的组。canonical group 的 frame/delta 在组合后真实应用，仍只进入一次 `commitWorkbook` 内容历史。

API 捕获 workbook 实例、revision、历史 generation、用户与选区；在预览、发布和 prepareView 处重新检查。受影响图表复用既有 chart source/action 权限；切片器删除复用 `BeforeSlicerChange` 撤销钩子并在成功后发 `SlicerRemoved`。权限/guard 失败不写入内容历史或成功选区。

[sheet-drawing-selection.ts](../../packages/office-studio/src/api/sheet-drawing-selection.ts) 按 API / workbook / worksheet 缓存选区；非活动宿主不能写另一工作台，替换 owner 后旧 dispose 不能清新 owner。辅助选区记录按真实 content history 的 generation/revision 索引，并按保留的 undo/redo 条目裁剪；它不建立第二份内容历史。实际内容 Undo/Redo 仍由 OfficeStudio 快照事务统一承担。

主任务在 `api/office-studio.ts` 已接 registrar，已有图片 arrangement handler 委托本 runner，避免重复注册。后续已接 ExcelWorkbench scene/typed selection/完整顺序、Canvas 与 ChartOverlay 输入路由、七项 FWorksheet wrapper 和包出口，工作表复制也使用 typed subtree ID 映射；最终 native lifecycle 在 filter/auto-height 后执行并拟合祖先组框。新增八项对齐走 `sheet.command.align-sheet-drawings`，固定叶 operation ID 前缀已修正为 `sheet.operation`；方向键走固定 `sheet.command.move-drawing` 并复用实际 skeleton 边界和连续快照历史。具体文件、状态和未运行草稿见[工作台证据](SHEET-21-drawing-workbench-affine.md)。**这些接线仍不作为实际浏览器编辑闭环的通过证据**。

实现差异：本项目用 Vue/Pinia 的应用状态、独立模型/planner 和原生 workbook 事务替代上游 DrawingManager/JSON1 mutation。上游 groupBaseBound 与 child 更新由其服务分解；本地以现有叶世界几何为权威，公开 group DTO 中 parent.transform 可作为目标 frame，ungroup 的 parent/children 作为结构快照身份检查，不重复应用其子 transform。全部上游 DTO 字段、operation 和 `SHEET_DRAWING_PLUGIN` 导入转换尚未做完整兼容验收。

## 契约草稿清点

| 文件 | `test` 声明数 | 有意义的待执行场景 |
| --- | ---: | --- |
| [sheet-drawing-groups.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-groups.contract.test.mjs) | 16 | 非连续成员、typed 同 ID、坏图、嵌套命中、多组取消组合、父子重叠、稳定层级、NaN/Infinity、只读/不可用/控制能力、resize/shear/reflection、多选、逆事务及陈旧 guard |
| [sheet-drawing-reconcile.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-reconcile.contract.test.mjs) | 9 | singleton/empty 祖先、显式删除、缺媒体保留、完整嵌套 copy、缺映射/冲突拒绝、独立复制、显式矩阵与同事务源删除/新身份 |
| [sheet-drawing-host.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-host.contract.test.mjs) | 9 | 真实 native owner/轴、资源/样式保留、直接正反重放、后续 revision 拒绝、隐藏 Both、前后权限 footprint、Delete/主图保护、图表 mask 所有权 |
| [sheet-drawing-commands.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-commands.contract.test.mjs) | 11 | 实际 OfficeStudio registry/commit/history、选区 restore、JSON reopening、多批组、宿主推导、no-op/参数拒绝、delegated arrangement、stale/Abort/权限、切片器事件、主图/源权限、隔离和 registry undo/redo |
| [sheet-drawing-model-fixture.mjs](../../packages/office-studio/tests/fixtures/sheet-drawing-model-fixture.mjs) | 0 | 本项目模型、图表 definition、Table/Slicer 实际工厂和序列化媒体 DTO；HTTP 图片未抓取/解码，不是 Office 或浏览器验收样本 |

**45 个草稿均未运行。** 集成文件中的 Vite hooks 只是待执行代码，本阶段没有启动任何 server。草稿覆盖假设仍需实际运行并修正，不能记录为「已通过」。

## 尚未验收与仍需实现

- 浏览器中混合选择、组内编辑、拖动/缩放/旋转/错切、多层嵌套、freeze/zoom、命中顺序、键盘 Delete、退出组及焦点恢复；Canvas2D/WebGPU 两后端与权限裁剪均未验收。
- 原子任务没有新增可替换 UI 控件。后续主任务的 Canvas 浮条/右键、图表输入桥/手柄及工作台夹具已静态核对为 `antdv-next`，使用 ConfigProvider/token；popup Escape/回焦和手柄 IME/修饰键保护已落盘。主题、焦点及 Windows 原生 IME 仍未运行验收，见[工作台证据](SHEET-21-drawing-workbench-affine.md)。
- 完整 shape 作者能力（Pro 预设、箭头、连接器、路径/文本编辑等）不由本批 DTO/几何接线自动完成；仍按完整 SHEET-21 目标继续实现与验收。
- slicer/timeline affine 与五类 leaf/组已补入[切片器增量](SHEET-25-slicer-timeline-affine.md)；[Pro 多目标连接](SHEET-25-slicer-connections.md)及其 UI/生命周期也已落盘，均未运行。Office major-axis、真实 UI/文件、完整交叉筛选成员状态/Pro API 和全部直接编辑仍需实现/验收。临时不可用或权限/位置锁拒绝不能缩减完整目标。
- 完整组对象 clipboard 未接线；纯模型 complete subtree copy recipe 不等于已实现组剪贴板。组操作 API/UI 和工作表复制映射已由主任务接入，真实 undo/redo、保存重开及组剪贴板仍需实现或验证。
- XLSX/DOCX/OOXML 等实际 Office 解析/写出、嵌套组和完整 affine 映射、crop/flip/资源及未知源数据保真未跑样本；JSON 草稿不能代替 Office 往返。没有依赖 Univer Server、Exchange 或商业 SDK。
- 大对象集合的主线程规划、命中/布局、Worker/Rust/WASM 路径选择、取消和内存生命周期尚无同数据集基准。本文没有声称性能改善、零拷贝或达到首屏/滚动/编辑/内存目标，WASM 与真实性能仍标「未验收」。

整体 SHEET-21、Pro、入口和文件闭环继续保留「未验收」；以上记录是阶段差异与真实缺口，不是目标范围缩减。
