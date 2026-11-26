# SHEET-25 / SHEET-21：切片器、时间线 affine 与 Ant 画布控件

记录日期：2026-11-26。状态：**实现和测试草稿已落盘，未验收**。

本批接续[共享绘图工作台](SHEET-21-drawing-workbench-affine.md)，把真实 Table 切片器和日期 Pivot 时间线接入同一绘图模型、命令、权限、内容历史及 Canvas 输入。全部 Sheets / Docs / Pro 目标继续保留；没有将本切片视为整个 SHEET-25 或 SHEET-21 完成。

遵循用户“先实现，后统一测试”：只读取源码、查询组件文档、编辑、人工静态回读及统计测试声明。**未运行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、基准或 `git diff --check`。** 下列场景和夹具没有执行；内部快照重开草稿不是独立 Office 文件验收。

## 1. 固定来源与实现边界

OSS 只读来源为 `E:/rust_test/univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 仅静态读取项目缓存 `.research.local/univer-1.6.0/packages` 的发布类型与 ES 片段，没有安装或执行商业 SDK。

| 来源 | 已查阅契约及差异 |
| --- | --- |
| `packages/sheets-drawing/src/services/sheet-drawing.service.ts`、`sheet-drawing-placement.ts` | `ISheetSlicerDrawing` 参与 Sheet drawing；Position/Both/None 锚点与 rotation/skew/flip 为模型空间 placement。这里保持 header-free CSS px 的完整矩阵和正 local extent，兼容位置字段由作者几何派生 |
| `packages/drawing-ui/src/controllers/drawing-update.controller.ts` | `_commandExecutedListener` → `SetDrawingAlignOperation` → `_drawingAlign` / 排序 → feature notification；不能把返回 true 的 operation 当作完成。本项目使用共享纯 planner 和一次 guarded 内容事务 |
| `packages/sheets-drawing-ui/src/controllers/sheet-drawing-update.controller.ts` | `_updateDrawingListener` → skeleton placement → `SetSheetDrawingCommand`；最终可见轴和持久化属于 Sheet owner |
| `sheets-slicer/lib/types/services/slicer-drawing-placement.d.ts` | `applySlicerDrawingPlacement` 复用 Sheet drawing placement，并可读取 skeleton |
| `sheets-slicer/lib/types/commands/commands/update-slicer-presentation.command.d.ts` | `UpdateSlicerPresentationCommand` 同一可撤销事务更新 hosted settings 与可选 model-space placement；kind 区分 Slicer/Timeline，参数含 unit/host/view/drawing。该类型没有 literal command ID；不宣称本地五个内容 ID 与商业内部 ID 全部相同 |
| `sheets-slicer-ui/lib/types/controllers/slicer-drawing-render.controller.d.ts` | 明确 “Sheet owns drawing placement, grouping and Transformer selection; Canvas owns control interaction.”；包含 interactive/printing、取消旧 unit/control 与宿主检查。本项目 Sheet host 保存数据，Canvas 管命中/手势，Vue/Ant 管可替换控件 |
| `sheets-slicer-ui/lib/types/menu/menu.d.ts`、`lib/es/index.js` 的 `lr` / `ur` schema 片段 | 插入、粒度/设置、连接与快速样式入口分开；格式标签只在一个 DRAWING_SLICER 聚焦时显示。四个格式根的静态顺序见下一节；动态显隐及所有连接行为尚未验收 |
| Antdv Next 官方 Slider / Dropdown / Button 文档及本地 `dist/{slider,dropdown,button}` 类型 | Slider 使用 `change` / `changeComplete`、`tabIndex` / `ariaLabelForHandle`；Dropdown 使用受控 open、popupRender、autoFocus；Button 字色用 component token。见 [Slider 文档](https://github.com/antdv-next/antdv-next/blob/main/docs/src/pages/components/slider/index.en-US.md)、[Dropdown 文档](https://github.com/antdv-next/antdv-next/blob/main/docs/src/pages/components/dropdown/index.en-US.md) |

模块遵守[完整架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：`domain` 保存可序列化模型，`features/slicer` 管来源绑定、presentation 与控制布局，`features/sheet-drawings` 管 native owner/最终轴生命周期，`features/drawing-groups` 管纯规划，`api` 管事务/历史/权限，Vue 管工作台输入与 Ant 控件。没有新增 Worker/WASM 消息或依赖；几何和控件规划目前在主线程 TypeScript，较大绘图集合的计算域选择仍需按[协议 2 与内存标准](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)实测。

<a id="entry-inventory"></a>

## 2. 入口清点与命令闭环

| 分类 | 可复查静态数量 | 状态 |
| --- | ---: | --- |
| 本地常规 Data/分析根 | 2，既有 | 切片器、时间线；本批增量 0，Sheets 常规总数仍 71（30/7/11/17/6） |
| 固定 Pro 切片器格式上下文根 | 4 | 快速样式、切片器设置、时间线设置、连接；不与常规根相加，不把两个条件设置项当作同时可见 |
| 本批顶部扩展 / 独立 More 根增量 | 0 / 0 | 没有新增工具栏产品能力；完整 Pro 动态装配与溢出数仍未核实 |
| Canvas 固定控件 slot 声明 | slicer 3、timeline 3 | 前者 clear/multiSelect/scroll，后者 clear/level/scroll；标题/滚动条件控制显示，不是六个独立业务功能 |
| Timeline level 菜单叶 | 4 | 年、季度、月、日，复用 presentation 内容命令 |
| Canvas 项目 / 日期周期 | 动态窗口 | 根据真实来源及可见区域派生；不能用样本项目数作为功能完成数 |
| 共享绘图浮条 / 右键 | 复用既有 4 控件 / 3 父项及删除 | 组、排列、对齐、删除仍共用[原命令](SHEET-21-drawing-workbench-affine.md#entry-inventory)，不重复计数 |
| 范围边缘 / 外框 handles | 编辑器内核 | 时间范围两个边缘与共享 transform handles 由 Canvas 命中/拖拽，不是原生表单或 Table 替代编辑器 |

| 标签 / 顺序 / 位置 | 入口、状态与内容链路 |
| --- | --- |
| 本地 DATA/pivot（分析），group order 30；classic：透视表、透视图、切片器、时间线 | `sheet.operation.open-slicer` grid r1c2、`sheet.operation.open-timeline` r2c2；本批沿用既有 Panel 和 `sheet.command.add-slicer`。无来源/活动宿主或编辑权限的能力判断仍按既有工作台处理 |
| 固定 Pro 切片器格式 tab order 100，styles order 0 | `sheet.menu.slicer-quick-style` order 0，r1c1 跨两行；后续[连接 UI 批](SHEET-25-slicer-connections-ui.md)已接上下文 tab/Ant 快速图库，真实运行未验收 |
| 固定 Pro 格式 options order 1 | `sheet.menu.slicer-settings` order 0 r1c1；`sheet.menu.timeline-settings` order 1 r1c1；`sheet.menu.slicer-connections` order 2 r2c1。后续[多目标连接批](SHEET-25-slicer-connections.md)已接四条件根和 shared cache/linked/owned 替换，完整运行闭环仍未验收 |
| 切片器 clear、多选、项目；时间线 clear、周期/边缘、粒度 | clear → `sheet.command.clear-slicer-filter`；item → `sheet.command.set-slicer-items`；period/range → `sheet.command.set-timeline-range`；level → `sheet.command.update-slicer-presentation`。按钮/键盘复用同一 native 内容命令与真实 Table/Pivot 筛选 |
| 垂直/水平 Ant Slider、wheel、内容键盘滚动 | 派生窗口 preview → inverse affine 位置 → 最后一次 `sheet.command.update-slicer-presentation`；切片器存 startItem，时间线存真实日期 serial scrollPosition。preview 不进持久模型/历史 |
| 绘图组合、对齐、排列、移动和删除 | shared planner → `sheet-drawing-host` → native slicer/timeline → 一次内容历史；外框方向键走 `sheet.command.move-drawing`，Enter/Space 进入内容、Escape 退出。Ctrl/Shift 在项目控件上保留筛选语义，组对象/外框仍由共享绘图控制 |

完整入口对照见[顶部表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)和[Sheets 来源表](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。本批没有用固定 slot 数或测试数关闭产品功能。

## 3. 模型、最终轴、命令及历史

- [model.ts](../../packages/office-studio/src/domain/model.ts) 增加 `SlicerModel.drawingGeometry`。`position/size` 是行列字段，`layout` 是兼容像素 AABB；author local width/height 和六个矩阵系数是独立模型，不能把 `size.rows/columns` 当成像素尺寸。
- [slicer-drawing.ts](../../packages/office-studio/src/features/slicer/slicer-drawing.ts) 复用共享可见轴解析 placement。Position 平移基底，Both 按最终世界 AABB 做 X/Y 缩放和平移，None 保留绝对矩阵。完全折叠的 Both 留住正 author extent，临时 owner 不可用；展开后恢复投影。
- [sheet-drawing-host.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-host.ts) 的 slicer/timeline 可组合、可变换；lockedPosition 阻止变换，None 不受 skeleton 移动约束。发布真实 native 模型，逆向按 revision 检查，不重复移动嵌套后代。
- [sheet-drawing-axis-geometry.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-axis-geometry.ts)、[sheet-drawing-slicer-geometry.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-slicer-geometry.ts) 在 final filter/auto-height 后、图片投影后、父组拟合前执行一次 slicer 投影；本次已明确修改的 author geometry 不再重投影。最终 axes、native geometry 和 group fitting 属于同一内容提交。
- [slicer-arrangement.ts](../../packages/office-studio/src/features/slicer/slicer-arrangement.ts) 将 native 排列委托共享 planner/host；权威顺序是 drawingState，不能再单独写 native zIndex 建立第二套层级。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 对五个 native mutation 捕获 workbook identity/revision、history generation、模型签名和用户；BeforeSlicerChange 后、candidate/prepareView 时重新检查，并检查所有改变控制的前后完整 footprint 的 View/Edit 能力。settled 事件的 payload 为每个监听器独立生成，保留已提交的 control revision；监听器重入或抛错不撤回已接受的编辑。
- [slicer-command-context.ts](../../packages/office-studio/src/api/slicer-command-context.ts) 导出 `expectedRevision?` / `host?: 'workbench'`，五个 `FSheetSlicer` mutation 均接入。默认 Facade 仍能操作绑定的 inactive unit；工作台调用显式要求 unit/sheet 活动。非法 host、版本过期、宿主切换及撤销钩子重入拒绝旧提交。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的 scoped helper 捕获 Studio/unit/sheet/user、绑定 Facade，提交及 await 后重新检查活动上下文。refresh 使用 prepared view，避免再做一次 auto-height；失败反馈只回到仍有效的宿主。presentation 的重复值不增加 revision/history，不能推断所有重复筛选动作都已有相同 no-op 语义。

## 4. Canvas、Ant 控件、焦点和取消

[slicer-renderer.ts](../../packages/office-studio/src/features/slicer/slicer-renderer.ts) 保留 local 布局并提供完整 matrix/inverse；所有场景 primitive 使用共同 `matrix/localClip`，Canvas2D/WebGPU 共用 affine 裁剪。production 使用 `controlUI: 'antd'`，停止自绘 clear、多选、项目、周期 track、粒度菜单和滚动控件；默认 Canvas 分支供 print/export、非交互场景和旧契约消费。

[SlicerControlOverlay.vue](../../packages/office-studio/src/components/spreadsheet/SlicerControlOverlay.vue) 使用 Ant Button、Dropdown、Slider、Tooltip 与嵌套 ConfigProvider。作者 accent、选中文字色和 header 字色由主题/component token 派生，disabled/hover/focus 保留库行为。Slider 的 tabIndex 与 ariaLabelForHandle 放到真实 handle，冻结副本仅有一个 Tab target。Canvas 保留 caption、日期标签、选区/范围边缘、命中和拖拽内核。

[slicer-control-overlay.ts](../../packages/office-studio/src/features/slicer/slicer-control-overlay.ts) 复用同一 affine、冻结 pane、preview 与完整 leaf paint order。逐个减去更高可见 owner 的凸 footprint，CSS 使用 nonzero path；多个遮挡重叠不会因 even/odd 再露出控件。该行为以共有矩形 drawing hit footprint 为边界，复杂透明 shape 的逐路径命中并未由本批完成。

[slicer-canvas-gesture.ts](../../packages/office-studio/src/features/slicer/slicer-canvas-gesture.ts) 捕获 unit/sheet/model 签名、geometry、revision、user、viewport；世界移动、local resize、keepRatio 和逆 affine 滚动由纯 helper 派生。[slicer-control-selection.ts](../../packages/office-studio/src/features/slicer/slicer-control-selection.ts) 统一普通、Ctrl/Meta、多选、Shift 区间与 all→null，拒绝已失效 value/index/anchor。

`SpreadsheetCanvas.vue` 将 Ant pointer 转给 Canvas owner；项目/范围/clear/滚动/粒度事件携带捕获或当前 workbook revision。Slider pointer capture 期间仅改 preview，pointerup 才提交，释放 capture 前清 pointer ID 避免 lostpointercapture 重入。键盘 Slider 完成后保留 handle 焦点，keyup 前转移焦点会取消未提交 preview。Menu 显式 owner/outside pointer listener 在卸载时移除，deactivate/宿主/模型/选区/viewport 变化关闭旧菜单或取消手势；Escape 回焦检查原上下文。真实浏览器主题、原生事件顺序、Tab、焦点和 Windows IME 仍未验证。

## 5. Clipboard、视图与文件边界

区域 clipboard 保留控制矩阵/正 local extent，粘贴对 chart/slicer/timeline 的 e/f 平移一次，完整 world AABB 决定是否完全包含；来源 Table/Pivot 按现有依赖规则处理。非法/奇异/非有限矩阵、零尺寸、伪造兼容 layout 和部分包含在 detached 规划前拒绝。values/formulas/formats/columnWidths 不夹带绘图对象。

[excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 双向 clone slicers，避免 view 与 authority 共用嵌套 geometry。内部快照保存/重开草稿保留源绑定、filter、affine 和组。没有携带可解析 Pivot 的跨 workbook timeline 仍拒绝；完整依赖迁移及组对象 clipboard 仍需实现，不能把拒绝保护记为跨 workbook 能力完成。

Office 继续复用本项目 parser/writer；本批未补齐 slicer/timeline affine、嵌套组、资源及所有 XLSX cache/parts 的独立文件往返，相关能力保持未验收，不引入 Univer Server/Exchange 或商业 SDK。

<a id="contract-drafts"></a>

## 6. 未运行的草稿

以下为人工回读并静态统计的新批，共 **98 条 test 声明 / 182 个参数化计划场景**。不是运行结果、功能计数或通过数；旧主任务 **82/98**、Sheets 图片 **66**、Docs 图片批 **57** 不并入本表，旧测试中的 host/排列断言更新也不重复计入。

| 文件 | 声明 | 计划场景 | 验证意图 |
| --- | ---: | ---: | --- |
| [sheet-slicer-drawing.contract.test.mjs](../../packages/office-studio/tests/sheet-slicer-drawing.contract.test.mjs) | 10 | 11 | 五种 leaf 与组、affine、lock、native arrangement、逆 revision、删除/作者保留 |
| [sheet-slicer-final-geometry.contract.test.mjs](../../packages/office-studio/tests/sheet-slicer-final-geometry.contract.test.mjs) | 10 | 13 | Position/Both/None、filter/auto-height/outline、折叠恢复、作者跳过、组框顺序 |
| [slicer-affine-canvas.contract.test.mjs](../../packages/office-studio/tests/slicer-affine-canvas.contract.test.mjs) | 11 | 22 | 旋转/反射/shear 命中、preview、pane、移动/resize/keepRatio、失效捕获 |
| [render-affine-command.contract.test.mjs](../../packages/office-studio/tests/render-affine-command.contract.test.mjs) | 7 | 9 | UV/local clip、winding、大平移、文本逆 bounds、Canvas2D 应用顺序 |
| [sheet-slicer-facade-history.contract.test.mjs](../../packages/office-studio/tests/sheet-slicer-facade-history.contract.test.mjs) | 19 | 53 | registry/Facade、历史/内部重开、Before/settled 重入、前后 View/Edit、五 mutation 版本/active host |
| [sheet-slicer-affine-clipboard.contract.test.mjs](../../packages/office-studio/tests/sheet-slicer-affine-clipboard.contract.test.mjs) | 8 | 17 | 区域 capture、copy/cut、AABB/矩阵/伪造拒绝、paste modes、未解析跨 workbook 依赖 |
| [slicer-control-overlay.contract.test.mjs](../../packages/office-studio/tests/slicer-control-overlay.contract.test.mjs) | 8 | 14 | affine/pane、多个重叠遮挡、owner order、View/available、nested group、Tab 去重 |
| [slicer-control-selection.contract.test.mjs](../../packages/office-studio/tests/slicer-control-selection.contract.test.mjs) | 6 | 15 | 普通/Ctrl/Meta/多选/Shift、all/empty、来源失效和纯派生 |
| [slicer-canvas-scroll.contract.test.mjs](../../packages/office-studio/tests/slicer-canvas-scroll.contract.test.mjs) | 4 | 10 | 反射/旋转/local scroll、preview/cancel、真实日期 offset、无 overflow |
| [sheet-slicer-drawing.browser.test.mjs](../../packages/office-studio/tests/sheet-slicer-drawing.browser.test.mjs) | 15 | 18 | Ant/native 项目/菜单/Slider、键盘外框/内容、Escape/焦点/主题、stale switch、混合组、冻结、历史/重开 |
| 合计 | **98** | **182** | **全部未运行** |

[slicer-drawing-model-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-drawing-model-fixture.mjs) 采用真实 Table 及 Date Pivot；scrollable 选项扩展为 48 项 Table 与连续 36 个月。Pivot 输出避开 Table，行列尺寸固定。[sheet-drawing-workbench.html](../../packages/office-studio/tests/fixtures/sheet-drawing-workbench.html) 的 `withSlicers` 保留默认三对象旧夹具，新增 control point、presentation、user remount 和真正 OfficeStudio/Pinia/Ant/Canvas 工作台路径。夹具未启动；真实 GPU、Office 文件、输入法和性能不能由静态源码证明。

## 7. 未完成与差异

- Both 的世界 X/Y AABB 锚点投影与 Office major-axis 行为尚未核实，非均匀/反射/shear/多层组的最终文件效果需真实样本对照。
- 所有新模型/命令/UI/保存重开草稿、Canvas2D/WebGPU parity、CSS 非矩形遮挡、控件事件顺序、主题/Tab/回焦、Windows IME 都保持未验收。
- 后续[连接批](SHEET-25-slicer-connections.md)已补格式四根/Ant 快速图库、共享缓存与多目标连接、linked/owned 联动和[复制/结构生命周期](SHEET-25-slicer-connections-lifecycle.md)；[SHEET-24 独立/关联创建与字段 API](SHEET-24-owned-pivot-charts.md)、[创建/字段 UI](SHEET-24-pivot-chart-ui.md)也已落盘，草稿均未运行。完整交叉筛选成员状态、日期/Pro API、缺依赖跨 workbook、同时创建表图/轴入口与 Office 仍需逐入口闭环。
- 完整 shape/connector 作者与绘制、组 clipboard、跨 workbook 依赖、Docs 跨页 anchor/混合对象/非均匀 affine、复杂 story 与 Office 资源仍保留完整目标。
- pane/owner 查询和凸多边形减法仍在主线程；遮挡碎片随更高 owner 增长，当前没有预算/Worker/Rust 大样本验证。未测设备、帧时间、long task、传输/复制及内存，不声称性能优化、达标或零拷贝。
- 后端继续暂缓；原子 native/affine/filter 变更、版本与 footprint 重验需求记录于[后端清单](../BACKEND_REQUIREMENTS.md)。selection、content focus、preview、popup、pointer capture 与 pane 属于临时 UI 状态。
