# SHEET-21 / SHEET-19 / SHEET-22：绘图工作台接线与作者几何

记录日期：2026-11-26。状态：**开发代码与测试草稿已落盘，未验收**。

本批继续[绘图组合与混合层级](SHEET-21-drawing-groups-and-mixed-order.md)，将共享选区、组/混合对象命令、图表作者坐标和图片最终轴投影接入实际工作台；同时接入[Sheets 图片剪贴板](SHEET-19-images.md)与[Docs 图片批量编辑](DOC-07-image-editing.md#image-batch)的公共出口。后续续作已接固定 `sheet.command.move-drawing`、实际工作表尺寸约束和连续微移历史。后续[切片器/时间线 affine 与 Ant 控件](SHEET-25-slicer-timeline-affine.md)已接五类叶/组、最终轴、native context 与 98/182 未运行草稿。完整形状、连接器、六类组合、Pro、Office 和性能目标保留，不能由本批接线关闭功能 ID。

遵循用户“先实现，后统一测试”的安排，本批只进行了源码读取、搜索、编辑、人工静态回读和组件文档查询。**没有运行测试、类型检查、构建、Cargo/Rust/WASM、Vite/dev、浏览器、服务、基准或 `git diff --check`。** 本文的命令和用例是待验证代码，不是执行结果。

## 1. 固定版来源与模块归属

上游只读检出为 `E:/rust_test/univer`，版本 **v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**。以下来源均指此固定版本。

| 来源 | 可复查的契约与本项目差异 |
| --- | --- |
| `packages/drawing-ui/src/menu/{schema,group.menu,arrange.menu,align.menu}.ts` | DRAWING/OTHERS 的组合、排列、对齐依次为 order 1/2/3；叶项为 2/4/8。对齐 operation 前缀实际为 `sheet.operation`，本批已修正菜单元数据中的八项 ID |
| `packages/drawing-ui/src/commands/operations/drawing-align.operation.ts` | AlignType 为字符串 0–8；八个叶 operation 转交 `sheet.operation.set-image-align`。handler 返回 true 后由共享 `DrawingUpdateController.onCommandExecuted` 监听实际处理，不能只凭 handler 返回值判断功能是否实现 |
| `packages/drawing-ui/src/controllers/drawing-update.controller.ts` | `DrawingUpdateController._drawingAlign / _applyAlignType / _sortDrawingTransform` 读取 transform frame，按整体 span/count 与一位小数步长规划分布，发布 `featurePluginUpdateNotification` 并刷新 transformer。`planDrawingAlign` 按这些坐标语义编排，最终压成一次本项目内容事务 |
| `packages/sheets-drawing-ui/src/controllers/sheet-drawing-update.controller.ts` | `SheetDrawingUpdateController._updateDrawingListener` 订阅 `featurePluginUpdate$`，保留 None 锚点或通过 skeleton 更新锚点，再执行 `SetSheetDrawingCommand`。这是 Sheets 的持久化适配链；不能据此推断 Docs 对齐未实现 |
| `packages/sheets-drawing-ui/src/controllers/shortcuts/drawing.shortcut.ts` | 绘图方向键/删除仅在 Sheets 绘图焦点下触发；公式栏、面板、形状文本输入排除。当前 Canvas、图表和共享绘图处理器保留输入法、文本输入、已消费事件及活动宿主保护 |
| `packages/sheets-drawing-ui/src/commands/commands/move-drawings.command.ts` | `sheet.command.move-drawing` 使用完整 transform bounding box、1px 平移、六位小数、skeleton 总尺寸和 None 锚点例外。已注册同一 ID；图片、图表和组选区共用 `planDrawingMove`，Shift 10px 是本项目扩展 |
| `packages/core/src/types/enum/direction.ts`、`packages/ui/src/services/undo-redo/undo-redo-group.service.ts` | Direction 是 UP/RIGHT/DOWN/LEFT = 0/1/2/3，timed 窗口为 1000ms。上游 drawing history 保存 delta，故用 append；本项目内容历史保存完整候选快照，保留第一份 undo、更新最终 redo，不能直接照搬仅替换一条 delta 的方案 |
| `packages/sheets-drawing/src/{controllers/sheet-drawing.controller.ts,facade/f-over-grid-image.ts}` | 保存 data/order，Facade 操作实际作者对象。这里保留本项目 typed drawingState 和真实 native owners，不复用上游内部服务或商业 SDK |

本项目模块职责遵守[完整架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：领域 ref/矩阵不依赖 Vue；`features/drawing-groups` 负责纯规划；`features/sheet-drawings` 连接模型 owner；`api` 负责权限、版本、事务和 Facade；Vue 工作台/Canvas 负责选区、命中、输入和 Ant 控件。几何与命令编排目前为有界主线程 TypeScript，未新增 Worker/WASM 接口；重计算和性能继续遵守[协议 2 与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。

<a id="entry-inventory"></a>

## 2. 入口清点与真实处理器

总表见[顶部入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)和[Sheets 来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。以下计数是本批涉及的展示位置，不能相加作为功能完成数。

| 分类 | 本地静态数量 | 边界 |
| --- | ---: | --- |
| 图片/背景常规 Ribbon 根 | 2 | 已存在于 `sheet-ribbon.ts`，本批纠正中央清单的旧记录；子菜单另计 4 |
| 共享绘图右键父入口 | 3 | 组合/排列/对齐；叶项 2/4/8，另有一个本地删除叶项，共 4 个顶层菜单节点、15 个叶动作 |
| 共享绘图 Canvas 浮条 | 4 | 组合/排列/对齐 Dropdown 和删除 Button；复用上述动作，不计为四个新增产品能力 |
| 本批新增常规/顶部扩展根 | 0 | 对齐、组操作和作者编辑由上下文呈现；没有创建新的常规 Ribbon tab |
| 独立 More 功能 | 0 | More 复用入口集合；动态插件/Pro 溢出完整数仍未核实 |
| Docs 图片批量常规根增量 | 0 | 选区/浮条/右键扩展见[Docs 批量入口清点](DOC-07-image-editing.md#image-batch)，不重复计入常规 Ribbon |

| 标签 / 分组 / 位置 | 标签与命令 | 显隐、状态与实际链路 |
| --- | --- | --- |
| INSERT/MEDIA，order 0，grid r1c1 跨两行；classic 媒体组首项 | 图片根 `sheet.menu.image`；子项 `sheet.command.insert-float-image`、`sheet.command.insert-cell-image` | 无活动 sheet 或不可编辑时禁用；`executeRibbon` → 图片 picker/preparation → 图片命令 → 模型/资源和既有内容历史 |
| VIEW/DISPLAY，order 0.5，grid r1c2 跨两行；classic 冻结之后 | 背景根 `sheet.menu.worksheet-background-image`；设置/删除命令 | 根检查 sheet 编辑能力，删除还检查背景存在；API 保存真实背景资源，历史/重开仍待验证 |
| DRAWING/OTHERS，order 1，子 0/1 | 组合/取消组合；`sheet.command.group-sheet-image` / `sheet.command.ungroup-sheet-image` | 组合至少两个可组成员；取消组合需组选中。Canvas Ant 菜单 → `useSheetDrawings` → registrar → detached planner/native owners → 一次 guarded commit |
| DRAWING/OTHERS，order 2，子 0–3 | 置顶/前移/置底/后移；`sheet.command.set-drawing-arrange` | 选中对象保持相对顺序；不可编辑禁用；边界 no-op 不写内容历史 |
| DRAWING/OTHERS，order 3，子 0–7 | 左/水平居中/右/上/垂直居中/下/水平分布/垂直分布；`sheet.command.align-sheet-drawings` | 共享选区至少两个可变换对象；所有对象权限/可用性在命令层再次检查；输入为字符串 AlignType，完整矩阵只平移 e/f |
| 共享 Canvas 浮条/右键与键盘 | 删除；`sheet.command.remove-sheet-drawings` | 组展开后删除选中 native owners，源 Table/Pivot 保留；只读普通方向键/删除被消费但不提交内容 |
| Sheets 绘图焦点下的方向键 / Facade | 移动；`sheet.command.move-drawing`，参数 `direction: 0/1/2/3`、可选 `step: 1/10` | 单图片、单图表、组和混合选择共用 registrar；Canvas 图片裁剪仍由裁剪会话消费方向键。此为键盘/API 入口，不新增 Ribbon 根或浮条控件 |

对齐菜单元数据的固定八项为 `sheet.operation.set-drawing-align-{left,center,right,top,middle,bottom,horizon,vertical}`；实际共享 UI 和 Facade 走本地 canonical 内容命令。没有把元数据记录误称为全部上游 operation 已注册。

人工回读当前常规 schema 后，Sheets 为 **71 根（30/7/11/17/6）**：INSERT 为 media 4 + cells 3，VIEW 为 display 4 + visibility 2。此前 69（30/6/11/17/5）是缺少图片/背景登记的旧清单；固定 OSS 仍为 63（30/9/12/6/6）。该数字包含条件项、排除 children/context，未执行 factory 或运行态装配。Docs 常规根保持 52（33/16/3）。

## 3. 共享选区、工作台和公共 API

- [use-sheet-drawing-selection.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-drawing-selection.ts) 用 typed refs 区分同 ID 的不同对象；去重、清理不存在的对象、排除已选父组的后代，additive 选择替换冲突祖先/后代。native 同步有重入保护。
- [sheet-drawing-selection.ts](../../packages/office-studio/src/api/sheet-drawing-selection.ts) 区分“从未发布”和“主动空选区”；按 Studio/unit/sheet 恢复。旧 owner 释放不清新 owner，选区只借用真实内容历史的 generation/revision，不建立第二套内容撤销栈。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 提供真实 scene、完整叶顺序、能力、typed selection 与命令宿主；开始单元格编辑或选择表头清绘图选区。[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 与 [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 共用这份选区和顺序。
- Canvas 判定真正顶层 owner，native chart DOM 只是输入桥；图片在图表 DOM 下方时仍按绘图 paint/hit 顺序接收事件。组/混合 pointerdown 后抑制延迟 click，避免将组重新选成单图表。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 已增加 `FWorksheet.groupDrawings / ungroupDrawings / transformDrawings / arrangeDrawings / alignDrawings / moveDrawings / removeDrawings`，`Enum.DrawingAlignType` 使用固定字符串，移动复用已有数值 `Enum.Direction`。bound unit/sheet/selection 在额外 transform 字段之后写入，不能被调用方重定向。
- Docs 的 `FDocument.alignImages / transformImages / arrangeImages / removeImages`、`Enum.DocumentImageAlignType` 和相关类型已接公共出口；布局 provider 和批量权限/历史由 Docs 模块实现，见[Docs 图片证据](DOC-07-image-editing.md#image-batch)。

<a id="model-lifecycle"></a>

## 4. 作者几何、最终轴投影与历史

[chart-drawing-affine.ts](../../packages/office-studio/src/features/charts/chart-drawing-affine.ts) 将 renderer local point、矩形四角、线轮廓和 pointer delta 投影到工作表/视口，反射、错切、旋转和作者 local size 均保留。先在 local frame 裁剪作者轮廓，再投影，避免旋转时外泄。native chart hitbox 使用完整叶顺序作 zIndex；标题、绘图区、元素浮栏及拖动使用同一投影。

冻结 pane 从 Canvas 命中同步交给 `routeCanvasEvent`，手势保存所用 pane；标题/plot/legend/content 的后续事件复用它。滚动、宿主/工作表切换和内容过期会取消旧手势。图表普通移动只提交 position，尺寸操作只提交 size，避免平移顺带改写 renderer 作者尺寸。普通单图表方向键现走统一 move 命令；公共 `chartDrawingKeyboardTranslation` 也委托同一 planner。尺寸/plot 手柄排除 composition、keyCode 229、Ctrl/Meta/Alt、已消费事件和过期/忙碌宿主。

[sheet-drawing-image-geometry.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-image-geometry.ts) 对最终轴变化重投影未被本次作者/结构命令改过 placement 的 affine 图片。`reconcileWorkbookDrawingStates` 在公式、filter refresh 与 auto-height 之后运行，standalone `prepareAxisWorkbookView` 复用同一最终 reconcile。先前已改 placement 的图片跳过，避免同事务重复投影；之后重新拟合祖先组框并修正缺失 owner/order。

图片平移统一使用 `sheetImageWithDrawingGeometry`：local extent 与 a/b/c/d 未改变时保留作者 size/crop/rotation，只更新兼容锚点与 e/f；真正变换按既有完整 helper 发布。HTML/PNG 与剪贴板接线见[Sheets 图片证据](SHEET-19-images.md)，不把 PNG 编码记为 Office 往返。

对齐、组/混合 transform、排列和删除都先规划 detached candidate，再重新检查用户、内容实例/revision、history generation、宿主、选区、取消和前后占用范围权限，最后进入同一内容事务。no-op 不追加历史；内容 Undo/Redo 使用 OfficeStudio 既有文档隔离历史，选区恢复是辅助视图记录。

<a id="keyboard-move"></a>

### 4.1 统一移动与连续历史增量

[drawing-group-move.ts](../../packages/office-studio/src/features/drawing-groups/drawing-group-move.ts) 对完整 affine bounds 求所有约束的交集，以同一个世界 delta 平移全部所选对象，保留 a/b/c/d、local extent、资源和组内距离；嵌套组的所有叶只变换一次。有 anchored 后代的组按其整个组框约束；None 图片的 `constrainMovement: false` 是临时 owner 能力，不写入 drawingState。

[sheet-drawing-move.ts](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-move.ts) 复用 `toExcelWorkbookView` 的实际行列数量及共享可见轴，含单元格扩展、filter、outline、auto-height 和默认宽高。它不再用 Excel 最大行列数量充当 skeleton 当前边界；几何统一排除 headers/scroll/zoom。极大或越界对象仍沿用固定 min/max clamp，包括零分量轴的纠偏，极端输入及最终边界保真待真实验证。None 对象仍受到本项目范围权限/最大可表示范围约束；不能将本批推断为任意位置均已验证。

[drawing-move-history.ts](../../packages/office-studio/src/features/drawing-groups/drawing-move-history.ts) 只保存一份最近 receipt 和 1000ms 窗口，内容快照继续由现有 history Store 管理。每次移动先准备、检查前后 footprint，再用 `commitWorkbook.amend` 保留第一次 undo 和最终 redo；没有丢弃中间移动结果或绕过权限。相同 typed 选区内换方向或 Shift 步长可继续同一 burst；超时、时钟倒退、内容/保存点/历史版本、用户、选区及 owner 变化开始新记录。其他命令、Undo/Redo 意图、workbook dispose 和 registrar 释放清除临时会话，不增设 timer 或第二份内容历史。

[sheet-drawing-selection.ts](../../packages/office-studio/src/api/sheet-drawing-selection.ts) 新增临时 selection generation：相同 typed ref 集的重复同步不改变代数，选择改走又恢复、owner 替换或释放会改变代数。旧 owner 释放仍不会改变新 owner。提交 receipt 用于选区辅助历史，过期内容/宿主/取消回调不能重新记录微移会话。单图片绘制/裁剪/输入仍在原 owner，普通微移通过 `nudge` 路由到共享命令；无 drawing scene 的独立图片 Canvas 保留既有几何更新接口。

## 5. Ant 控件与输入边界

本批涉及的图片/绘图浮条、图表输入桥、resize/plot 控件、上下文菜单、Tooltip/反馈及浏览器夹具使用 `antdv-next`，承接统一 ConfigProvider 和主题 token。Canvas 承担内容、轮廓、命中与拖动；没有引入原生表单控件或第二种 UI 库。

混合右键 Dropdown 使用唯一 `classes.root` 和 `popupRender` 处理 Escape。关闭时仅在焦点仍属于该 popup、host/sheet/revision/selection 未变化且组件仍活动时归还 Canvas；用户转移焦点后不抢回。内容、选区、视口、编辑、用户变化和 deactivate/unmount 关闭旧 popup。尺寸/plot 输入法保护已静态补齐，但原生 Windows IME、主题切换与实际焦点链尚未运行验收。

<a id="contract-drafts"></a>

## 6. 主任务测试草稿

| 文件 | `test` 声明 | 参数化后计划场景 | 待验证的行为 |
| --- | ---: | ---: | --- |
| [chart-drawing-affine.contract.test.mjs](../../packages/office-studio/tests/chart-drawing-affine.contract.test.mjs) | 8 | 10 | affine 正反投影、local delta、边界微移、裁剪后轮廓和作者 local size |
| [sheet-drawing-align.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-align.contract.test.mjs) | 10 | 17 | 六种对齐、两种分布、固定 span/count/一位小数、嵌套与只读/不可用原子性 |
| [sheet-drawing-selection.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-selection.contract.test.mjs) | 9 | 9 | dedupe、父子冲突、native 重入、Studio/工作表切换、主动空选区、旧 owner 释放 |
| [sheet-drawing-final-image-geometry.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-final-image-geometry.contract.test.mjs) | 8 | 10 | 最终尺寸/filter/auto-height、隐藏 Both 恢复、placement 跳过、父组重新拟合 |
| [sheet-drawing-ui-facade.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-ui-facade.contract.test.mjs) | 9 | 9 | 真实 registry/Facade、单次历史、取消/权限/宿主 guard、最终轴提交、复制与内部重开 |
| [sheet-drawing-workbench.browser.test.mjs](../../packages/office-studio/tests/sheet-drawing-workbench.browser.test.mjs) | 12 | 12 | 实际工作台混合选择/对齐/组拖动与取消、resize/rotate、覆盖命中、冻结作者交互、主题/焦点/只读、批删；移动增量含混合/组/单图连续历史及输入/保存点屏障 |
| [sheet-drawing-move.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-move.contract.test.mjs) | 10 | 15 | 四方向、完整 bounds、混合共同边界、嵌套组、None、六位精度和批量拒绝 |
| [sheet-drawing-move-history.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-move-history.contract.test.mjs) | 7 | 7 | 1000ms 窗口、续期、receipt、保存/重开/内容/选区屏障和释放 |
| [sheet-drawing-move-facade.contract.test.mjs](../../packages/office-studio/tests/sheet-drawing-move-facade.contract.test.mjs) | 9 | 9 | 实际 registry/Facade/历史、各步结果、保存点、选区 owner、目的权限、取消和观察事件重入 |
| 合计 | **82** | **98** | 全部未运行；声明数和计划场景数不代表功能数或通过数 |

[sheet-drawing-workbench.html](../../packages/office-studio/tests/fixtures/sheet-drawing-workbench.html) 使用实际 OfficeStudio/Pinia/ConfigProvider/ExcelWorkbench/Canvas/ECharts、真实 PNG 和独立工作表。测试读取 native title/plot target，按 frozen pane 投影；没有打开浏览器。静态复核把历史草稿中的无变化“置顶”改为实际“置底”，并补手柄 composition/修饰键保护场景；仍须统一运行后修正夹具或实现。

浏览器文件原 9 条现为 **12 条**，增加混合连续微移/完整 undo-redo、affine 组在 zoom 下微移、单图共享命令及输入/保存点屏障 3 条。初始工作台批为 53/64，移动增量增加 29 条声明、34 个计划场景，故当前为 82/98。此前组合批 45 条、Docs 图片批新增 57 条、Sheets 图片批 66 条各在原证据记录，不并入本表。合成 composition 只测试事件保护，不能当作原生 Windows IME 验收。

## 7. 保留的缺口

- 全部真实 UI、内容/选择历史、保存重开、冻结/缩放、明暗主题、焦点、原生剪贴板、Windows IME 与两种 Canvas 后端均未验收。
- [slicer/timeline affine、五类叶/组、final axes 和 Ant 控件](SHEET-25-slicer-timeline-affine.md)已落盘但未运行，独立 98/182 草稿不并入本批 82/98；[Pro 多目标连接](SHEET-25-slicer-connections.md)及其 UI/生命周期已另批落盘。完整 shape/connector/SmartArt 作者与渲染、交叉筛选成员状态/Pro API、Office major-axis 与真实文件仍有缺口；DTO 和临时拒绝不能作为最终交付。
- 组对象完整 clipboard、跨 workbook 依赖、Office affine/嵌套组/crop/资源写回与独立重开仍需实现或验证。
- 固定 `sheet.command.move-drawing`、实际 skeleton 范围与连续微移快照历史已落盘，尚未运行验收；极端组/None/越界修正和直接作者 resize 的最终效果仍需与固定行为核对。
- Docs 跨页 anchor 迁移、混合对象/非均匀变换、Office 资源与复杂 story/segment 组合仍保留，见 Docs 证据。
- 图片/绘图的批量规划、索引、布局、Worker/Rust/WASM、取消/内存与同数据集大样本指标尚无实测；没有声称优化、零拷贝或性能达标。

后端继续暂缓；持久 drawingState/affine/native owner、资源关系、Docs batch/segment 权限和原子确认需求同步到[后端需求](../BACKEND_REQUIREMENTS.md)。selection、preview、layout provider、popup/PNG ticket 和原生输入桥不进入第二份内容日志。
