# SHEET-07 / SHEET-20：View 入口与视图链路续核

日期：2026-11-27。固定参考为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；本次读取的 `E:\rust_test\univer` 工作树无在途改动。

状态：十字高亮既有工作台接线已补，测试草稿未运行。冻结、网格线、视口、缩放和十字高亮仍未验收；本次记录不代表完整 Sheets 或 Pro 功能已完成。

## 六个根入口与分类

固定 [Sheets UI schema][schema] 的 View 声明四个根；[drawing schema][drawing-schema] 增加背景图，[crosshair schema][crosshair-schema] 增加十字高亮，合计六个根。常驻核心声明为 4、OSS 插件扩展为 2；扩展根属于同一个常规 View 集合，不能重复相加。classic 按分组和 order 排列，grid 坐标属于声明位置，背景图与缩放的上游 column=2 声明存在重叠，运行时布局仍需核实。

| 标签 / 分组 | order / grid 声明 | 上游入口与动作 | 本地入口及当前状态 |
| --- | --- | --- | --- |
| 冻结 / DISPLAY | 0；r1c1、跨两行、showLabel | `sheet.toolbar.sheet-frozen`；六项 selection / row / column / first-row / first-column / cancel | `sheet.command.freeze-selection` split，四项 selection / first-row / first-column / unfreeze；行/列独立入口缺失 |
| 背景图 / DISPLAY | 0.5；r1c2、跨两行、showLabel | `sheet.menu.worksheet-background-image`；add / delete 两项 | 同名根及既有两项，属于 SHEET-19，本批未改 |
| 缩放 / DISPLAY | 1；r1c2 | `sheet.menu.zoom-ratio` → `sheet.command.set-zoom-ratio-from-toolbar`；50/75/100/125/150/175/200/400 | `sheet.operation.set-zoom`，相同八档，位于 r1c3 |
| 十字高亮 / DISPLAY | 2；r2c2 | toggle / set-color；16 个主题 token，色板八列两行 | 同名两项 operation、八色菜单，位于 r2c3；本批恢复实际接线 |
| 网格线 / VISIBILITY | 0；r1c1 | `sheet.command.toggle-gridlines`；激活态跟随活动表配置 | 同名根与持久模型字段，已接 `commitHistory` |
| 全屏 / VISIBILITY | 1；r2c1 | `ToggleFullscreenOperation` / `FullscreenMenuItemFactory` | `sheet.operation.toggle-fullscreen`，使用工作台宿主与 fullscreenchange；本批未改 |

冻结六个子项、背景图两个子项、缩放八个选项及十字高亮十六个颜色选择分别记录，不将其算成新的根。窄屏 More 复用现有根与子项，不增加功能数；本批未执行窄屏运行态清点。

同功能的上下文声明另计：冻结在 MAIN_AREA、COL_HEADER、ROW_HEADER 各一个根，分别包含 6 / 3 / 3 个动作声明；底栏另声明网格线一个根及十字高亮一个根。背景图插件另在 MAIN_AREA、COL_HEADER、ROW_HEADER 声明同一个 `SaveCellImagesCommand`，属于 SHEET-19 的三处挂载，不是三个新功能，也不属于 View 常驻根。对象浮动工具及其他 Pro 注册不在本节重新计数，完整目标继续按全局入口表推进。

冻结及网格线的上游 factory 使用 workbook editable / worksheet edit 权限；冻结行、列的上下文动作根据当前末选区 primary / range 首坐标显隐。缩放、十字高亮按 Sheets 宿主显隐，十字高亮不要求内容编辑权限。上游颜色选择跟随 ThemeService；本地颜色列表仍是固定 RGBA，完整主题色板对齐尚未完成。

## 模型、坐标、历史与保存

[冻结 UI 命令][freeze-ui-command] 使用当前末选区的 primary 首坐标，缺失 primary 时才取 range 首坐标；0 边界转为 1，整行 / 整列选择仅冻结对应轴。它同时使用 `sheetViewStartRow/Column`，将 `startRow/startColumn` 与 `ySplit/xSplit` 分开表达从当前可视起点冻结的区间。[SetFrozenCommand][freeze-command] 和 [SetFrozenMutation][freeze-mutation] 保存这些字段，并登记正向 / 逆向 mutation。

本地 [sheet-workbench.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-workbench.ts) 的 `frozenPaneCountsForSelection` 仅返回 range 左上角的行列计数；[ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 的冻结选择调用该函数，首行 / 首列固定为 1，行列标题上下文则采用 index+1。它没有表示非零可视起点的冻结模型，也没有按 primary 或 rangeType 选择轴。本批保留这些差异，未以局部计数修改宣称完整冻结对齐。

[SpreadsheetGeometry](../../packages/office-studio/src/rendering/spreadsheet/layout.ts) 已真实建立滚动正文、冻结行、冻结列及冻结交点四个 pane；cellRect、hitTest、scrollForCell、选区及输入布局共用尺寸索引和冻结变换。[SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 按工作表保存运行期 scrollLeft/scrollTop，工作台按工作表保存选区、anchor、activeRangeIndex 和 zoom。不能将现有实现概括为只画一条冻结线。

[excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) 双向保留 frozenRows / frozenCols 与 showGridlines。[Facade](../../packages/office-studio/src/api/facade.ts) 的冻结 / 网格线设置通过 `mutateWorkbook`；工作台目前直接修改视图模型并调用 `commitHistory`，命令入口与 Facade 仍需统一验收。上游网格线同样是持久 mutation 和内容历史；本地网格线不应改成纯视图临时值。

[缩放 toolbar 命令][zoom-toolbar] 限定 10–400% 并转入 [SetZoomRatioCommand][zoom-command] / [SetZoomRatioOperation][zoom-operation]。本地八档选择及底栏减 / slider / reset / 加四项控件复用工作台 zoom 值，实际 Canvas 使用 zoom/100；运行期每表会话存在，但公开 `zoom/getZoom` 及原生快照重开后的每表 zoom / scroll 恢复没有闭环。上游对可见单元格编辑器的缩放限制、本地输入中缩放行为以及 chart sheet 的视口政策另留未核实项。

## 本批十字高亮实际改动

[上游 operation][crosshair-operations] 的 toggle 只切 enabled，set-color 在设置颜色时开启高亮。[服务][crosshair-service] 默认关闭，维护独立视图状态及主题 token，不登记内容 mutation。[render controller][crosshair-render] 跟随选择及活动表，排除覆盖整行 / 整列的选择，合并多选择的外部横纵范围；公式引用选择使用单独服务。

本批开始时，现有 Ribbon 根、八色菜单、Canvas prop 和 [scene](../../packages/office-studio/src/rendering/spreadsheet/scene.ts) 绘制已经存在，但 ExcelWorkbench 没有 enabled / color 状态、operation handler 或 Canvas 传参，入口无法产生高亮。本批补回以下实际链路：

- 工作台维护默认关闭的 enabled 和当前 RGBA；toggle 切换，set-color 更新并开启。
- Ribbon 反映 enabled，八色菜单反映当前颜色，Canvas 接收同一视图状态并复用现有跨 pane 裁剪。
- 两项 operation 跳过通用 finishEdit；有输入会话时调用现有 `restoreCellInputFocus`，无会话时恢复 Canvas 焦点。未写工作簿、modified 或内容历史。

本批沿用 `antdv-next` Button / Menu / Popover 及既有 ConfigProvider，不新建原生控件。当前色板仅八色，与上游十六个主题 token 不等价；底栏十字高亮入口、独立 enable / disable Facade、公式引用选择切换和服务跨工作台实例生命周期尚未接通。本批没有修改 `crosshairHighlightRanges` 算法或 Worker/WASM 协议，没有性能优化结论。

## 草稿与未验收项

[sheet-crosshair-highlight.contract.test.mjs](../../packages/office-studio/tests/sheet-crosshair-highlight.contract.test.mjs) 移除以源文本正则代替行为的断言，直接检查真实 Ribbon / Menu 的 operation、当前颜色，以及冻结、隐藏轴、滚动、缩放场景中的单元格中心裁剪和去重。原场景误用 x 坐标判定选区被覆盖，现改为验证选区中心不在高亮矩形内。

[sheet-view-toolbar.spec.ts](../../tests/browser/sheet-view-toolbar.spec.ts) 与 [独立配置](../../tests/browser/sheet-view-toolbar.playwright.config.ts) 是未运行草稿：覆盖 toggle / 选色自动开启、内容与历史不变、当前颜色、Canvas 焦点、单元格输入草稿保留与取消。未验证 Windows IME、明暗主题、More、原生文件保存/重开、WebGPU / Canvas2D 实际绘制或大样本性能；SHEET-07 / SHEET-20 保持未验收。

本批只读检查：限定文件 diff 空白检查通过，十字高亮 `.mjs` 草稿通过 `node --check` 静态语法解析；本文 10 个本地文件目标与 12 个固定版上游源码目标均存在，提交一致。浏览器草稿没有运行，未执行类型、构建或性能验证。

[schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/menu/schema.ts
[drawing-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-drawing-ui/src/menu/schema.ts
[crosshair-schema]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-crosshair-highlight/src/menu/schema.ts
[freeze-ui-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/set-frozen.command.ts
[freeze-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/commands/set-frozen.command.ts
[freeze-mutation]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets/src/commands/mutations/set-frozen.mutation.ts
[zoom-toolbar]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/set-zoom-ratio-from-toolbar.command.ts
[zoom-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/commands/set-zoom-ratio.command.ts
[zoom-operation]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/operations/set-zoom-ratio.operation.ts
[crosshair-operations]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-crosshair-highlight/src/commands/operations/operation.ts
[crosshair-service]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-crosshair-highlight/src/services/crosshair.service.ts
[crosshair-render]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-crosshair-highlight/src/views/widgets/crosshair-highlight.render-controller.ts
