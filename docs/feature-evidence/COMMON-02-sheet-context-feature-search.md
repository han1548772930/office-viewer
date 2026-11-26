# Sheets 右键菜单接入功能搜索

更新日期：2026-11-26。涉及 COMMON-02/03、SHEET-02/03/08/16/19/26/27/30；状态：未验收，开发进度：待验证。

本批让功能搜索直接读取当前 Sheets 选区的右键命令树。普通单元格、整行、整列、全表使用已有菜单工厂、翻译、权限和目标参数；读取搜索上下文不打开右键菜单、不改变选区或内容、不派发命令。

## 来源与入口口径

固定上游：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地 `E:\rust_test\univer`。

| 来源 | 本批复用的可观察契约 |
| --- | --- |
| `packages/sheets-ui/src/controllers/render-controllers/contextmenu.render-controller.ts`，`_init` | 根据 `RANGE_TYPE.ROW/COLUMN` 路由主区域、行头、列头菜单；选区内保留目标 |
| `packages/sheets-ui/src/menu/schema.ts`，`ContextMenuPosition.MAIN_AREA/ROW_HEADER/COL_HEADER` | 快捷复制/剪切/粘贴、分类清除、行列结构、尺寸和冻结分组 |
| `packages/sheets-ui/src/menu/menu.ts` | 范围和权限分别控制显隐、禁用；命令保持统一 |
| 本项目 `components/spreadsheet/sheet-context-menu.ts` | 边框、图片、超链接、迷你图、大纲及现有结构/剪贴板命令的真实 schema 和参数 |
| 本项目 `platform/ribbon/feature-search.ts` | 搜索过滤不可执行父子项，按命令与参数去重，通过 `contextId` 使用上下文命令执行路径 |

完整顶部入口清点沿用 [Sheets 固定源码入口清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和 [顶部功能区清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。本批新增常驻、插件扩展、溢出 Ribbon 根均为 0，索引 4 种既有 Sheets 上下文：单元格、整行、整列、全表。没有新增右键业务命令；动态图树和叶子数量取决于实际选区、权限、剪贴板及对象，不把搜索数量作为功能完成证明。

## 实现与接口

| 文件 | 实际变化 |
| --- | --- |
| `packages/office-studio/src/components/spreadsheet/sheet-context-menu.ts` | 新增 `sheetContextMenuSelectionHit`、`translateSheetContextMenuItems`、`sheetContextMenuFeatureSearchContext`；保留显式 `rangeType`，缺省类型按当前表尺寸推导；递归保留命令、参数、禁用原因并去除搜索中的分隔线 |
| `packages/office-studio/src/components/spreadsheet/SheetContextMenu.vue` | 使用共享翻译函数；数量占位符、无占位符数量后缀、分隔线和原菜单顺序保持原有行为；组件公开 props 不变 |
| `packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue` | 右键和搜索共用 `createSheetContextMenuContext`；复用完整结构/移位/可见性范围权限、合并首格图片、迷你图来源权限和大纲 owner；公开 `featureSearchContext(): RibbonGroup[]` |

Canvas 增加可选 `activeRangeIndex`，由工作台传入活动选区索引；未提供时沿用末范围。多选时可让较早的整行或整列范围成为活动上下文，getter 不重排范围。读取返回值不修改菜单 open/hit/位置、工作簿版本、内容历史、活动表或焦点。大纲参数继续携带 `unitId/subUnitId/expectedRevision/expectedHistoryGeneration/expectedUserId/host`，迷你图参数保留来源范围和版本。

主工作台负责把 getter 交给 `WorkbenchRibbon.featureSearchContext`，并通过既有 `executeSheetContextCommand(commandId, params)` 执行命令。本批没有改 `ExcelWorkbench.vue`、共享 Ribbon 或其他代理负责的文件；主线接线状态以主线差异为准。

## 验证与剩余项

新增 5 个契约测试声明及 3 个浏览器测试声明，均未运行：

- `sheet-context-menu.contract.test.mjs`：显式/推导选区类型、数量翻译、完整轴参数、真实剪贴板 capability 隐藏/禁用、搜索过滤、图片/链接/迷你图参数和大纲 owner。
- `sheet-context-menu.browser.test.mjs`：实际 Canvas getter 的选区/内容/焦点不变、未打开右键菜单时的文本转数字和迷你图、行列/隐藏/保护状态变化、大纲 owner 更新、较早活动多选范围。既有行菜单断言改为检查当前完整轴 `ranges` 参数。

本批只做静态阅读和 `git diff --check`。按用户要求，测试、类型、前端构建、Rust/WASM、浏览器和基准全部后置，未启动 dev。

本批没有新增控件。涉及的生产菜单、快捷按钮和提示继续使用 `antdv-next` 的 Menu/Button/Tooltip；Canvas 输入、主题和最小原生事件桥沿用既有实现。明暗主题、真实 Windows 输入法、菜单关闭后的焦点恢复、两种渲染后端和文件保存/重开尚未运行验收。

搜索上下文在主线程读取已有模型及菜单 capability，没有改变 Rust/WASM 协议或计算权威来源。大样本读取耗时、主线程长任务、内存及搜索执行后的完整编辑/撤销/保存/重开仍未验收；本批不声称性能改善或全部功能已完成。
