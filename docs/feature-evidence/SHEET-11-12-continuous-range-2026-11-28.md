# SHEET-11 / SHEET-12 连续区域、排序确认与取消

日期：2026-11-28。开发状态：已接线，待统一验证；验收状态：未验收。本页续接 [原固定入口清点](SHEET-11-12-sort-filter-entry-review.md)，其“待确认方案”属于此前快照，本文记录已落盘的连续范围与面板链路。

## 固定契约与入口数量

参考 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`：[expand-range.ts](../../../univer/packages/sheets/src/basics/expand-range.ts) 的角/边/合并主格探测、[sheets-sort-ui.service.ts](../../../univer/packages/sheets-sort-ui/src/services/sheets-sort-ui.service.ts) 的 `_detectSortLocation` / `showExtendConfirm`、[ExtendConfirm.tsx](../../../univer/packages/sheets-sort-ui/src/views/ExtendConfirm.tsx) 的保持/扩展/取消，以及 [sheets-filter.command.ts](../../../univer/packages/sheets-filter/src/commands/commands/sheets-filter.command.ts) 的单格四向、单行向下 smart filter。

原清点的排序/筛选常驻根保持 2，固定子项 7，本地子项 8（多一个筛选条件面板）。classic/grid 仍共用相同命令；本批未增加常驻、插件扩展、上下文或溢出根，也没有增加算法名称。新排序范围弹窗属于既有自定义排序流程。

## 实际调用链

[sheet-data-range.ts](../../packages/office-studio/src/features/sheet-sort-filter/sheet-data-range.ts) 已有共享稀疏连续范围计算，按存储值、富文本运行和 `metadata.publicBody` 判断内容，保留 0/false、空种子、角探测、单向扩展和合并主格语义。规范化产生的空 `richText:[]` 不自动算内容；公开 `p:{}` / 空 dataStream 的存在性由现有 `publicBody` 保留，不新增重复权威字段。

[use-sheet-sort-filter.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-sort-filter.ts) 已将显式扩展排序与 smart filter 接入原 Facade/命令；扩展后同步范围并保留原活动列。工作簿、工作表、用户、选区、内容版本、history generation、活跃性和阻塞状态构成宿主边界。取消或切换使旧排序失效；主线的 Facade 与命令守卫在计算返回、派发和内容发布边界继续检查。

本批补齐 [SheetDataPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataPanel.vue) 此前未消费的 `sourcePrimary` / `editable` / `requestSortRange` / `assertCurrent` props。自定义排序先用 `antdv-next` Modal / Form / RadioGroup / Radio 显示保持/扩展/取消；确认后才读取多关键字设置，初始排序键使用原活动列。取消只关闭视图，不写内容和历史。单格或单行的筛选设置使用同一 smart range，实际条件仍由原引擎求值。

面板提交使用原命令，筛选 DTO 绑定本地 WeakMap 守卫，命令观察者重入切换宿主后不能写回。排序传递 AbortSignal 与宿主断言，输入法期间禁止确认/关闭，但保留正在组合的输入控件。失活、权限/选择变化和关闭会取消旧任务。筛选提交被命令取消时不显示成功状态。

[sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 消费已有工作台 `dataRangeEditable` 能力，排序根、扩展子项、自定义排序和筛选子项同步禁用；直接命令入口仍执行相同保护。显式扩展与自定义排序允许从单格开始，普通当前区域排序仍要求多行。

## 草稿和边界

- [sheet-data-range.contract.test.mjs](../../packages/office-studio/tests/sheet-data-range.contract.test.mjs)：6 个声明，覆盖内容身份、空种子/斜角、smart 方向、合并、方向/范围错误和保持/扩展的不可变模型。
- [sheet-data-panel.contract.test.mjs](../../packages/office-studio/tests/sheet-data-panel.contract.test.mjs)：新增公开空正文快照/清除/写入/内部重开、计算取消后禁止同步回退、计算与 before-command 守卫，以及筛选命令观察者切换工作簿的真实合同草稿。
- [sheet-data-toolbar.spec.ts](../../tests/browser/sheet-data-toolbar.spec.ts)：原自定义排序草稿经过实际范围确认；新增取消零内容/历史变更、扩展后原活动列、多关键字关联行和单行 smart filter 的浏览器草稿。

所有草稿未运行。SFC、模板、TypeScript 和 mjs 只做语法读取，不执行测试、类型检查、构建、dev、浏览器、Rust/WASM 或基准。新增控件使用现有 ConfigProvider/主题，实际焦点、IME、两渲染后端、Office 文件、大样本、完整比较/日期/颜色条件、列头筛选、全部关联资源重排和协作仍未验收；本批不能作为 SHEET-11/12 全量完成证明。
