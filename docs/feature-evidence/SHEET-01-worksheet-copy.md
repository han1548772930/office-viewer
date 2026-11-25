# SHEET-01 工作表复制闭环（2026-11-25，未验收）

固定源 `E:/rust_test/univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

## 源码契约

- `packages/sheets/src/commands/commands/copy-worksheet.command.ts`：`sheet.command.copy-sheet`，`unitId/subUnitId`，克隆配置、唯一名称、新 sheet ID、源表后插入、激活副本；interceptor 接 `targetSubUnitId/targetSubUnitName/copyContext` 组织插件资源与逆 mutation。
- `packages/sheets/src/facade/f-workbook.ts`：`duplicateSheet`（约 702 行）、`duplicateActiveSheet` 复用同一同步命令。
- 本地原 `ExcelWorkbench.vue::copySheet` 只复制表内数据和设置，明确提示关联图表、透视表、结构化表格不创建；本批删除此范围缩减行为。

## 本地实现

`features/sheet-workbench/sheet-worksheet-copy.ts` 独立候选规划；`api/sheet-lifecycle-commands.ts` 注册固定命令 ID；UI 和 FWorkbook 的 duplicateSheet/duplicateActiveSheet/copySheet 共用命令。原子发布、公式重算、内容历史、撤销重做及观察者隔离沿用 OfficeStudio 的工作簿事务。

复制源宿主上的 tables/pivots/charts/slicers；分配新对象及列/透视字段 ID、工作簿唯一表名和切片器名，保留位置和选择条件。数据源指向被复制表时改副本，其他表/外部 unit 数据源保留。透视图和切片器绑定、透视生成单元格 metadata 与下钻字段映射同步重映射。

稀疏单元格、完整轴 span、隐藏行列、冻结、范围和未知元数据直接克隆，避免按可见剪贴板范围截断。显式自表公式、结构化表名、迷你图和内部链接重映射；字符串、外部 A1 与 3D 引用保持。校验/条件规则、大纲及保护分配新身份。整个源表可见权限包含远处保护区域，不只检查当前可见区域。

## 差异及未核实

- 本地名称以 ` (2)` 起，保持 Excel 31 UTF-16 上限和大小写无关冲突；不复刻上游内部 ID/插件 interceptor。
- 新保护规则的 owner 是执行复制的用户，原规则的编辑/查看策略与 allowedUserIds 保留，permissionId 独立；完整持久 ACL 和公开 Pro 复制权限仍待服务端对应。
- local defined names 的 sheetId/subUnitId/sheet/localSheetId 与 refersTo 保留/重映射；其他上游扩展作用域键、外部结构化表引用、媒体和未建模绘图资源复制尚未核实，不能据此声称完整 Office 保真。
- 上游大表分块，本地目前是候选快照；Worker 路径、真实大样本内存/耗时未验收。
- `sheet-worksheet-copy.contract.test.mjs` 已写资源身份/绑定、稀疏轴、自表/外部引用、名称、保护和 DTO 草稿，未执行。API 历史、浏览器主题/焦点/IME、JSON/Office 保存重开和资源独立编辑场景仍待统一验证。

本批没有运行测试、类型检查、构建、Rust/WASM、浏览器、dev 或基准；SHEET-01 保持未验收。
