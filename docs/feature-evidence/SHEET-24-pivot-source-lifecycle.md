# SHEET-23 / SHEET-24：透视报告与 owned / linked 图表的来源生命周期

记录日期：2026-11-26。状态：**实现与合同草稿已落盘，未验收**。

本批把普通透视报告、linked 图表和 owned 图表的来源/cache 映射接到现有结构编辑、工作表复制/删除/改名及剪贴板事务。用户要求先实现、后统一测试；仅做静态读取、编辑、声明清点与文档路径核对，未执行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`。

## 固定来源

OSS：`E:/rust_test/univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 固定发布包位于项目 `.research.local/univer-1.6.0/packages/`，仅查阅公开类型和 ES 文本，没有安装或执行商业 SDK。

| 固定发布包内的具体来源 | 观察及本地差异 |
| --- | --- |
| `sheets-pivot-chart/lib/types/controllers/pivot-chart-sheet-change.controller.d.ts`、`pivot-chart-sheet-copy-plan.d.ts`；对应 ES 的 `Q._getRemoveSheetMutations`、`rn` copy plan | chart 宿主生命周期与 source 生命周期分开；undo 保留 definition 及 owned source payload；自包含来源复制会迁移 source/cache。本地沿用同一现有文档内容历史，没有照搬 mutation 或私有 payload |
| `sheets-pivot-chart/lib/types/services/pivot-chart-source.service.d.ts` | owned source 与 linked report 为两种可观察来源。本地 owned source 指向原始数据范围，linked chart 指向报告生成矩阵；两者使用现有同一透视聚合器 |
| `sheets-pivot/lib/types/controllers/sheets-pivot-ref-range.controller.d.ts`；公开 ES `Fo._getStructureMutations`、`_getRemovedFieldMutations` | source/cache 保留身份，按 canonical dataFieldId 删除 Row/Column/Value/Filter/Hidden 的所有报告内实例；共享 cache 内插列共用新 data field。本地保留 report-local ID，源字段身份由 cache ID 与 sourceFieldId 共同确定 |
| `sheets-pivot/lib/types/services/sheets-pivot-cache.service.d.ts`；ES `SheetsPivotCacheService.moveSources` | 完整来源范围移动保留 cache ID 并更新 source key。本地完整矩形 cut 对共享与 disconnected cache 使用同一个范围/列映射；跨工作簿缺依赖仍由现有剪贴板契约拒绝 |
| `sheets-pivot/lib/types/controllers/sheets-pivot-copy-sheet.controller.d.ts`、`sheets-pivot-remove-sheet.controller.d.ts`、`models/pivot-sheet-copy-plan.d.ts`；ES `go/vo/_o/bo` | 自包含来源复制重映射全局 data field ID，而报告内 local instance ID 保留。本地新 cache 保留 canonical 字符串，使用 `(cacheId, sourceFieldId)` 隔离身份，未引入上游全局 DataFieldManager；这是明确实现差异 |
| `E:/rust_test/univer/packages/sheets/src/commands/commands/copy-worksheet.command.ts`、`remove-sheet.command.ts` | worksheet 命令通过 interceptor 汇总正向/逆向资源变化。本地使用 detached candidate 与现有 Studio 内容事务，不能据此声称大表分批、历史清空或性能一致 |

补查的 pivot ES 文本来自固定 URL `https://unpkg.com/@univerjs-pro/sheets-pivot@1.6.0/lib/es/plugin-VCBFwx6j.js`，保存在 `.research.local/univer-1.6.0/packages/sheets-pivot/lib/es/plugin-VCBFwx6j.js`。这是研究来源，不是运行依赖。未运行上游演示；源删除后的 cache records 续算、按 sheetName 回退和外部 linked report 删除行为仍没有运行证据。

本批没有新增顶部入口。完整分类和 Pro 入口清点继续由[工作区清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)与[透视图字段/设置记录](SHEET-24-pivot-chart-ui.md)维护，helper、缓存数和断言数不能作为功能完成证明。

## 模块与实际行为

遵循[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：`features/pivot/services` 统一 source/cache/派生矩阵生命周期，`features/slicer` 保留控制缓存及连接清理，`features/sheet-workbench` 提供 detached 资源事务，`api` 继续负责命令、权限、事件和文档隔离历史。没有解析、写出或 UI 逻辑进入该服务。

新增 [pivot-source-lifecycle.ts](../../packages/office-studio/src/features/pivot/services/pivot-source-lifecycle.ts)：

- 物理列改变前冻结 canonical 来源身份，保留报告内字段实例 ID、collapse keys 和 Show Values As 的 local references。删除源字段会移除所有重复实例，压紧 position，清理悬空 value filter、Show As 及失去后代的 collapse state。
- source/cache 在内部行列排列变化时仍 dirty；真正 no-op move 不增 revision。共享及 disconnected cache 使用同一 mapping；插列分配未配置 source field，不擅自增加报告字段，已退休序号不复用。
- 清理旧 `axisOutput` 和 drill-down，再用现有 `renderPivotTable` 重建 worksheet 输出。owned report 清理 worksheet 派生字段，由原有图表投影按需计算；没有增加第二套聚合器。
- 刷新前对共用 before 对象的宿主 worksheet 脱离引用，避免候选规划修改原始报告单元格。复制生成单元格先改到新 report ID，再重建矩阵，因此缩小后的旧生成尾部可由原清理逻辑删除。
- source 改名只更新描述性 sheetName，source sheet ID/cache ID 继续权威。source 删除保留未解析 ID/cache 引用和最后 linked cells，删除旧 axes/drill-down；同名新 sheet 不复活来源。
- 剪贴板内容失效使用实际 value/formula 差异，样式变化及相同内容不 dirty。完整 source cut 迁移所有关联及 disconnected cache，报告源字段按相同列偏移更新。

[pivot-chart-report-lifecycle.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-report-lifecycle.ts) 保留 `synchronizeLinkedPivotChartReports` 导出名，覆盖 owned 及 linked。linked source 对齐最终报告矩阵；owned source 对齐其实际原始数据范围。一次来源事务的多次同步通过 before revision 保持幂等。

## 接线与接口

`PivotSourceCacheTransform` 接收 `sheetId/range` 与可选 `column/sourceWorksheet/sourceDirty`；`PivotSourceStructureTransform` 另接收 `point/outputRange/anchorWorksheet/preserveGeneratedCells`。`reconcilePivotSourceCaches` 返回按 cache ID 分类的退休 canonical field IDs，Slicer 用它清理真正失去来源字段的控制资源。新增接口没有变更 Worker 或 Office writer DTO。

| 文件 | 本批接线 |
| --- | --- |
| [slicer-connections-lifecycle.ts](../../packages/office-studio/src/features/slicer/slicer-connections-lifecycle.ts) | cache 变换委托给新 source 服务，控制连接清理继续保留 |
| [sheet-axis-structure-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-resources.ts) | 行列插入/删除/复制/移动完成共享 report/cache mapping、worksheet 输出刷新和图表同步 |
| [sheet-cell-shift-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-resources.ts) | cell-band 变换共用 mapping，不能静默移除局部生成输出 |
| [sheet-worksheet-references.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-references.ts)、[sheet-worksheet-lifecycle.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-lifecycle.ts) | source/cache rename、源删除失效与图表 revision 接入真实 worksheet 候选 |
| [sheet-worksheet-copy.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-copy.ts) | self-contained source 在一个 copy batch 中分叉一个 cache，external source 保留原身份，copy local IDs/collapse keys 保留；刷新 copied 普通输出 |
| [sheet-clipboard-workbook-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-workbook-resources.ts) | source cut/move、copied output ownership、真实 value/formula 失效及最终输出统一在原候选内完成 |

原有 `sheet.command.insert-col/remove-col/copy-rows/move-rows`、`sheet.command.copy-sheet/remove-sheet/delete-sheet/set-worksheet-name/insert-sheet` 和剪贴板命令继续进入现有 command/permission/history/save 路径。本批不另增 Facade 或命令；真实历史、重入和 save/reopen 的断言写在合同草稿，尚未运行。

## 未运行的草稿

[pivot-source-lifecycle-fixture.mjs](../../packages/office-studio/tests/fixtures/pivot-source-lifecycle-fixture.mjs) 使用真实稀疏 source：行 2..6、列 3..7，Region/Segment/Amount/Other/Date；两份普通报告、两个 linked chart、一个 owned chart 共用 cache，Amount 有 SUM/AVG 两个 local 实例，North 折叠，并有可选 self-contained/disconnected 情形。

静态声明清点：[模型/资源合同](../../packages/office-studio/tests/pivot-source-lifecycle.contract.test.mjs) **32 条声明、41 个计划用例**；[命令/Facade 合同](../../packages/office-studio/tests/pivot-source-lifecycle-facade.contract.test.mjs) **8 条声明、11 个计划用例**；合计 **40 条声明、52 个计划用例，全部未运行**。参数化/循环展开是草稿统计，没有通过数。

草稿覆盖来源外及内部列编辑、重复 value 退役、依赖清理、内部 move/no-op、行复制/删除、cache 身份隔离、宿主输出坐标、owned 无虚构 cells、whole-band shift、候选拒绝隔离、copy/self-contained/external、rename/remove/同名重建、snapshot reopen、完整 source cut、复制后矩阵缩小、真实 clipboard value 更新、presentation-only 和同步幂等。Facade 草稿使用注册命令、同一内容 history、undo/redo、save/reopen、取消及重入；模型断言不能替代浏览器或 Office 文件验收。

## 保留缺口

1. source 删除后，本地 cache 目前没有原始 records 快照，owned report 无法继续聚合；linked 仅保留最后实际生成 cells，旧轴及 drill-down 已失效。恢复来源通过内容 undo 恢复真实数据。本批没有把该差异写成源删除后的完整缓存计算能力。
2. XLSX writer 的共享原生 cache parts、owned source payload、关系与 Office 续算保真仍需补后端；本批只记录契约，没有改 Rust/WASM 或解析写出。
3. 当前 source lifecycle 和聚合路径仍在主线程 TypeScript。后续 records/cache revision/字段映射及输出 invalidation 的 typed Worker 协议须支持取消、分批、所有权和 lease；真实大样本路径按[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)测量。没有复制字节数、TypedArray lease、传输及内存基准，不能声称零拷贝、优化或达标。
4. 未完成独立 XLSX 样本、固定上游运行行为、浏览器编辑闭环、主题、焦点和 IME 验收。本批没有修改可替换 UI 控件，UI 的 `antdv-next` 与 ConfigProvider 检查仍属统一验收。

Sheets、Docs 和全部 Pro 目标保留；上述是未完成事项，不是最终范围缩减。
