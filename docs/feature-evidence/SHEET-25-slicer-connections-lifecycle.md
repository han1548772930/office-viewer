# SHEET-25：共享切片器连接生命周期与快照保留

记录日期：2026-11-26。状态：**实现和契约草稿已落盘，未验收**。

本批负责共享来源/控件缓存穿过输入归一化、视图适配、剪贴板、工作表复制/删除/改名和结构编辑时的资源生命周期。连接命令、筛选计算、Facade 和连接 UI 由主任务及 UI 任务衔接。本记录不表示整个 Sheets、Docs、SHEET-25 或 Pro 范围已经完成。

遵循用户“先实现，后统一测试”：仅静态读取与编辑，**未执行测试、类型检查、构建、Cargo/Rust/WASM、dev、服务、浏览器、性能基准或 `git diff --check`**。下列契约草稿尚未运行，内部模型快照重开也没有经过独立 Office 文件验收。

## 固定来源与可观察契约

OSS 基线：`E:/rust_test/univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 类型基线：项目内 `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/`，只静态查阅，没有安装或执行商业 SDK。

| 具体来源 | 查阅结果、实现差异及未核实项 |
| --- | --- |
| `types.d.ts` 的 `SheetSlicerSourceSpec`、`ISheetSlicerConnections` | 透视来源具有 `pivotCacheId/sourceFieldId`，报告集合同时容纳 `pivotTables` 和 `ownedPivotCharts`。本项目使用同样的稳定来源身份及分类型引用；本地 JSON 模型合并 slicer/timeline 控件集合，未复刻私有资源格式 |
| `facade/f-slicer.d.ts` 的类说明、`getCacheId/getSource/getCache/getConnections/setConnections` | source/cache/filter 共享，presentation 属于各 view；查询返回防御性数据，来源身份不可经 presentation 替换。本批保留共享缓存及独立 affine/settings；公开 getter/setter 与权限/历史的实际调用验收由主任务继续完成 |
| `facade/f-pivot-table.d.ts`、`facade/f-pivot-chart.d.ts` 的 fieldId 参数 | 明确使用 `getSourceFieldsInfo().id` 所代表的原始来源 ID，不能使用名称或报告内字段实例 ID。本地 `sourceFieldId` 与 `field.id` 分开，缓存记录来源列；旧输入按原来源的相对列迁移 |
| `models/sheets-slicer.model.d.ts` 的 `getSnapshot/prepareBatch/_syncRetainedSources` | 持久资源与派生状态分离，所有参与模型先提交再发布。本批使用 detached candidate，源表单元格不因资源归一化再次深复制；最终提交、事件和历史仍走现有 Studio 内容事务 |
| `services/slicer-sheet-effects.d.ts` | 原生 sheet effects 具有正向/逆向模型，未知生成效果不能静默省略。本批将所有共享缓存加入同一个 clipboard/worksheet candidate，不用 view-only 状态替代资源变更 |
| `utils/slicer-persisted-state.d.ts` | 持久 guard 使用 JSON 语义，包括省略的可选键。本地复制保留 omitted connections 的形状，连接集合身份比较沿用主任务的集合比较；全部 stale guard 和监听器重入属于主任务验证范围 |
| `E:/rust_test/univer/packages/sheets/src/commands/commands/copy-worksheet.command.ts`、`remove-sheet.command.ts` | 固定版工作表命令通过 interceptor 组织内容 mutation 和逆向历史。上游大表分批/清空历史路径与本地完整 detached 事务不同；本批没有声称两者性能或大表历史行为一致 |

本批没有新增顶部常驻、扩展、More 或上下文根入口。共享连接入口的标签、条件、命令和完整 Pro 入口清点仍由主任务及[已有 SHEET-25 清点](SHEET-25-slicer-timeline-affine.md#entry-inventory)维护，不把 helper、资源数量或测试声明数计为产品入口完成数。

模块归属遵循[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：`domain` 负责可序列化缓存/报告模型；`features/slicer` 负责复制与生命周期；`features/sheet-workbench` 负责 detached 资源事务；`api` 保持命令及历史边界；parser/view adapter 只做输入和模型转换。计算在现有主线程 TypeScript 资源规划中执行，没有新增 Worker/WASM 消息、DOM 操作或第二套筛选引擎。大资源集合的路径选择及内存/取消指标仍需按[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)实测。

## 实际落盘行为

[slicer-connections-lifecycle.ts](../../packages/office-studio/src/features/slicer/slicer-connections-lifecycle.ts) 提供三项可测试资源操作：

- `copyPivotReportCache`：保留原始来源字段 ID 和 `nextFieldId`；相同工作簿、相同来源复用 source cache；来源改变时在同一个复制 batch 内生成一个共享新缓存。新来源列坐标按偏移更新，完整字段元数据从原缓存或已归一化的 detached 原缓存复制。
- `copySlicerViewCache`：完整 remap 已复制的普通报告、独立 owned PivotChart 和 Table 列；同工作簿尚未复制的合法报告继续指向原报告。只有来源与连接集合改变时分叉 control cache，多份 copied views 共用该 batch 的新 cache。selection/timeline mirrors 从最终 cache 重建，affine/settings 仍按 view 独立复制。
- `reconcileSlicerConnectionResources`：只清理 before 中真实存在、宿主匹配且已经删除的报告引用，更新 surviving refs 的宿主和 legacy primary hints。删除 primary 不删除其他宿主上的 view；删除最后报告后保留 disconnected source/cache/selection。原本未知、宿主不匹配的引用及原本不一致的来源保留，后续 live 操作明确校验，不能经无关编辑悄悄改成合法数据。

结构编辑以同一个 range/column transform 更新完整 source cache 和 report/owned report。`cache.fields` 记录真实来源列；删除绑定列时删除对应 views/control cache。插入来源内新列使用单调 `source-field:N` 并保存 `nextFieldId`，包括没有报告的 disconnected caches；已删除的数字字段 ID 不重新解释为当前相对列。旧 cache 没有字段元数据且来源不可解析时仍保留，不能宣称其字段生命周期已经完整恢复。

工作表改名更新 source cache 与所有报告的 `source.sheetName`。剪切完整来源矩形会移动保留的 source cache 和字段列，包含无报告的 disconnected sources；cache ID 仍属于同一文档资源。原本未知的 Table 来源不会因无关 cell shift 被直接丢弃。

资源归一化只在可变的 detached/model 输入处调用主任务的 `normalizeSlicerCaches`。普通来源/字段业务错误保留原资源，非预期异常继续传播。只读查询本身不因本批而变成资源创建入口。

## 接口与修改文件

复制 helper 的 `SlicerCopyConnectionsMapping` 接收普通报告、owned chart、Table/列映射，以及 batch 共享的 `pivotCacheIds/slicerCacheIds`。结构 helper 的 `SlicerCacheStructureTransform` 接收 `sheetId`、`range`，并可选 `column(column, sourceRange)`。新增接口没有改变 Worker 协议或 Office writer 输入契约。

| 文件 | 本批变化 |
| --- | --- |
| [model-utils.ts](../../packages/office-studio/src/api/model-utils.ts) | canonical/partial/public 输入保留 `pivotCaches/slicerCaches`，输入阶段归一化共享身份 |
| [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) | 双向保留两个 cache 集合；完整深复制 Chart，避免 owned report、options、affine 与 authority 别名 |
| [workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts) | 保留导入 report 的 cache 身份、source provenance、stable field ID，归一化共享 source cache |
| [pivot_sidecar.rs](../../crates/office-parser/src/pivot_sidecar.rs) | 原生 cache part 路径派生 `office-pivot-cache:<part>`，相同 native part 的报告保留同一 ID；原生字段输出字符串 source ID |
| [sheet-clipboard-types.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-types.ts) | 资源 DTO 增加两个 cache 集合 |
| [sheet-clipboard-workbook-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-workbook-resources.ts) | clone/capture、per-tile copy/remap、cut/move、资源预算和 changed keys 保留共享 cache；fingerprint 使用同一 capture 自然覆盖连接及字段元数据 |
| [sheet-clipboard-resource-validation.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-resource-validation.ts) | cache/source-field/nextFieldId、类型化 refs、duplicate refs、owned report、来源和锁定字段一致性校验；chart type 使用既有 `isChartType`；总连接预算为 100,000 |
| [sheet-clipboard.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts)、[sheet-clipboard-commands.ts](../../packages/office-studio/src/api/sheet-clipboard-commands.ts) | candidate 发布及 recent paste restoration 包含两个 cache 集合，后者为主任务明确协调的两处 keys 修改 |
| [sheet-worksheet-copy.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-copy.ts) | 完整 source/cache/report/owned/control remap；报告字段实例 ID 保留，source ID 独立；兼容既有 drawing group mapping |
| [sheet-worksheet-lifecycle.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-lifecycle.ts)、[sheet-worksheet-references.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-references.ts) | 删除只移除真实 target；source sheet 改名包含 source cache 和 owned report |
| [sheet-axis-structure-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-resources.ts)、[sheet-cell-shift-resources.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-resources.ts) | ordinary/owned reports 和 disconnected source cache 共用结构生命周期；owned report 不创建虚构输出 cells；None 控件保留绝对 placement |

没有修改 App.vue、连接 UI 或筛选计算内核；没有覆盖前轮图片、affine 或主任务的 model/Facade/Studio 实现。

## 未运行的契约草稿

[slicer-connections-lifecycle.contract.test.mjs](../../packages/office-studio/tests/slicer-connections-lifecycle.contract.test.mjs) 与独立的 [slicer-connections-lifecycle-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-connections-lifecycle-fixture.mjs) 已落盘。与 UI/主任务的 `slicer-connections-fixture.mjs` 已隔离，避免夹具路径覆盖。

静态源码包含 **26 条 test 声明、32 个计划场景**（25 个普通声明与一个含 7 项的参数化声明）。这是声明统计，没有执行或通过结果，也不是产品功能计数。

草稿使用 Region/Amount/Date 稀疏数据、两个普通报告、可选独立 owned report、一个共享 source/control cache 和两个独立 affine views，覆盖以下意图：

- canonical/partial 快照重开与 view 深复制隔离；synthetic native parser DTO 的同 part 来源 ID。synthetic DTO 不是实际 XLSX 样本。
- 工作表副本的 report/owned/field/cache remap、self-contained source 分叉、view-only/omitted connections 缓存复用及 legacy report-local 字段迁移。
- primary/最后 target 删除、owned target 存活、source 改名、未知引用与已经不一致来源保留。
- 插入/删除来源列、稳定非数字 ID、cache-only disconnected metadata、`nextFieldId` 单调新增及真实绑定列删除。
- 区域 clipboard 完整 cache 捕获、所有目标 remap、affine 保留、完整依赖跨工作簿 copy、无报告来源 cut/move、缺依赖/不兼容 uncopied target 拒绝。
- cache owner、来源 cache、duplicate refs、非有限来源列、非法 nextFieldId、owned/linked 来源冲突、伪造字段身份等拒绝；相同范围的明确 authored cache IDs 不合并。

上述是测试计划与已写断言，**没有通过数**。连接设置后的真实筛选、独立报告输出、命令权限、撤销重做、事件、浏览器主题/焦点/IME，以及 Office 独立文件样本仍由统一验收覆盖，不能用资源模型断言代替完整闭环。

## 当前边界和剩余缺口

1. 同工作簿 source sheet 复制可能让 copied/uncopied reports 指向不同 source cache；当前拒绝整个 detached copy，要求一起复制所有连接报告，没有静默丢弃外部连接。跨工作簿 copy 要求来源与所有连接目标完整携带；缺外部来源、缺报告或无报告可映射的 disconnected cache 仍拒绝。拒绝保护不等于全部迁移能力完成。
2. 本批只补原生 Pivot reader 的 source identity，没有完成共享 native cache 写出、slicer/timeline 定义和缓存 parts、owned PivotChart、selection/filter、affine 或相关 relationships 的 XLSX 独立文件往返。[现有 xlsx writer](../../crates/office-writer/src/xlsx.rs) 仍按报告生成 cache parts；writer 的这些模型和保真契约需要后续扩展。不得把内部快照保留写成 Office 保真完成。
3. 无公开源码的 Pro cache 私有布局没有被复制。以公开类型的来源、连接、快照和副作用边界独立实现；没有执行上游演示或标准 Office 文件对比，最终交互和文件差异仍未核实。
4. 本批没有新增或修改可替换 UI 控件，连接 UI 的 `antdv-next`、ConfigProvider、主题、焦点恢复及 Windows IME 检查归 UI/主任务，均不能在这里标为已验证。
5. 资源预算沿用现有同步剪贴板边界；没有性能优化声明，也没有设备、浏览器、首屏/滚动/编辑、长任务、Worker/WASM 传输或内存测量。真实大样本保持未验收。

全部 Sheets、Docs 和 Pro 目标继续保留。以上缺口是后续实现与统一验收事项，没有改写为最终范围缩减。
