# SHEET-23/24 值显示方式命令、Facade 与格式投影

日期：2026-11-26。状态：实现及合同草稿已落盘，**未验收**。固定 v1.6.0 的 15 种模式、参数/算法依据及底层 1.5.2 cache 边界见 [共享引擎证据](SHEET-23-24-pivot-show-data-as.md)。历史 `percentOfParentGrandTotal` 独立保留，不计入固定 15 项。本记录只说明 API、共同提交与图表投影增量；UI 另批记录。

## 实际行为

| 行为 | 命令、模型与边界 |
| --- | --- |
| 普通透视表 | `FPivotTable.set/resetShowDataAs` → `sheet.command.set-pivot-show-data-as` → 共用 `commitPivotChange`；`set-pivot-field-setting` 可同一次提交汇总/格式/显示方式 |
| owned/linked 透视图 | `FSheetPivotChart.set/resetShowDataAs` → `edit-pivot-chart-report` 的 `setShowDataAs`；本地整份字段草稿可与外观一次 configure |
| 基准读取 | `getShowDataAsBaseItems` 使用真正 grouped 成员；基准 field ID report-local-first，目标字段 canonical-first 并消歧。返回配置 clone，不让读方修改权威报告 |
| 权限与上下文 | 普通报告新入口检查完整 source/output read 与 output edit；可传 content/history/user/host。拒绝过期、取消、无权限，不写入被忽略的配置 |
| 关联图表 | 共同 prepare 在切片器同步之后运行 `synchronizeLinkedPivotChartReports`；普通 Pivot 命令或 fan-out 均重建 linked source，已增加 revision 时不再次增加，保留视觉设置/owned report/未知 imported refs |
| 事件 | 普通 Pivot 提交捕获并发布相关 PivotChartChanged，带真实 reportRevision；观察者重入不改变后续 payload 的事务身份 |
| 错误/空值 | 错误-only 聚合保留 categories、原始错误 records 和 value-series 绑定；两种 value direction 都保留 null/error，不以是否存在数值删行或删列 |
| 格式 | owned 一次 calculatePivotReport 提供 numberFormats；linked 读输出 cells 格式，无 outputRange 走同一 data-only 计算。hierarchy/heatmap 透传 formattedRecords，与普通源保持同一请求口径 |
| 保存/历史 | field.showDataAs、base refs、原始聚合源和 chart 作者设置进入既有快照与文档历史；重开仍重新从真实报告读取。内部快照不等同 Office 原生支持 |

模块为 Pivot services/Chart data projection 与应用命令边界，Vue 不实现聚合。新增变换复用原聚合 callback，不建立第二权威引擎。本批未实测大型数据，也未完成 retained Rust/Worker 计算，不能声称优化。

## 本批文件与计划验证

- [pivot-report-permissions.ts](../../packages/office-studio/src/features/pivot/services/pivot-report-permissions.ts)：普通报告读/编辑检查。
- [pivot-chart-report-lifecycle.ts](../../packages/office-studio/src/features/pivot/services/pivot-chart-report-lifecycle.ts)：共同 linked source/revision 投影。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)、[facade.ts](../../packages/office-studio/src/api/facade.ts)、[index.ts](../../packages/office-studio/src/index.ts)：新 command/field setting/context/读写/导出片段。
- [chart-pivot-data.ts](../../packages/office-studio/src/features/charts/chart-pivot-data.ts)、[chart-source-compute.ts](../../packages/office-studio/src/features/charts/chart-source-compute.ts)：error-only/null 和 formattedRecords 片段。
- [pivot-show-data-as-facade.contract.test.mjs](../../packages/office-studio/tests/pivot-show-data-as-facade.contract.test.mjs)：**13 条声明 / 24 个计划场景**。真实 OfficeStudio、PivotChart source cells、Facade/history；owned/linked/PivotTable、汇总+视觉一次提交、累积/排名/基准、reset/no-op、非法/删除依赖、直接报告事件、取消/过期/权限、error-only 双方向、百分比格式与 lifecycle 去重。

引擎独立草稿为 [34 条声明 / 65 个计划用例](SHEET-23-24-pivot-show-data-as.md)，不重复计入上述 13/24。只做静态回读和声明/参数清点，所有草稿未运行。新 UI 无新增顶部根，15 个模式属于字段设置枚举，不能与按钮数量相加。

## 保留缺口

完整成员 DTO/空白基准项、Office 边缘语义、所有 Pro Builder/Facade/controls 与字段/缓存迁移仍需继续核对。真实主题/焦点/Windows IME、Canvas 两后端、快照与文件重开、权限边界、Worker/Rust/WASM 性能/内存/取消均未验收；不把公开方法存在计为完成。保存与服务协议需求见 [后端需求](../BACKEND_REQUIREMENTS.md)。
