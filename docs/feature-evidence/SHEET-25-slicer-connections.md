# SHEET-25：共享切片器多目标连接

记录日期：2026-11-26。状态：**实现和测试草稿已落盘，未验收**。

本批接续 [affine 与 Ant 画布控件](SHEET-25-slicer-timeline-affine.md)，补共享来源缓存、替换报告连接、多个报告筛选、Facade 与同一次内容历史。UI 和资源生命周期分别见 [连接界面](SHEET-25-slicer-connections-ui.md)、[复制与结构生命周期](SHEET-25-slicer-connections-lifecycle.md)。整个 Sheets / Docs / Pro 范围继续保留。

## 固定依据与入口清点

OSS 参考为 `E:/rust_test/univer` 的 v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据为 `.research.local/univer-1.6.0/packages/` 的发布类型和可取得的 ES schema；没有安装或执行商业 SDK。`sheets-slicer/lib/es/index.js` 引用的内部 chunk 未取得，不能把类型反推写成已阅读其内部算法。

| 具体来源 | 已确认的公开契约与本地实现 |
| --- | --- |
| `sheets-slicer/lib/types/facade/f-slicer.d.ts` 的 `getConnections/getConnectionCandidates/setConnections`，约 250–300 行 | getter 返回防御性共享快照；Table 来源为 undefined/空候选；同 workbook/cache/source field 的候选；替换而非追加；linked chart 归一为所属 PivotTable，owned chart 独立；无效/权限/不可用为 false，未知命令或监听器异常 reject |
| 同包 `types.d.ts` 的 `SheetSlicerSourceSpec/ISheetSlicerConnections/ISlicerConnectionCandidate/SlicerIntentType.ChangeConnections` | 来源身份与报告集合分离；本地 `SlicerCacheModel` 持有来源、connections、selection、timelineRange，view 只保存宿主与 presentation/affine |
| 同包 source、transaction、model 服务类型及 `models/sheets-slicer.model.d.ts` | 多模型准备后发布、共享 cache 与 retained source；本地在现有工作簿候选事务中准备全部报告与输出，失败不发布部分结果 |
| `sheets-pivot-chart/lib/types/types.d.ts` 的 Range/PivotTable source | Range 生成独立 owned report，PivotTable source 为 linked；本批补 owned 模型/计算/连接，完整 owned 创建与字段编辑入口由后续批继续闭环 |
| `sheets-slicer-ui/lib/types/views/SlicerConnectionsDialog.d.ts`、UI service、open-dialog operation、menu 类型及 ES schema | 设置与连接使用不同 operation；格式 tab 的 styles/options 两组、四个条件根按固定顺序接入，本地使用 Vue/antdv-next |

本批常规根增量为 **0**：Sheets 保持 **71（30/7/11/17/6）**，Docs 保持 **52（33/16/3）**。以下上下文声明与常规根分开统计，也不把同一入口的 More 呈现再加一次。

| 分类、顺序与布局 | 入口与真实命令 |
| --- | --- |
| Data/分析既有 2 根；classic 按透视表、透视图、切片器、时间线排列 | `sheet.operation.open-slicer` / `sheet.operation.open-timeline` → `sheet.command.add-slicer`；本批沿用位置 |
| 1 个格式上下文 tab，order 100；styles order 0 | 快速样式 `sheet.menu.slicer-quick-style`，r1c1 跨两行 → `sheet.command.update-slicer-presentation` |
| options order 1，项 order 0/1/2 | 切片器设置、时间线设置 r1c1 → `sheet.operation.open-slicer-settings-dialog`；报表连接 r2c1 → `sheet.operation.open-slicer-connections-dialog` → `sheet.command.set-slicer-connections` |
| Canvas 右键与浮条连接位置 | 复用同一 open-dialog operation，不增加独立内容命令或常规根 |

格式根共 **4 个静态声明**；两种设置项按 view kind 条件互斥，Table 来源隐藏连接。焦点、宿主、来源和权限控制可见/可执行状态，不能由静态计数推导运行态同时可见数或功能通过数。

## 模型、计算与提交

- [model.ts](../../packages/office-studio/src/domain/model.ts) 增加 `PivotCacheModel`、`SlicerCacheModel`、分类型连接引用与 `ChartModel.ownedPivot`。明确 authored/imported cache ID 不因同一 range 被合并；旧模型仅在没有明确身份时按完整来源 tuple 迁移。缓存借用源 cells，只保存来源和字段元数据，字段的 canonical `sourceFieldId` 与 report-local `id` 分开。新增列使用单调 `source-field:N/nextFieldId`。
- [slicer-connections.ts](../../packages/office-studio/src/features/slicer/slicer-connections.ts) 提供读回、候选、替换 planner、linked owner 去重和资源归一化。零连接保留来源、字段元数据与选中项，后续仍能读成员并重新连接；无效保留资源写入 `metadata.slicerResourceDiagnostics`，live 操作明确拒绝，不静默删除或重绑下一列。
- [slicer-report-filter.ts](../../packages/office-studio/src/features/slicer/slicer-report-filter.ts) 借用 source cells，仅复制变化报告和输出 sheet 的 cells map；先清全部旧输出，再准备全部新输出，检查冲突后统一发布。owned report 使用同一纯 Pivot 计算，不写 worksheet cells。普通输出冲突转为业务拒绝，未知异常继续抛出。
- [slicer-engine.ts](../../packages/office-studio/src/features/slicer/slicer-engine.ts) 把 items/timeline/clear/连接落到共享 cache；同来源多个 views 同步 selection/revision。外部 Pivot/Table 字段筛选也更新 cache 并联动其余报告；同一次编辑为一个 cache 指定冲突选择时拒绝整个候选。
- [slicer-permissions.ts](../../packages/office-studio/src/features/slicer/slicer-permissions.ts) 检查源范围读取、全部共享 view、普通报告前后输出和 owned chart footprint 的编辑权限；UI 禁用和命令验证共用这些来源/范围约束。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) 注册连接命令并使用现有候选事务和内容历史。校验 unit/revision/user/history/工作台宿主及事件重入；普通无效目标返回 false、满足的 no-op 返回 true、取消或监听器异常继续 reject。提交后通知变化的全部共享 view 和 PivotTable，隔离 observer payload 与异常。删除真实目标后先 reconcile，再 normalize/sync，不因 primary 被删除清掉存活连接。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 增加 `getCacheId/getConnections/getConnectionCandidates/setConnections`；插入和工作台 mutation 接显式 revision/host。默认绑定 Facade 可操作 inactive unit，工作台上下文拒绝已切换宿主。

图表数据路径同时修改 `chart-pivot-data.ts`、`chart-source-compute.ts`、`chart-source-diagnostic.ts`、`chart-engine.ts`、`chart-settings-draft.ts`、`chart-settings-data.ts` 与 `services/chart-statistics-service.ts`：owned 来源进入同一源读取、签名与权限流程；全部成员过滤掉时输出空记录及字段元数据，避免旧分类继续显示。Pivot 计算公共接口及资源/协议导出见 `pivot-engine.ts`、`index.ts`。

模块按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)分层。未新增 Worker/WASM 协议、浏览器绑定或商业依赖；报告计算仍在现有 TypeScript 路径，大样本迁移、取消、复制/内存计量和真实阈值须遵守[协议 2 与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。没有性能优化或端到端零拷贝结论。

## 测试草稿与 UI 边界

用户要求先实现、后统一测试，本批只读写和人工静态回读。**没有执行测试、类型检查、前端/Rust/WASM 构建、Cargo、浏览器、dev、基准或 `git diff --check`。**

| 主任务新增草稿 | 静态声明 | 计划场景 | 断言意图 |
| --- | ---: | ---: | --- |
| [slicer-connections.contract.test.mjs](../../packages/office-studio/tests/slicer-connections.contract.test.mjs) | 13 | 18 | 候选/防御性快照、6 类非法目标、linked 去重、owned local ID、替换、共享联动、timeline、零连接恢复、detach、冲突原子性、owned 无输出 cells、空图表与稳定字段 |
| [slicer-connections-facade.contract.test.mjs](../../packages/office-studio/tests/slicer-connections-facade.contract.test.mjs) | 12 | 13 | 一次历史/undo/redo/内部保存重开、no-op/false、权限/冲突、2 类监听器异常、取消、observer、过期/宿主、外部筛选与 primary 删除 |
| 合计 | **25** | **31** | **全部未运行，没有通过数** |

共享 [slicer-connections-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs) 只用 North/South 与日期/金额的合成稀疏数据，包含两个兼容 report、同 range 不同 authored cache、linked/owned chart、共享 region views、独立 timeline 和 Table slicer。UI 与 lifecycle 的草稿分别在其证据中统计，旧 affine 98/182、绘图 82/98、Sheets 图片 66、Docs 图片 57 保持独立，不能合并成已覆盖功能数。

本批可替换交互控件使用 `antdv-next` 的 Form、Modal、Checkbox、Button、Select 等，受统一 ConfigProvider/theme token 管理；Canvas 内核仍管理绘制/选区/拖拽。没有新增原生文本/IME 桥。真实明暗主题、焦点恢复、弹层键盘、输入法和两种渲染后端仍待统一浏览器验收。

## 未核实项与后续缺口

1. 当前 detach 选择保留报告最后 field filter，之后独立编辑；重新连接应用 cache 当前选择。取得的固定类型没有明确 detach 后过滤器处理细节，Office/上游实际运行行为待核实，不能标成功能等价证据。
2. 完整交叉筛选成员可用状态、cache options、全部日期条件/粒度、所有 Pro Facade/事件/移动入口仍须逐项核对。owned 创建和字段编辑不是本批模型/计算支持的完成证明；后续 [SHEET-24 独立/关联创建与字段 API](SHEET-24-owned-pivot-charts.md)、[字段引擎](SHEET-24-pivot-chart-engine.md)及 [创建/字段 UI](SHEET-24-pivot-chart-ui.md)已落盘，均未验收。
3. 复制缺失来源/目标或部分外部依赖时目前拒绝整个 detached paste/copy；跨 workbook disconnected 等依赖闭环仍待补齐，详见 lifecycle 证据。拒绝保护不等于完整迁移功能。
4. 内部 JSON 保留资源与真实 XLSX 共用 Pivot cache、slicer/timeline parts、独立 owned report、稳定字段/缓存身份及关系导出分别验收；仅 reader 身份字段补齐，Office 写出与独立文件重开仍缺，需求登记在[后端/IO 清单](../BACKEND_REQUIREMENTS.md)。
5. 真实 UI、复杂文件、性能/内存/取消和两后端均未验收。全部能力目标保留，不以四个上下文根、声明数量或本批静态实现宣称 SHEET-25/Pro 已完成。
