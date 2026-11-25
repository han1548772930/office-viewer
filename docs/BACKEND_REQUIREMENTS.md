# 后端需求与后续实施范围

记录日期：2026-11-25。状态：待开发、未验收。当前顺序是先完成前端对齐；本文件保存后端需求，不表示服务已经实现。

## 用户要求与项目归属

- 必要的后端能力使用 **ntex** 作为 Rust Web 框架，可按实际需要添加持久化、认证、任务调度等依赖。
- 按 Univer **v1.6.0** 的源码、公开发布类型和可观察行为确认服务在做什么，再在本项目独立实现；不依赖 Univer Server、Exchange 或商业 SDK，不复制私有协议。
- 服务放在 `E:\rust_test\office-viewer` 内，属于用户指定的 `E:\rust_test` 工作区域。建议入口为 `services/office-server`，实际创建时与现有 Rust crate、构建脚本和锁文件协调。
- 后端暂缓，继续优先实现 Sheets / Docs 前端。完整交付仍包含下面的服务范围；暂缓不等于删减目标。
- 测试、类型检查、构建、Rust/WASM、浏览器及性能测量在统一验证阶段执行，可随实现编写测试和样本草稿。

执行依据：[完整工作台架构](../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)、[功能清单](UNIVER_FEATURE_PARITY.md)、[计算与平台标准](RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。

## 固定源码与公开契约依据

只读参考仓库为 `E:\rust_test\univer`，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下是接口/客户端可观察契约，不能据此声称掌握闭源服务端实现。

| 来源 | 需对照的行为 |
| --- | --- |
| `packages/protocol/src/ts/universer/v1/snapshot.ts` | `ISaveChangesetRequest/Response`、`IFetchMissingChangesetsRequest/Response`、`IGetLatestCsReqIdBySidRequest/Response`：提交确认、缺失变更补取与请求身份 |
| 同文件 | `IEnsureSnapshotRequest/Response`、`ISaveSnapshotRequest/Response`、`IGetUnitOnRevRequest`：快照与特定版本内容 |
| `packages/protocol/src/ts/universer/v1/history.ts` | 历史版本相关公开数据与请求；实现前逐项展开并记录差异 |
| `.research.local/univer-1.6.0/packages/collaboration-client-ui/lib/types/services/local-cache/web-local-cache.service.d.ts` | 客户端本地缓存、离线与重连边界 |
| `.research.local/univer-1.6.0/packages/edit-history-ui/lib/types/services/history-ui.service.d.ts` | 版本列表、读取、恢复与 UI 状态边界 |
| 本地缓存中的 collaboration / collaboration-client / edit-history / exchange-client 发布包，以及[公开协作](https://docs.univer.ai/zh-CN/overview/features/shared/collaboration)、[历史](https://docs.univer.ai/zh-CN/overview/features/shared/history)、[权限](https://docs.univer.ai/zh-CN/overview/features/shared/permissions)规格 | 实施前核实具体版本、可用源码/类型、任务请求及状态；未核实部分保持待核实 |

后端 HTTP、WebSocket 和任务 DTO 使用本项目的独立版本化协议。上游方法名用来定位语义，不要求复制其 URL、二进制编码或内部存储。

## 功能范围与真实闭环

| 功能 ID | 服务责任 | 必须形成的闭环 |
| --- | --- | --- |
| COMMON-06 / SHEET-30 | 身份与权限 | 文档、表、范围、对象及 Docs story/控件的查看、编辑、复制、评论、导出、打印分别判断；每次服务端写入重验权限；权限变更撤销非法待提交操作 |
| COMMON-07 | 实时协作与同步 | 持久 changeset、确认版本、去重、补取、重连和离线队列；远端选区/presence 独立于内容日志；多个真实独立客户端一致 |
| COMMON-08 | 持久历史版本 | 快照加增量日志、版本列表、指定版本读取；服务重启后可查询和重放；不使用本地撤销栈代替 |
| COMMON-09 | 语义对比与恢复 | 内容、样式、结构、对象与资源差异；历史视图只读；恢复生成新的事务和版本，保留原日志并广播 |
| SHEET-17 / DOC-10 | 持久线程评论 | 评论身份、锚点、回复、编辑、解决/恢复、权限与用户；锚点随结构或文本修改变换；重启和跨客户端保留 |
| COMMON-12 / COMMON-10 | 媒体和嵌入关系 | 二进制资源、关系、引用计数、权限；父子文档身份/引用或快照模式、保存和删除协调；复制/恢复不丢资源 |
| SHEET-29 / DOC-22 / SHEET-28 / DOC-21 | 同步作业与异步任务 | 导入导出、已确认版本快照、Office 转换和需要服务端的打印/PDF任务有真实排队、执行、取消、失败及结果；复用本项目解析/写出和同源布局，不接入 Exchange |

任务必须实际完成所请求的工作。创建任务后仅返回固定成功状态、空结果、定时器进度或提示消息，不算实现。

## 协议、持久化和执行要求

1. 内容事务至少带 `protocolVersion/unitId/clientId/transactionId/baseServerRevision`，按用户身份和文档绑定验证。重复提交返回同一次确认，不重复修改内容；重复 ID 配不同内容必须拒绝。
2. 服务端按文档排序和原子确认。客户端对未确认 mutation 做确定性 OT 重基；文本、行列、范围/锚点、对象及资源分别定义变换规则。若某域改用 CRDT，须先记录域、ID、冲突、历史和迁移协议。
3. 快照、内容日志、权限、评论、媒体关系、版本和去重索引持久化。确认响应只能发生在相关数据可靠写入之后；崩溃重启不能重复应用或丢失已确认内容。
4. 补取支持确认版本和分页/数量预算。日志裁剪后提供可重建的快照与后续增量，不让客户端无限重试已经不可用的版本。
5. 每个任务记录 `jobId/unitId/sourceRevision/kind/status`、实际阶段/处理量、取消状态、结果引用和结构化错误。状态至少区分排队、运行、成功、失败、取消；取消/失败不发布部分文档结果。
6. 同步作业读取明确的已确认版本；作业启动、排队、重试和发布都核对权限与目标版本。过期结果不得覆盖用户的新编辑。重试和恢复需防止重复写入。
7. CPU 重计算不阻塞 ntex 的请求处理。采用有界执行器、队列、内存/文件大小预算与超时；业务计算可复用本项目 Rust 纯内核，浏览器 WASM 不承担后端网络职责。
8. 本地 Office 文件解析和编辑可以离线运行。前端通过服务接口适配接入协作/持久版本；未配置服务时仍使用真实本地模型，不以空回调伪装联网功能。
9. 配置至少覆盖监听地址、持久化目录/数据库、认证来源、资源预算与任务并发。提供项目内启动说明和恢复方案；正式依赖和锁文件随实现正常生成，本次不安装或启动服务。

## 2026-11-25 前端内容协议增量

以下需求续接[富文本字体与编辑](feature-evidence/SHEET-04-05-rich-text-font-editing.md)、[默认样式来源](feature-evidence/SHEET-04-default-style-inheritance.md)、[Docs 结构块](feature-evidence/DOC-13-14-16-17-structured-blocks.md)、[Canvas/辅助宿主](feature-evidence/DOC-13-14-16-17-canvas-interactions.md)和[Code/段落入口](feature-evidence/DOC-13-14-16-17-code-popmenus.md)，服务和原生文件闭环仍待开发/验收。

| 功能 ID | 持久内容与原生文件要求 | 临时状态与计算边界 |
| --- | --- | --- |
| SHEET-04 / SHEET-05 | 保存最终单元格值、类型、数字格式及全部 rich runs；保留来源身份、嵌套字体/公开别名、显式 false、颜色、链接、空标记和未知扩展。XLSX 解析/写出复用本项目链路，扩展 rich run 与主题字体映射，实际 Excel 保存/重开不得扁平化或丢作者字段。 | input trace、typing 偏好、native/Canvas selection、滚动和焦点只属于编辑视图，不进入持久内容或历史版本。文本范围 changeset 明确定义 UTF-16 offsets、来源版本、运行/锚点变换及权限，过期草稿不能覆盖新内容。 |
| SHEET-04 / SHEET-03 / SHEET-08 | 保存 workbook/worksheet 默认、样式目录、行列点/spans/priority，以及 cell authored patch、reference、reset mask、theme 和物化标记。完整 `style` 不能作为唯一权威或覆盖稀疏作者来源；目录更新、行列变换、重开和恢复不将旧缓存冻成 direct。[公开轴](feature-evidence/SHEET-03-public-axis-import.md)已在前端映射 h/ah/ia/w/hd，原始对象及未知扩展存档随结构变换保留；17 个草稿未运行。服务端当前值与来源存档分开，完整 auto-height/Office 样式与单位映射仍待实施/验收；[短字段读回](feature-evidence/SHEET-04-public-style-readback.md)仅为边界投影。 | 剪贴板/格式刷的已解析区域快照与 pending cut 属于会话来源，未粘贴的 snapshot、笔状态和 view 不进入持久内容。成功粘贴/刷格式/填充提交作者结果；NO_FORMAT 保留目标来源。超同步预算继承边框/填充由分块 Worker/Rust/WASM 和取消协议落实，100,000 预算不是验收指标。 |
| DOC-13 / DOC-14 / DOC-16 / DOC-17 | 保存 Code/Column/Quote/Callout 的真实块树、段落运行、来源 body 与配置。Column 的 horizontalPadding/minHeight 按 CSS px 保持模型值，默认 8/72；公开 body NumberUnit 在未修改时保留作者单位，作者修改后输出 `u=3`。内部快照保留不等于 DOCX/ODT 原生结构保真，需扩展解析/写出及来源关系并验证 Office 样本。 | Code block/host 的水平滚动、结构浮动入口及拖拽预览属于视图，不进内容历史；实际结构变更通过同一内容命令、权限与原子事务。 |
| DOC-13 / DOC-16 / DOC-17 | Code 字体保存在真实段落/run，language/wrap/showLineNumbers/tabSize 保留作者值及未知扩展；结构类型转换保留容器/子段落 ID、内容/资源和 `metadata.documentStructuredTypeSources` 可恢复配置。changeset、历史与原生文件映射不得仅保存七类高亮颜色或丢转换来源。 | Prism Token/UTF-16 Canvas 分区是派生视图，字体登记、Popover/Drawer、More 草稿及 target/revision/focus 不进内容日志。表后/辅助 host/嵌套内容变更使用明确位置及同一锁/权限检查；[Quote 分页切片](feature-evidence/DOC-16-quote-pagination.md)已接共享布局，复杂对象/表格分页仍待前端补齐和统一验证。 |
| SHEET-22 / SHEET-24 | 图表作者 DTO 保存标题/副标题样式、标题 manualLayout/rotate、绘图区相对完整外框的 manual rectangle、图例配置、类型特有配置与资源 ID/二进制关系；恢复历史和复制时资源仍可解析。标题/plot 拖动释放生成一次作者事务；真实 XLSX/DOCX/DrawingML 可编辑写回按图型逐项补，不以内部 JSON 为等价证据。 | 图例 selected/page、标题/plot hover/selection/preview、Sankey 临时拖动位置、Relation viewport 等 renderer view 不进入内容日志。target bounds/automatic seed metadata 仅供同源命中，不序列化成作者 option；故事/宿主和过期 revision 在发布时重验。详见[标题](feature-evidence/SHEET-22-title-direct-editing.md)和[绘图区](feature-evidence/SHEET-22-plot-area-direct-editing.md)。 |

富文本测量/布局、批量格式、图表变换及索引等重计算仍归项目纯内核与 Worker/Rust/WASM，遵守类型化控制 DTO、transferable、lease、memory.grow 视图重建、复制字节统计和取消契约；不放进 ntex 网络处理层。服务端作业需要同源计算时复用纯 Rust 内核并由有界执行器调度。当前不启动服务或执行统一验证。

## 2026-11-26 前端内容协议增量

| 功能 ID | 持久内容与原生文件要求 | 临时状态与计算边界 |
| --- | --- | --- |
| SHEET-04 范围主题 | 保存 workbook 自定义九区段目录、worksheet 规则 ID/name/物理范围、可选逻辑 styleRange 与未知扩展；逻辑范围可含负起点且必须包含物理成员。目录注销保留 dangling 规则，重注册可恢复。九区段保存/应用形成一个内容 changeset，服务端重验所有同名引用区域、expected revision/history，不拆成先注册后应用。行列/表复制、局部移位及 cut/all 使用规则变换/ID 重映射。公开 `SHEET_RANGE_THEME_MODEL_PLUGIN` 资源已前端导入，完整输出与 XLSX/theme/tint 仍须扩展本项目解析/写出，不以 JSON 代替 Office。 | Ant 草稿/区段/边框位置/错误/pending、区域快照、剪贴板与回焦 generation 不进入内容日志。set/remove 和目录变更通过共享事务，移除权限覆盖整条规则；局部移位已接[共享候选/确认](feature-evidence/SHEET-02-cell-shift-transactions.md)。大量规则索引/样式投影归 Worker/Rust/WASM；复杂组合仍待前端补齐与验收，详见[基础 46 个草稿](feature-evidence/SHEET-04-range-themes.md)、[九区段编辑 45 个新草稿](feature-evidence/SHEET-04-range-theme-editor.md)及[局部结构 32 个草稿](feature-evidence/SHEET-04-range-theme-structure.md)。 |
| DOC-02 / DOC-16 分页 | 保存四项可选段落 boolean（keepLines/keepNext/pageBreakBefore/widowControl）、默认/作者来源及 Quote 自动生成前的恢复信息；公共 body 输出 0/1，保持未修改来源扩展。DOCX 必须映射真实 w:keepLines/w:keepNext/w:pageBreakBefore/w:widowControl，Quote 插入/退出/转换保留作者覆盖。 | 实际行测量、note floor、分页和辅助 host 都是前端布局派生结果，不把预测高度、page/column cursor 或 UI draft 当作者内容。共享 paragraph 命令和 Facade 接线已落盘，37 个草稿未运行；复杂对象/表格规则和原生文件仍待实施/验收，见[Quote 分页](feature-evidence/DOC-16-quote-pagination.md)。 |
| SHEET-22 / SHEET-24 对象操作 | 同 workbook 粘贴保存新图表 ID、原 source/pivot 绑定、作者格式及资源引用，一次内容历史；删除 prune 无引用资源。跨 workbook 依赖导入、真实 chart-sheet、通用/分组 drawing 与 Office 外部 clipboard 仍是前端和文件适配缺口。服务端只承担内容确认/权限/持久化，不兜底缺失的客户端操作。 | 编辑打开面板、复制、PNG 导出不增加内容历史；clipboard ID/ticket/session、Ant 菜单与借用/owned Canvas 不进入快照。真实 Canvas PNG 编码的像素复制、预算/释放与迟到回调已有接线，34 个草稿未运行，无性能通过结论；详见[图表对象操作](feature-evidence/SHEET-22-chart-object-context-actions.md)。 |
| SHEET-02 / SHEET-03 / SHEET-08 局部移位 | 保存实际移动格的作者来源、公式/名称、CF/DV/主题片段、保护、表格/筛选、图表/透视绑定与切片器锚点；一次内容历史同时记录选择前后。后续 changeset 定义横纵带结构 edit、旧版本/资源变换和重基，服务端重新检查所有实际来源/目标权限；不能只保存按钮 ID。XLSX 复用本项目 parser/writer，补通用 drawing/media 和独立 Office 重开。 | 确认 request、baseRevision/history generation、UI host identity/signature、输入/回焦和 pending 不进内容日志。四核心+四确认命令、Facade、候选/权限/验证和 observer guard 已落盘，56 个草稿未运行；同步矩形预算不是性能证据，大范围资源变换/索引/取消归 Worker/Rust/WASM，详见[移位事务](feature-evidence/SHEET-02-cell-shift-transactions.md)。 |
| SHEET-22 / SHEET-24 面板 route | 仅应用后的 chart 作者 patch 与源/资源绑定进入内容模型、历史和服务日志；同一图表反复打开面板保持未提交 draft，不能把导航误记为作者变更。 | tab/section/axisId/seriesId/titleKey、导航 key、Collapse/Tabs、focus/composition/session 属于 UI，不同步为 chart 内容。P17 route 接真实 Ant Drawer/控件，33 个草稿未运行；全部 element 命中/直接编辑、外部 drawing/跨 workbook 等仍归前端与文件适配，详见[面板路由](feature-evidence/SHEET-22-chart-edit-panel-routing.md)。 |
| SHEET-22 线型/空值 | 保存 `area.lineStyle` 的 line/smooth/step 与 invalidValueStrategy；专用写入/reset 清除旧 smooth/空值覆盖，保留其他作者字段，一次更新进入内容历史。XLSX 原生线型映射、Step 诊断及真实 Excel 重开继续扩展本项目 parser/writer，manifest 或图片不替代可编辑文件。 | 采样几何、effective legacy 状态和渲染缓存属于派生数据；导航和未应用的 Select draft 不进入内容日志。P18 已接二维/3D 与公开 Builder/FChart，34 个新草稿未运行；曲率/组合最终效果、两后端及大样本 Worker/WASM、取消/内存/复制计量未验收，详见[线型证据](feature-evidence/SHEET-22-line-and-area-styles.md)。 |
| DOC-15 / COMMON-03/05 / DOC-01/03/12 LaTeX | 保存稳定 range/run ID、原始 latex source、字符/段落来源及 live DOC_LATEX_PLUGIN 配置；named resource 优先，删除不从旧 nested wrapper 复活。公开段落 text/offset 读真实 dataStream source，header/footer segment ID 复用保存分配及 aliases。公开插件仅 inline，本地 block 另用本项目扩展。协作/历史保留 source↔atom 坐标映射及锁/权限；模板应用后仍保存原 source，真实 DOCX/OMML 读写继续扩展本项目链路，内部 JSON 不代替 Office 验收。 | Canvas/数学 AST/布局、71 模板目录、unit editor owner/句柄、Ant source draft/错误、上下文 tab/selection/session/revision 和迟到回焦不进内容日志。四 public 命令、FDocumentLatex 与模板 operation/Facade 绑定使用同一 canonical 内容路径；原 43 个及新 47 条声明（计划 117 用例）未运行。其他段落/浮动/移动入口、完整 TeX/OMML、Worker/Rust/WASM 布局/取消/复制计量和真实 IME/两后端/性能仍有缺口，详见[LaTeX 编辑](feature-evidence/DOC-15-latex-editing.md)与[模板/上下文](feature-evidence/DOC-15-latex-templates-ribbon.md)。 |

上述增量不改变后端暂缓和未来 ntex 的安排，完整前端/文件/服务验收目标保留。

## 后续验收

- 两个独立客户端并发修改、请求重复/乱序、相同 ID 不同内容、离线编辑和断网重连，核对确认版本、模型与资源一致。
- 已确认内容、权限、评论、任务和历史在服务重启后可恢复；恢复历史产生新版本并同步，未授权客户端不能读取或写入受限内容。
- 任务排队、运行中取消、超时、执行错误、崩溃恢复与重试均有真实结果；Office 导出使用已确认版本并独立重开验证。
- 同一数据集记录客户端输入延迟、服务响应/排队/任务时间、传输字节、内存和取消耗时。没有真实大样本和故障场景证据，保持未验收。

待实施时继续逐功能展开源码依据、数据/命令/权限、失败及恢复语义。本记录不将接口数量或架构设计算作已完成功能。
