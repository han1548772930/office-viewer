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

## 后续验收

- 两个独立客户端并发修改、请求重复/乱序、相同 ID 不同内容、离线编辑和断网重连，核对确认版本、模型与资源一致。
- 已确认内容、权限、评论、任务和历史在服务重启后可恢复；恢复历史产生新版本并同步，未授权客户端不能读取或写入受限内容。
- 任务排队、运行中取消、超时、执行错误、崩溃恢复与重试均有真实结果；Office 导出使用已确认版本并独立重开验证。
- 同一数据集记录客户端输入延迟、服务响应/排队/任务时间、传输字节、内存和取消耗时。没有真实大样本和故障场景证据，保持未验收。

待实施时继续逐功能展开源码依据、数据/命令/权限、失败及恢复语义。本记录不将接口数量或架构设计算作已完成功能。
