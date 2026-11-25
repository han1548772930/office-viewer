# SHEET-22 / SHEET-24：Relation 与 WordCloud 资源闭环补充

2026-11-25。前端实现继续推进；本记录全部保持 **未验收**。没有运行测试、类型检查、构建、Cargo/WASM、浏览器、开发服务或性能基准。45 项图型目录、同名按钮和下述测试草稿均不代表全部功能完成。

## 固定依据与入口边界

- Univer 固定源：`E:\rust_test\univer`，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，v1.6.0。
- Pro 声明：[Relation Builder](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/builders/f-relation-chart-builder.d.ts)、[WordCloud Builder](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/builders/f-word-cloud-chart-builder.d.ts)。商业包仅静态参考，没有安装商业 SDK 或依赖其服务。
- 可复查运行源码：[engine-chart 归档](../../.research.local/univer-1.6.0/archives/engine-chart-1.6.0.tgz) 中 `package/lib/es/index.js`。`Hae` 先调用 `va` 识别矩阵，再选择 `Wae` / `Gae`；`va` 的 converted DTO 条件为 N 条数据、N+2 dimensions、N+1 source columns。`Gae` 校验所有端点及有限非负权重，保留零权重、重复有向边，self incident 权重只加一次；`Wae` 只累计出边权重。
- WordCloud 的公开 `setMaskImage(resourceId)` 明确接受宿主资源 ID；不能将 ID 当网络路径。公开六个专属方法为 shape/reset、mask/clear、repeat/reset；本地 `setMaskImageResource` 为宿主适配扩展，不能计入上游六项。
- Relation 的公开十二个专属方法覆盖 layout、value sizing、emphasis、circular label rotation、force、node shape 的设置与恢复。本轮没有据十二个方法存在声称它们已经通过真实交互验收。
- 新词云边栏继续使用既有图表设置上下文入口，不增加 Ribbon 根入口。轮廓来源有 3 个本地选择；形状有 10 个值，内置遮罩有 20 个 ID。候选图片引用/上传各 1 项，重复开关 1 项，字号/旋转数值 3 项，并有清除/恢复动作。选择值数量不是功能数量。固定版完整边栏位置、分组顺序、显隐效果及商业图片轮廓仍未核实，完整清点继续受[图表审计](SHEET-22-chart-type-audit.md)约束。

## 本轮实际落盘

### Relation 自动矩阵识别

- [chart-relation-data.ts](../../packages/office-studio/src/features/charts/chart-relation-data.ts) 对本地普通单元格的 row-major 输入识别 N 行、名称/分类两列、N 列权重；首列节点名须逐行匹配后 N 列标题，名称非空且唯一。显式 `mapping.relationFormat` 和显式 edge mapping 优先；自动识别不写成第二份作者配置。
- [chart-record-projection.ts](../../packages/office-studio/src/features/charts/chart-record-projection.ts) 使用同一识别结果确定 primary fields，避免先按 edge roles 验证而拒绝有效矩阵；隐藏矩阵主列仍按共同可见性规则处理。
- [chart-relation-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-relation-settings-controls.ts) 和 [ChartSettingsFields.vue](../../packages/office-studio/src/features/charts/views/ChartSettingsFields.vue) 显示自动识别结果，矩阵下隐藏 edge 字段；显式配置与恢复自动识别继续走 detached draft → `sheet.command.update-chart`。
- 本地 N+2 row-major 输入与上游 converted DTO 的 N+1 source columns 不同；匹配名称及显式 override 是本地无歧义规则。没有宣称上游内部 DTO 已被逐字复刻。矩阵数值的有限非负校验也是本地诊断边界；上游 `Wae` 对 `typeof number` 的直接接受行为仍有差异。
- 较早审计中的“所有图关系跳过零值并合并重复边”通则不适用于本项目当前 Relation：本地实现保留零值、原始空白名称和每条重复边，契约草稿单独覆盖这些差异。

### WordCloud 宿主图片资源

- [model.ts](../../packages/office-studio/src/domain/model.ts) 增加图表图片 DTO；浏览器 Blob/ImageBitmap/FileReader/URL 对象不进入作者模型。上传候选通过 `wordCloud.maskImageResource` 暂存，正式发布后移入 `metadata.chartImageResources` 并清除候选载荷，避免同时持久化两份图片。
- [chart-image-resources.ts](../../packages/office-studio/src/features/charts/chart-image-resources.ts) 集中校验 ID、MIME、data URL/base64 实际长度及 URL 类型。单张 8 MiB，注册表字符串估算 32 MiB、至多 4096 项；拒绝持久化 blob URL、裸资源 ID 作为 URL、冲突 ID 和不一致引用。预算是本地功能边界，不是堆内存实测。
- `OfficeStudio.commitChartChange` 的同一候选内容事务发布引用与注册表；remove-chart / worksheet remove 派生剩余引用并清理资源。多个图表可共享相同资源；历史快照各自持有注册表，撤销/重做恢复相应资源及引用。
- [chart-word-cloud-mask.ts](../../packages/office-studio/src/features/charts/chart-word-cloud-mask.ts) 先解析当前候选/工作簿资源，再处理内置 ID 和明确 URL/data URL；缺失 ID 返回稳定诊断，不再 `fetch(resourceId)`。响应校验 MIME、声明/实际字节并流式有界读取。完成、取消或版本过期关闭 ImageBitmap、清空 img、撤销 object URL；无 signal 的过期 decode 同样释放迟到 bitmap。
- 注册资源优先于同名内置 ID；直接 URL 的响应需要明确图片 MIME。候选上传使用内联 data URL，本轮没有外部上传服务。注册表边界已与主任务协调，后续 COMMON-12 的通用资源模型迁移仍需保持所有权/保存协议。
- [chart-statistics-service.ts](../../packages/office-studio/src/features/charts/services/chart-statistics-service.ts) 从宿主传入资源解析器，authoring 不再跳过 Relation/WordCloud 真实布局。无可排布词条时也尝试验证图片遮罩；失败不会被“只有配置保存”当成成功候选。authoring 准备实际结果不进入保存模型。
- 注册表/data URL 转 Blob、字形 raster/Worker buffers 仍有复制；预算计入源字符串、chunk/Blob 和画布估算。解码器内部峰值及 fallback img 的真实分配尚未测量，没有声称零拷贝或性能优化达标。

### UI、API 与生命周期

- [ChartWordCloudSettings.vue](../../packages/office-studio/src/features/charts/views/ChartWordCloudSettings.vue) 使用 `antdv-next` Form.Item / Select / Input / Upload / Button / Switch / InputNumber / Alert，替换原通用词云专属控件；继承工作台统一 ConfigProvider/token，没有新增原生可替换表单控件。
- Upload 的 `beforeUpload` 返回 false，读取只更新 detached draft；图片引用和载荷一起交给现有 Apply。读取期间 Apply 禁用；关闭、切换 session、草稿替换、disabled 或组件卸载会取消 FileReader，迟到结果不能更新其他草稿。Upload 与其 Button 同时禁用，文件读取同步失败也释放监听与回调。
- [ChartEditorPanel.vue](../../packages/office-studio/src/components/spreadsheet/ChartEditorPanel.vue) / [ChartInsertDialog.vue](../../packages/office-studio/src/features/charts/views/ChartInsertDialog.vue) 传递 active/session 与资源 loading；组件不承担文件解析、Office 写出或布局求解。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 的 Builder / FChart 增加本地 `setMaskImageResource`；`setMaskImage` / `clearMaskImage` 切换时清候选载荷，继续复用同一图表命令、权限和文档历史。
- Context7 的 `/antdv-next/antdv-next` Upload 文档作为 `beforeUpload=false` / disabled 行为参考，查询传入本地版本 1.8.2，返回 main 文档；不是固定 1.8.2 的运行验证。主题、焦点恢复、选择文件后的焦点和 IME 均仍需浏览器验收。

## 契约草稿与剩余验收

- [chart-relation-word-cloud.contract.test.mjs](../../packages/office-studio/tests/chart-relation-word-cloud.contract.test.mjs)：矩阵/edge 优先级、原始名称/零值/重复边、hidden primary columns、资源预算/冲突/JSON 重开、流式 MIME/字节拒绝、取消/过期 bitmap 与 FileReader 回调释放、词云 all-zero/repeat/font mapping。
- [chart-relation-word-cloud-facade.contract.test.mjs](../../packages/office-studio/tests/chart-relation-word-cloud-facade.contract.test.mjs)：同一 Apply 的引用/注册表、undo/redo、保存重开、冲突/缺失原子拒绝、源版本变化期间的迟到资源、Relation 矩阵 live commands。使用独立物化 tuple transport；[平台 fixture](../../packages/office-studio/tests/fixtures/chart-resource-platform-fixture.mjs) 的签名占位字节和 Canvas mock 仅用于隔离宿主，不能作为真实图片/求解器验收。
- [chart-word-cloud-settings.browser.test.mjs](../../packages/office-studio/tests/chart-word-cloud-settings.browser.test.mjs)：草稿计划覆盖真实 antdv-next Upload、真实 Canvas PNG/bitmap 解码、关闭/替换取消、统一主题和输入焦点。文件仅写入，没有启动测试或浏览器。
- 全部草稿未运行。仍缺真实 Worker/Rust 结果对照、上游运行/商业遮罩效果、完整侧栏入口清点、跨宿主 Docs 图表资源适配、ACL/协作/共享/异步取消全场景，以及 XLSX/DOCX/native Office 保存重开与视觉样本；不将 JSON 保留称为原生 Office 渲染等价。
- Sheets / Docs / Pro 完整工作台范围持续有效；后端后置、未验收项和本轮局部前端工作均不构成最终范围缩减。
