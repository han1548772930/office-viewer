# SHEET-22：图表对象上下文命令与剪贴板 / PNG 导出

状态：**前端接线与测试草稿已落盘，未验收**。本文件记录 P16，不重复计算 P14 标题 / 副标题或 P15 绘图区入口。四个默认上下文根已接到实际命令；可选面板 route 已由后续 [P17](SHEET-22-chart-edit-panel-routing.md) 继续接入真实表单，仍未验收。跨工作簿导入、图表工作表和外部 Office 绘图剪贴板仍有前端缺口。

## 固定来源与入口分类

固定对照版本：Univer **v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**。基础仓库为 `E:\rust_test\univer`；本批 Pro 可读证据来自项目内固定版本包的 ES / 类型快照，未安装或运行商业 SDK。

- [sheets-chart-ui ES](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/es/index.js)：`Ii` / `getChartContextMenuTarget`（约字符偏移 19426）、`Ri`（约 20119）、`zi`（约 20382）、`Bi`（约 20529）、`Vi`（约 20795）；`za/Ba/Va/Ha` 与 `Ua` 菜单 schema（约 93196）。偏移为固定 ES 的查阅定位，不是源码行号。
- [focused-chart.command.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/commands/commands/focused-chart.command.d.ts)：四个公开命令及 `route?: IChartEditPanelRoute`。
- [chart-context-menu.d.ts](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/menu/chart-context-menu.d.ts) 与 [menu schema](../../.research.local/univer-1.6.0/packages/sheets-chart-ui/lib/types/menu/schema.d.ts)：菜单入口声明。
- [chart-edit-panel-route.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/routing/chart-edit-panel-route.d.ts) 与 [float-menu/types.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/float-menu/types.d.ts)：`tab/section/axisId/seriesId/titleKey` 路由结构与枚举。
- 复制链追踪：同一 ES 的 `Fi.copyChartFromContextMenu`（约 10438）临时锁定实际目标，调用通用复制命令并在 `finally` 清除目标；clipboard hook 保存单个未分组图表 / drawing 快照。粘贴校验 `copyId`、默认粘贴与原 drawing 未变化，`_generateChartCopy` 区分 GRID / CHART 目标，`_generateCopyPasteMutations` 创建新 ID，并配置 add/remove drawing、chart、style 的 redo/undo。

| 固定入口 | 命令 ID | 分组 / 顺序 | 显隐及状态 | 本地接线 |
| --- | --- | --- | --- | --- |
| 编辑图表 | `sheet.command.chart.open-focused-panel` | DRAWING / OTHERS，0 | 无图表目标隐藏；面板激活状态来自 `sheet.chart.edit.panel` | 当前真实图表 → 既有 Ant Drawer 与作者命令；P16 的显式 route 缺口由 P17 继续接入，见独立证据 |
| 复制图表 | `sheet.command.chart.copy-focused-chart` | DRAWING / OTHERS，1 | 无图表目标隐藏 | 权限与目标快照 → 内部图表剪贴板 + 浏览器 HTML / 文本来源标记；不即时增加对象 |
| 导出图片 | `sheet.command.chart.export-focused-image` | DRAWING / OTHERS，3 | 无图表目标隐藏 | 实际 ECharts Canvas → 独立快照 → PNG Blob → 下载 anchor |
| 删除图表 | `sheet.command.chart.delete-focused-chart` | DRAWING / OTHERS，4 | 无图表目标隐藏；图表工作表主图隐藏 | 实际目标 → 一次内容提交 / 资源清理；主图按固定 ES 返回 `false` |

分类计数：**固定上下文命令根 4、常驻 Ribbon 根 0、新增插件扩展根 0**。本地“更多”按钮复用这 4 个根，是另一种呈现，不计成 4 个新增功能。F2、Ctrl/Cmd+C、Ctrl/Cmd+V，以及本地 `sheet.command.chart.paste-copied-chart` 是调用方式或粘贴接线，不冒充固定上下文 schema 的第五项。此处计数不构成功能完成证明。

上游目标选择拒绝多焦点 drawing，接受 chart；图表工作表无焦点时可回退主图。本地当前只有单图表选中状态，并通过版本化 `officeChartSheet = { version: 1, primaryChartId }` metadata adapter 实现主图回退 / 禁删；没有由此补齐真正的图表工作表创建、导入和布局。

## 本地实现、模块与数据边界

本批涉及 **16 个文件**：6 个新增图表模块、4 个既有图表接线文件、5 个测试 / fixture 文件及本证据。中央清单由主任务统一更新。

- [chart-object-commands.ts](../../packages/office-studio/src/features/charts/chart-object-commands.ts)：固定命令、菜单顺序、目标别名 / 版本 / 快照校验、主图 adapter、共用能力判断与弱引用 UI 宿主注册。
- [chart-object-clipboard.ts](../../packages/office-studio/src/features/charts/chart-object-clipboard.ts) 与 [chart-object-clipboard-store.ts](../../packages/office-studio/src/features/charts/chart-object-clipboard-store.ts)：可测试的复制 / 粘贴规划与 Pinia 应用剪贴板状态。作者、资源、目标来源均脱离引用；ticket 拒绝迟到发布，关闭源工作簿及 Studio 销毁会清理在途状态。
- [chart-image-export.ts](../../packages/office-studio/src/features/charts/chart-image-export.ts)：浏览器平台 I/O；实际 `drawImage`、`toBlob('image/png')`、object URL 和连接到 body 的下载 anchor，含取消、超时、失败 / 迟到回调清理。
- [use-chart-object-actions.ts](../../packages/office-studio/src/features/charts/views/use-chart-object-actions.ts)：Vue 宿主会话、忙碌状态、原生剪贴板读取、原生 grid paste 接线和浏览器导出。切表、切 Studio、宿主停用、工作簿关闭 / Studio 销毁会取消旧 UI 会话；导出下载前重查权限，菜单完成后恢复焦点还需匹配原宿主 session。
- [ChartObjectContextMenu.vue](../../packages/office-studio/src/features/charts/views/ChartObjectContextMenu.vue)：Ant Dropdown / Tooltip / Button 右键与“更多”菜单，同一目标快照、同一命令能力；Shift+F10 / ContextMenu、Escape 和焦点恢复。
- [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)：四固定根与本地粘贴注册。删除与既有 `remove-chart` 共用真实作者删除 / pruning；粘贴复用 `commitChartChange`，不再初始化源图表默认设置，新增可选宿主当前性断言，沿用现有统计准备、权限、版本、资源发布和一次历史。
- [echarts-renderer.ts](../../packages/office-studio/src/features/charts/echarts-renderer.ts)：新增 `getImageFrame` 只读借用表面 getter；调用方必须在首次 await 前复制，不能自行释放 renderer 持有的表面。
- [ChartOverlay.vue](../../packages/office-studio/src/components/ChartOverlay.vue) 与 [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue)：单击选中，编辑命令打开真实面板；按 `data-chart-id` 找目标，移除 DOM 索引推断；右键 / 快捷键 / API 共用命令。标题 → 图例 → 绘图区的命中优先级与标题 / 绘图区 Delete 语义保留，文本输入 / IME 不抢图表复制粘贴。

计算域沿用完整工作台架构：剪贴板 DTO 与粘贴规划属于 charts 模型 / 控制层，图表作者提交及重计算属于既有命令 / 统计服务，Canvas 编码与系统剪贴板属于浏览器平台 / UI 宿主。粘贴没有新增公式、源数据或布局的权威引擎；已有 Worker / Rust/WASM 准备链照常复用。本批未增加 Rust 计算实现，也未以新依赖清单代替平台调用。

### 复制 → 粘贴 → 历史 / 保存链

1. 捕获 `unitId/sheetId/chartId`、内容 revision、作者 JSON 快照和可选 UI session；校验宿主及源区域 / pivot 源与输出的查看权限。
2. 复制保存作者与引用的遮罩图片资源；不保存计算数组、hover、选区或 renderer bitmap。HTML 与文本标记携带同一个 opaque clipboard ID，使仅有 `writeText/readText` 的浏览器也能核对来源。
3. 浏览器剪贴板不可写时保留内部复制。可读的外部内容优先校验，不能因外部文本、旧 ID 或冲突标记重新粘贴旧图表。原生 grid paste 的图表标记失效时被拦截；没有图表标记的外部文本继续使用既有单元格粘贴规划。
   浏览器写入一旦交给平台，旧请求仍可能完成系统剪贴板写入；本地保证失效结果不发布到新宿主 / 内部缓存，不声称可以撤回平台写入。
4. 同工作簿跨工作表粘贴生成新 chart ID，显式保留原数据所在 `source.sheetId`、pivot 绑定、标题 / 副标题、手工布局、轴、图例、样式、配色与大小；默认位置偏移 +24 px，可用命令参数指定位置。复制源对象 / 图片资源已改变或删除会拒绝粘贴。
5. 新图表引用已有不可变图片资源 ID；引用由全部作者派生，复用既有 pruning 与内容快照，不再维护另一套 refcount。候选通过原有作者 / 统计校验后才整体发布模型、资源和一次内容历史。
6. 复制、编辑打开面板、PNG 导出不创建内容历史。粘贴和删除分别走一次作者内容历史；内部 JSON 保存 / 重开使用原有模型序列化。Clipboard ID、UI session、菜单、编码画布不进入文档内容。

### PNG 表面、复制与内存约束

- 导出前停止拖动 / 草稿和 transient 指示器，重新准备当前作者对应的实际 Canvas；不把 CSS / Ant 选区或浮动控件编码进图片。已提交的作者样式及当前图表显示状态来自真实 renderer 表面。
- 明确的一次像素复制为 renderer Canvas → owned Canvas 的 `drawImage`；源与 owned 表面禁止相同。快照在首次 await 前完成，编码回调期间不依赖仍会变化的借用表面。
- source / owned 对应尺寸最多 **16,777,216 像素**，单个 RGBA backing-store 上界约 **64 MiB**；PNG Blob 输出上限 **64 MiB**，编码超时 **15 s**。成功、失败、取消、超时后 owned width / height 归零，迟到 callback 不再下载。
- 下载成功 click 后 object URL 延迟 1 s 回收；失败立即回收，anchor 始终移除。文件名默认 `${chartId}.png`，本地额外替换 Windows 非法文件名字符。
- 剪贴板 DTO 上限 **32 MiB**，先以作者 / 快照下界拒绝过大作者，再验证最终序列化估算；capture 与 Pinia publication 均有作者 / 引用资源的脱离引用复制。该预算为 UTF-16 控制数据估算，不是整个浏览器实际峰值内存。
- Canvas 编码内部复制、Blob / object URL 的引擎实现内存、source 与 owned 同时存活的峰值、Worker / WASM 既有作者任务传输均尚未实测。没有宣称整条链路零拷贝或达到性能标准。

## 接口与公开契约差异

- 新增四固定命令注册与本地粘贴命令；可从 `FOffice.executeCommand` / Studio 命令 API 调用同一实现。copy / export / paste 声明为本地异步命令，`executeCommandSync` 会在 dispatch 前拒绝；edit / delete 保持同步结果。固定 ES 的 export 本身为 async，copy 经 `Fi` 间接异步；不能只看 `ICommand` 类型声明就称全部同步。
- `ChartOverlay` 增加 `active?: boolean`、`openChartEditor?: (target) => boolean` 与 `select-chart` / `object-change` 事件。事件只驱动选中和从权威模型刷新工作台视图，不创建第二次历史；原 standalone `remove-chart` emit 保留。
- `EChartsCanvasRenderer.getImageFrame` 返回借用 `HTMLCanvasElement/sourceVersion`。宿主会话和 Clipboard 状态由应用拥有，内容模型不存宿主 token。
- P16 时固定 open-focused 的可选 `route` 尚未映射，显式请求会拒绝。后续 P17 已接入本地 Ant 表单、DTO / 能力 / 系列校验和会话；完整差异和未验收项见独立证据，没有标为完整 API 验收。
- 上游复制 hook 纳入通用 drawing clipboard / paste，并处理跨工作簿和 GRID / CHART placement；本地现阶段使用图表专用剪贴板与本地粘贴命令，按 HTML / 文本来源标记接入 grid paste。未补齐一般 drawing group、多对象复制、Office 外部图形 clipboard 格式或跨工作簿依赖导入。

## 草稿与验收状态

本批写入 **34 个 `test` 场景草稿**，按测试入口计数，参数分支不另计：

| 文件 | 场景数 | 验证意图 | 本批运行 |
| --- | ---: | --- | --- |
| [chart-object-clipboard.contract.test.mjs](../../packages/office-studio/tests/chart-object-clipboard.contract.test.mjs) | 10 | 作者 / 图片脱离引用、同工作簿跨表绑定、篡改 / 旧源 / ID / 位置拒绝、资源引用、双格式来源、权限 / 主图、目标别名 / session、ticket、IME / 快捷键、DTO 预算 | 未运行 |
| [chart-image-export.contract.test.mjs](../../packages/office-studio/tests/chart-image-export.contract.test.mjs) | 7 | 实际源身份与一次 snapshot 调用、预算 / preflight、借用所有权、abort / late callback、失败 / 旧作者、超时、anchor / URL 释放 | 未运行；mock 编码不是 PNG 样本验收 |
| [chart-object-commands.contract.test.mjs](../../packages/office-studio/tests/chart-object-commands.contract.test.mjs) | 9 | Facade / 命令复制粘贴历史与 JSON、跨表 / 跨簿边界、事件 / 权限、veto / 旧目标、同步拒绝、主图、route 转发 / 无效组合（P17 改写原预期）、迟到复制 / 关闭、真实资源引用历史管线 | 未运行；资源 / 数值 fixture 不替代 Worker 或图片验收 |
| [chart-object-actions.browser.test.mjs](../../packages/office-studio/tests/chart-object-actions.browser.test.mjs) + [真实工作台 fixture](../../packages/office-studio/tests/fixtures/chart-object-actions.html) | 8 | ExcelWorkbench / Ant / ECharts / Canvas2D 的菜单、Drawer 编辑、剪贴板 / grid paste、真实 PNG 下载及解码像素对照、删除历史、主题 / 主图 / 焦点、迟到 callback 与重开会话不夺焦点、IME 与外部文本边界 | 未运行 |

浏览器 fixture 不替换实际编辑器：使用真实 `ExcelWorkbench`、统一 ConfigProvider、Facade / 作者提交、ECharts 和办公 Canvas；仅观察 / 延迟平台完成回调以覆盖取消，PNG 场景记录实际 owned 编码位图并计划与真实下载文件解码结果比较。浏览器草稿选择 Canvas2D，未由此证明 WebGPU 后端。

本批只读固定来源和当前文件、手工静态审阅并写入代码 / 文档。**没有执行测试、类型检查、前端构建、浏览器、dev / Vite 服务、Cargo / WASM、性能基准、服务验收或 `git diff --check` 等自动检查；没有提交或打包发布。** 所有按钮、菜单、模型、文件和浏览器闭环均保持“未验收”。

## 仍需实现或验证

1. 前端实现 / 验收：`IChartEditPanelRoute` 后续接线与完整 element 直接编辑缺口见 P17；跨工作簿源工作表、pivot 与图片依赖导入；一般 drawing clipboard / 分组多对象；外部 Microsoft Office 图形剪贴板互操作。
2. 前端实现：真正 chart-sheet 创建 / 导入 / 独立布局和 GRID / CHART paste placement；当前 metadata adapter 与 +24 px 默认偏移不构成完成证明。
3. 文件验收：内部 JSON 草稿仍须运行；XLSX / OOXML 的 drawing、chart、源绑定与图片资源写回 / Office 重开仍需本项目解析 / 写出链路和真实样本。不能把内部 JSON reopen 或 PNG 下载草稿当成 Office 转换验收。
4. 跨模块验收：真实系统剪贴板能力 / 权限失败、Canvas2D 与 WebGPU、字体 / 缩放 / DPR、焦点与输入法、合作权限变化 / 服务 ACL、undo/redo 与宿主关闭需要浏览器 / 文件 / 服务证据。
5. 性能验收：同数据集首屏、滚动、编辑、缩放、作者 Worker / WASM 任务、编码耗时、主线程长任务、帧时间、传输与峰值内存均未测；保留“未验收”，不声称已优化或达标。

以上是持续对齐范围内的缺口，不是最终范围缩减；其中路由、跨工作簿导入、图表工作表和外部 clipboard 是前端工作，不能转记为后端自动补齐。
