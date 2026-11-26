# SHEET-25：切片器显示设置与排序 UI

记录日期：2026-11-26。状态：**实现、合同和浏览器草稿已落盘，未验收**。

本批接通普通 Pivot/Table 切片器的八项 `SheetSlicerSettings`、共享 cache 四项选项、文档样式预览及实际提交边界保护。按用户要求先实现、最后统一测试：仅静态读取和编辑，没有执行测试、类型检查、构建、Rust/WASM、dev、服务、浏览器、基准、计数脚本、链接检查或 `git diff --check`。下文计划场景全部未运行。

## 固定来源与入口边界

固定 OSS：`E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本地已发布包是同版本静态研究资料，没有安装或运行商业 SDK。

| 精确来源 | 可观察契约及本地实现 |
| --- | --- |
| `.research.local/univer-1.6.0/slicer/lib/types/definition.d.ts` 第 78–87 行 `ISlicerViewConfig` | name、caption、rowHeight、startItem、columnCount、showCaption、lockedPosition、style；本地使用同名设置 DTO，保留可选作者值 |
| 同文件第 46–75 行 cache options | sortOrder、customListSort、crossFilter、hideItemsWithNoData 属共享来源字段 cache；本地独立 Ant Modal 提交同一 cache setter |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/types.d.ts` 的 `SheetSlicerSettings`、PatchView/PatchCacheOptions | settings 是 view 配置，cache options 不属于某个 UI 草稿；本地不复刻内部 intent 或资源层级 |
| 同包 `facade/f-slicer.d.ts` 第 89–98、153–168、453–477、531–544 行 | getSettings/updateSettings、getCacheOptions/setCacheOptions、零基 setStartItem；本地正式接口采用这些名称，不引入 setSettings |
| `.research.local/univer-1.6.0/sheets-slicer-ui/lib/es/index.js` 的 `G.getSettings`、`pn`、`xn`、`kn` | 真实 getter、显示设置 dirty、caption 回退、列数/标题/锁定/样式设置；本地 Vue/Ant Form 维护受控草稿，行高保留 EMU 精度，而非保存 UI 整数显示值 |
| 同 ES 的 `gn/yn/wn` 与 `getSlicerPresentation` | 文档默认及内置/自定义样式；caption??name、showCaption=true、columnCount=1、locked=false 为派生值；本地缺省不写成显式作者默认值 |
| 同 ES 的 `lr/ur` | 四个上下文 Ribbon 根和三个 drawing context 项；本批复用既有设置根及 Drawer，不新增顶部根入口 |

入口分类继续见[连接 UI 清点](SHEET-25-slicer-connections-ui.md)和[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)。本批在既有设置根内新增“排序与无数据项”嵌套入口及“还原显示默认值”动作；不把字段数、按钮数或测试计划数当成功能完成证明。Timeline 保持独立展示/时间设置，不提供普通 Slicer settings/cache options；后续正式 Timeline 设置/日期筛选 UI 见[独立证据](SHEET-25-timeline-settings-ui.md)，不计入本文件的 40 声明 / 107 个未运行场景。

## 责任模块与实际行为

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，UI 位于 `components/spreadsheet`，可测试草稿/提交服务位于 `features/slicer`，权威模型、成员投影、校验、命令和内容历史复用既有服务。UI 仅编排主线程交互，没有新增权威排序/过滤引擎或 WASM bulk 协议；大型查询和性能仍受[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)约束。

- 显示设置读取真实 `getSettings()`。纯显示修改提交一次 `updateSettings(patch)` / `sheet.command.update-slicer-settings`；显示与布局/锚定/纵横比混合草稿继续一次 `updatePresentation` / `sheet.command.update-slicer-presentation`，不拆成两条内容历史。未变几何不回写投影 AABB。
- `SlicerPanelEditDraft` 归一化名称和文本标题，但保留 optional 作者缺省。标题显示 `caption ?? name ?? ''`；“使用名称作为标题”清 authored caption，改名称不同时写 caption。“还原显示默认值”保留 name、rowHeight 重置 24，清其余可选显示属性。
- 首个显示项 UI 从 1 计数，命令 startItem 从 0 计数。按钮高度保存 `round(px * 9525) / 9525` 的 EMU 精度。名称唯一性、caption/行高/列数/offset 校验复用正式 settings 校验；非法草稿不提交。
- `listSlicerStyleNames` 纳入文档自定义样式和内置目录。导入的未知引用作为“当前样式”保留，只有明确选择文档默认值才清作者引用。预览调用同一 `slicerPalette(style, type, theme, workbook)`，跟随文档资源及统一主题。
- “排序与无数据项”使用四个 Ant Select，读取真实 `getCacheOptions()`，一次 `setCacheOptions` / `sheet.command.patch-slicer-cache-options`。默认 sentinel 仅存在 UI；命令用 own undefined 清作者值，显式 false 保留。crossFilter=none 禁用 hide 开关但保留已有作者值。
- cache 弹窗维护独立 dirty、stale、busy 和 generation。外部来源/历史/用户/宿主变化或关闭/卸载后拒绝旧提交；重新读取允许新请求，旧 finally 不清新 busy。自己的真实 Before 命令按 normalized patch 对象身份识别，发布事件按已进入执行的请求及克隆 DTO 等价识别；外部同命令 ID/revision 不冒充自己的发布。
- cache 保存使未提交的父显示草稿 stale，保留输入到明确重新读取。读取、打开、重置草稿、取消、主题切换不直接写内容历史；模型提交的撤销、重做及保存继续使用文档唯一历史源。
- `withSlicerPanelCommandOwner` 在真实 `BeforeCommandExecute` 精确匹配本次 patch，并在 `BeforeSlicerChange` 再检查 captured owner/cache/UI generation。Facade 等待期间即使内容 revision 未变，历史清空或用户变化也拒绝旧命令；finally 释放两个监听。
- 父任务将公开 getInfo 改为正式 `ISheetSlicerInfo`，getSlicer/getSlicers 只返回普通切片器。本批合并 Panel/连接 state/夹具用本地 `getSlicerControl(s)` 和 `getModel()` 读取内部两类控制；正式显示/cache 服务继续普通 `getSlicer`、`getSettings` 和正式 setter。

## UI、主题与输入

本批可替换交互全部使用 `antdv-next`：Form/Form.Item、Input、InputNumber、Checkbox、Segmented、Select、Button、Popover、Modal、Drawer、Tooltip 和 Alert。图库样本仅为非交互装饰，按钮由库管理；Canvas 绘制、命中和选区继续既有编辑器内核。

Drawer/Modal 弹层和图库从统一 ConfigProvider / `theme.useToken()` 获取文字、背景、边界和选中颜色。cache 关闭后只向仍连接且 owner/宿主有效的 Ant 触发按钮归还焦点；顶层 Drawer 延续工作台 Canvas 焦点恢复协议。

显示 Form 捕获 composition，组合期间阻止提交、Enter/Escape 和非文本设置动作。来源变化时保留当前组合输入，compositionend 后 stale 草稿禁用。此次使用库 Input，没有新增原生文本桥。真实 Windows IME、Ant 关闭动画和最终焦点行为尚未运行验收。

## 文件与接口

| 文件 | 本批责任 |
| --- | --- |
| [SlicerPanel.vue](../../packages/office-studio/src/components/spreadsheet/SlicerPanel.vue) | 八项设置、caption fallback、默认清除、样式资源、cache 入口、受控草稿及 composition |
| [slicer-settings-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-settings-panel-state.ts) | 防御性设置读取、presence-aware 显示/几何 diff、校验、单次真实 setter |
| [slicer-cache-options-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-cache-options-panel-state.ts) | 防御性 cache 读取、UI sentinel、optional diff、真实 setter 与 prepared DTO 回调 |
| [slicer-panel-command-owner.ts](../../packages/office-studio/src/features/slicer/slicer-panel-command-owner.ts) | 两个实际命令边界的 captured owner 取消及监听释放 |
| [use-slicer-cache-options.ts](../../packages/office-studio/src/features/slicer/use-slicer-cache-options.ts) | 弹窗独立草稿、真实发布识别、stale/busy/generation、旧完成隔离 |
| [SlicerCacheOptionsDialog.vue](../../packages/office-studio/src/features/slicer/SlicerCacheOptionsDialog.vue) | 四个 Select、反馈、默认清除、提交/取消、主题与触发焦点 |
| [slicer-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-panel-state.ts)、[slicer-connections-panel-state.ts](../../packages/office-studio/src/features/slicer/slicer-connections-panel-state.ts) | 配合正式 Facade 分类接口迁移合并控制 lookup/list 与 raw model 读取 |
| [slicer-settings-ui-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-settings-ui-fixture.mjs) | 真来源、文档样式、可选作者设置与 March 上下文 |
| [slicer-connections-workbench.html](../../packages/office-studio/tests/fixtures/slicer-connections-workbench.html) | 真实 Ant/Canvas/Facade 宿主，settings/cache/displayMembers、patch presence 和等待 gate |
| [slicer-settings-ui.contract.test.mjs](../../packages/office-studio/tests/slicer-settings-ui.contract.test.mjs)、[slicer-settings-ui.browser.test.mjs](../../packages/office-studio/tests/slicer-settings-ui.browser.test.mjs) | 新批合同与浏览器草稿 |
| [slicer-connections-panel.contract.test.mjs](../../packages/office-studio/tests/slicer-connections-panel.contract.test.mjs)、[slicer-connections-ui.browser.test.mjs](../../packages/office-studio/tests/slicer-connections-ui.browser.test.mjs) | 旧草稿只迁移 raw 模型读取和纯 caption 的 settings 等待 gate；不增加旧批统计 |

正式 settings DTO 与 Facade/core 校验由父任务维护。本批新增服务接受 captured snapshot、optional 显示/cache patch 和当前 UI owner 回调；cache 的可选 `prepared` 回调只用于捕获实际 normalized command DTO。`updatePresentation(): Promise<this>` 拒绝抛错，两个正式 settings/cache setter 返回 boolean；服务正确处理这两种结果。

## 未运行草稿与样本

按 `rg -n '^test'` 读取声明并人工展开参数，没有执行计数脚本：

| 草稿 | 声明数 | 计划场景 | 状态 |
| --- | ---: | ---: | --- |
| `slicer-settings-ui.contract.test.mjs` | 21 | 64 | 全未运行 |
| `slicer-settings-ui.browser.test.mjs` | 19 | 43 | 全未运行 |
| 合计 | **40** | **107** | 不是通过数 |

合同草稿覆盖八项 getter/setter、纯/混合单事务、own undefined 与显式默认、名称唯一性和非法参数、文档自定义/未知样式、shared/独立 cache、四项排序/六种交叉显示组合、no-op/undo/redo/JSON 重开、权限、pending 来源/历史/用户/宿主/关闭/卸载、Before listener 重入、真实发布和旧 guard/finally。浏览器草稿覆盖 classic/grid、Pivot/Table、Ant 成员排序/隐藏、父草稿保留、焦点、主题、synthetic composition 和 Timeline 隔离。

真实样本复用[连接夹具](../../packages/office-studio/tests/fixtures/slicer-connections-fixture.mjs)：unit `slicer-connections-book`，source `source`，host `reports`，other `other`；A1:C5 Region/Date/Amount、North/South、10/20/30/40；普通/shared slicer、Timeline 与 Table 控制。March 上下文中 North 无数据、South 有数据；自身 North selection 保留，availability 排除自身字段筛选。JSON 保存重开不属于 XLSX/Office 验收。

新批草稿单列；不重计[连接 UI 旧批](SHEET-25-slicer-connections-ui.md)、[cache 核心批](SHEET-25-slicer-cache-options.md)、ΣValues 或其他字段筛选批。

## 未验收与剩余缺口

1. 所有新旧合同、类型、前端/Rust/WASM 检查及实际编辑闭环尚未执行；合并 Facade 接口需要在父任务完成后统一验证。
2. Ant 真关闭动画/焦点、Windows 中文 IME、用户切换的生产响应式 UI、窄屏、WebGPU/Canvas2D 行为仍未验收；synthetic IME 草稿不能代替系统输入法。
3. 真实 XLSX 的 slicer/cache/style parts、排序 options、affine 布局与 native filter 的解析/写出及独立 Office 重开未验收；不以 JSON 替代文件保真。
4. 上游完整 locale/日历自定义排序、文档自定义样式管理与未公开 Pro 行为继续按完整范围追踪，本批只接已有文档样式资源和四项 cache options。
5. 大成员查询/排序的 Worker/Rust/WASM、取消、TypedArray leases、复制/传输字节、长任务、帧时间和真实内存/性能没有测量，保持未验收，不声称优化或端到端零拷贝。

上述事项继续属于完整 Sheets / Docs / Pro 交付范围，没有缩减为显示版或空入口。
