# DOC-03 页面方向初始化与可选 metadata

日期：2026-11-28（Asia/Shanghai）。状态：**已修改，未验收**。

用户报告初挂载错误：`Cannot read properties of undefined (reading 'documentPageOrientation')`，来源为 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 的 `props.modelValue` immediate watcher。

## 来源和实际原因

[DocumentModel](../../packages/office-studio/src/domain/model.ts) 明确将 `metadata` 定义为可选。`normalizeDocument` 归一化块结构、文本和统计，保留元数据的有无；它不会给缺省元数据强行添加对象。原 watcher 已调用归一化，但最后仍直接读取 `next.metadata.documentPageOrientation`；因此省略可选字段的合法文档在初始化时报错，API 输入状态同步也存在相同读取。页面设置服务另有三处不受保护的模式读取，导致同样文档进入页面设置或有效节布局后再次失败。

固定 v1.6.0 的页面设置入口、命令、模式与几何契约继续见 [DOC-03 分节和页面设置证据](DOC-03-section-breaks.md#doc-page-settings-command-transaction)。本批不改变入口、命令 ID、布局分组或可见性；常规、扩展、上下文与溢出根增量均为 **0**。

## 实际修改

- [document-settings.ts](../../packages/office-studio/src/features/document/services/document-settings.ts) 增加纯函数 `documentPageOrientation`：显式且合法的 `portrait` / `landscape` 优先；否则按实际纸张宽高确定，正方形默认纵向。它不修改模型，不生成作者元数据。页面设置候选、旧内容宽和有效节几何明确允许缺省元数据，沿用传统模式默认和现代节方向的既有默认。
- [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 在任何候选发布及标题、纸张、缩放、方向读回前归一化输入；初挂载、同文档模型替换和 API 输入状态同步使用同一个方向解析函数。没有用空方向或固定方向掩盖真实几何，也不改变元数据的可选契约。
- [document-settings.contract.test.mjs](../../packages/office-studio/tests/document-settings.contract.test.mjs) 新增七组方向场景和一个缺省元数据的有效布局／页面设置场景，检查横向、纵向、正方形、显式覆盖、未知值、源模型不变及未主动添加 `documentFlavor`。
- [document-toolbar.browser.test.mjs](../../packages/office-studio/tests/document-toolbar.browser.test.mjs) 新增两个真实 Workbench 草稿：缺省元数据的首次挂载、同 ID 更新和显式方向读回；页面设置成功后继续编辑、撤销和重做。使用既有 Canvas、Ant 配置及父层历史夹具，不代表已经覆盖生产 Facade 历史或真实 Office 文件。

## 页眉页脚 desktop / mobile 来源续核

固定版 [desktop ui.controller.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/ui.controller.ts) 注册 [DocHeaderFooterRibbonOperation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/operations/doc-header-footer-ribbon.operation.ts)，复用 `sidebar.operation.doc-header-footer-panel` ID 显隐上下文 Ribbon。只有 [mobile ui.controller.ts](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/mobile/ui.controller.ts) 注册基础 Sidebar handler。因此本地 desktop 的 Ribbon 操作保持正确，不能只读基础侧栏处理器就将它替换成 Modal／侧栏。

方向初始化批没有改动页眉页脚生产代码或新增侧栏。原生节／story 导航、字段、输入和内容历史仍复用既有链路，后续 owner 续修见下文，完整功能范围不变。

### 页眉页脚首个输入会话的 owner 续修

后续载入检查发现另一个实际缺口：原 Workbench 的 editor 注册监听只在 `getInputState()?.valid` 时绑定 owner，但 `OfficeStudio` / `FDocument` 构造与 `activateDocument` 不会创建首个输入会话；该监听的 source 又只有 Facade 引用和历史 generation，首次光标选择不会重试注册。因此新载入文档可出现公开打开命令始终找不到实际 Workbench owner 的情况。

[DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue) 现将 owner 绑定到存在的 Facade 和当前历史 generation，不要求预先存在输入会话；打开 story 可正常建立首个会话。执行时原有 source、revision、history、input generation、已存在输入会话的 valid 检查、权限及文档模式判断全部保留；这不是放开陈旧／无效会话写入。

[document-story-input.browser.test.mjs](../../packages/office-studio/tests/document-story-input.browser.test.mjs) 新增一处 `test.each`，展开 classic / grid **2 个计划案例**：明确初始输入会话不存在，执行真实 Studio 打开命令创建 story 会话；通过固定 desktop operation 隐藏再显示上下文 Ribbon 时保留 story 和输入，最终关闭恢复正文且无内容／历史变化。[document-header-footer-ribbon.contract.test.mjs](../../packages/office-studio/tests/document-header-footer-ribbon.contract.test.mjs) 同步既有断言到已在 HEAD 接入的公开关闭命令 `doc.command.close-header-footer`，不是新增功能或测试案例。

## 验证边界

方向初始化批新增 **8 个契约计划案例、2 个浏览器计划案例**，运行 **0**。该批限定只读检查通过四份代码／草稿文件的五段脚本（含一段浏览器夹具）、一个 Vue 模板及当时本记录七条本地链接（含显式锚点）；限定 `git diff --check` 通过。第一次语法检查因根目录没有直接导出 `@vue/compiler-sfc` 而未开始解析，已改用本地安装的 `vue/compiler-sfc` 完成上述只读解析，未安装依赖。页眉页脚 owner 续修另增 **2 个浏览器计划案例**，其静态检查单独收尾。

owner 续修后的限定静态收尾通过 **6 份代码／草稿文件、8 段脚本（含 2 段夹具）、1 个 Vue 模板、154 处本地导入和本记录 10 条本地链接**，限定 `git diff --check` 通过。扫描覆盖 Workbench 同步 immediate watch、Canvas 输入桥及相关弹窗首挂载状态；保存的 execute/release 回调不等同于绑定时已经执行，不把延迟闭包中的后声明依赖误判为首次挂载错误。

测试、类型检查、前端构建、Rust/WASM、dev、浏览器、Office 往返和性能基准继续统一后置。真实首屏、API 会话更新、主题、焦点、IME、Office 文件及性能仍保持未验收；静态结果不替代运行验收。
