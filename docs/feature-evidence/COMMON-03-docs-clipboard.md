# COMMON-03 Docs 剪贴板

核对日期 2026-11-26，固定 Univer v1.6.0 提交为 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。当前状态为未验收。

## 固定入口

| 位置 | 入口与状态 | 固定版来源 |
| --- | --- | --- |
| 开始 → HISTORY | 粘贴 `doc.command.paste-special` 的 classic 顺序 4。grid 为第 1 行第 3 列，跨两行，显示名称。三个选项依次为保留源格式、匹配目标格式、仅保留文本。工厂按文档编辑权限禁用，形状文本编辑期间隐藏。 | [schema](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts) · [工厂](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/context-menu.ts) |
| 正文右键 → QUICK、FORMAT | QUICK 顺序为复制、剪切、普通粘贴。FORMAT 的选择性粘贴在删除前，各模式与功能区相同。复制和剪切检查非折叠选区及复制权限，剪切额外检查编辑权限。 | [右键 schema](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/menu/schema.ts) · [命令](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/clipboard.command.ts) |
| 画布粘贴 | 原生系统事件取得 text/HTML 及其他格式，Ctrl/Cmd+Shift+V 将下一个粘贴设为纯文本。另有贴在当前光标的三模式浮层，允许修改同一次粘贴的历史项。 | [渲染控制器](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/render-controllers/doc-clipboard.controller.ts) · [模式控制器](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/render-controllers/doc-paste-options.render-controller.ts) · [剪贴板服务](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/clipboard/clipboard.service.ts) |

本批对齐现有常驻粘贴根 1 个及其 3 个子项，补齐 classic 顺序。开始区已有复制、剪切两个本地扩展项，继续沿用。正文右键复用已有 QUICK 三项和 FORMAT 一个子菜单根，下设三个模式。常驻、扩展、独立溢出、上下文根和子项的新增数均为 0，More 仍由共享 Ribbon 呈现。

## 模型、命令和交互

- 工具栏、右键和原生 Ctrl/Cmd+C/X 共用选区序列化，系统剪贴板同时写入 `text/plain` 与 `text/html`。复用矩形表格的 TSV/HTML 和多图片序列化。末尾正好落在下一段起点时保留换行，避免带出下一段格式或 legacy 代码正文。部分代码复制只写选定源文本。复制要求可查看和可复制，剪切另需可编辑，并由文本、表格或图片命令判断实际锁及目标。
- 纯文本写入降级或系统写入失败时保留待剪切内容。成功剪切通过现有 `doc.command.insert-text`、`doc.command.remove-image` 或 `doc.command.batch-images` 进入文档内容事务。原生回调在同一事件内写 DataTransfer，菜单使用异步 Clipboard API。异步结果校验文档、选区、活动宿主、历史世代、权限和选中对象，请求过期后不会覆盖新内容或恢复旧焦点。
- 功能区、右键和原生事件共用粘贴入口。普通粘贴与原生 Ctrl/Cmd+V 默认保留源格式，Ctrl/Cmd+Shift+V 对下一次原生事件选择纯文本。各模式继续使用原有 `doc.command.paste-blocks`、`doc.command.insert-text` 与表格范围命令，代码源始终接收文本，富文本保留源强调与目标字符格式的行为沿用现有 planner。三种模式均通过内容提交写入模型和历史，快照保存读取提交后的文档模型。
- 可替换控件沿用 `antdv-next` Ribbon、Menu 和 Input.TextArea，主题和弹层继续取统一 ConfigProvider。根据 [Antdv Next Menu 本地文档](../llm.md) 并对照安装中的 1.8.2，菜单使用已有 items、disabled 和 click 约定，画布 TextArea 只转发系统剪贴板事件。

实现文件为 [DocumentWorkbench.vue](../../packages/office-studio/src/components/DocumentWorkbench.vue)、[DocumentCanvasViewport.vue](../../packages/office-studio/src/components/document/DocumentCanvasViewport.vue)、[document-ribbon.ts](../../packages/office-studio/src/components/document/document-ribbon.ts) 和 [document-clipboard.ts](../../packages/office-studio/src/features/document/services/document-clipboard.ts)。本批新增 Workbench 与 Canvas 私有剪贴板事件，无公开 Facade 或文件格式改动。

## 未执行草稿和剩余差异

追加了 4 条 [模型和权限契约草稿](../../packages/office-studio/tests/document-paste-special.contract.test.mjs)，覆盖反向部分选择、Unicode、尾段与代码边界、权限和直接文本粘贴。补了 2 条 [浏览器草稿](../../tests/browser/docs-toolbar-table.spec.ts)，覆盖 Ctrl+Shift+V 后下一次普通粘贴、匹配目标格式，并修订了现有剪切、Rich/text 模式、过期读取草稿及 [菜单接线草稿](../../packages/office-studio/tests/document-settings.contract.test.mjs)。按统一验证安排，未执行测试、类型检查、构建、浏览器、WASM、文件重开或基准。

最近粘贴模式浮层及同一历史项的改写仍待接入。当前 HTML 与文本剪贴板支持有限，原生图片文件、固定版自有剪贴板片段与内存回退、正文跨结构块选择的完整资源与编号、legacy 页眉页脚原生桥、形状文字及移动端行为均未验收。还需验证真实系统剪贴板/IME、权限与内容锁组合、明暗与 RTL、焦点恢复、内部快照和独立 Office 文件重开，以及大文件选区的内存和响应时间。
