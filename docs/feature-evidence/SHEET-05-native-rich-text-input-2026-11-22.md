# SHEET-05 原生输入范围与富文本编辑轨迹

记录批次：2026-11-22（Asia/Shanghai）。涉及 SHEET-02/05、COMMON-03/05；开发进度为待验证，验收状态为 **未验收**。

本批按用户安排仅实现、编写测试草稿和静态回读；未执行测试、类型、构建、Rust/WASM、语法、浏览器或基准检查。以下源码事实、已落盘实现与待执行场景分别记录，不把声明数量写成通过数量。

## 1. 固定来源与入口

Univer v1.6.0 固定提交：`eaf4cfa04fabdbcfc922c692f3450459f8a29546`，只读参考根为 `E:\rust_test\univer`。顶部入口清点沿用 [Sheets 固定源码入口表](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)和 [Docs / Sheets 完整入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)。本批沿用 F2、双击、直接键入、单元格编辑、公式栏及确认/取消入口，没有增加 Ribbon 根、插件根或上下文根，未重算运行态溢出数量。

| 固定源码 | 可核实契约 | 本地对应与差异 |
| --- | --- | --- |
| [cell-edit.operation.ts][cell-edit]、[editor.shortcut.ts][editor-shortcut] | F2 通过编辑器可见性 operation / shortcut 打开编辑 | 本地 F2 与公式栏开启 `mode: 'edit'` 会话；直接键入开启 `overwrite`。入口名称与运行保留测试分别记录 |
| [editor-bridge.service.ts][editor-bridge] | 单格编辑使用独立 document model；具有 unit / sheet / cell 归属 | 本地共享一份单元格草稿及 `sessionId`，画布和公式栏分别绑定原生宿主 `hostId`；没有在组件中再创建 OfficeStudio |
| [editing.render-controller.ts][editing-render]：`RichTextEditingMutation`、编辑快照、`SetRangeValuesCommand`、`redoUndoId` | 单格文档修改与最终单元格提交分开；确认使用单元格值命令与历史；验证拒绝可以回滚并阻止移动 | 本地 journal 只记录未提交草稿，最终通过 `sheet.command.commit-cell-input` planner 与公共 API 提交一次内容历史；草稿取消不写入内容历史 |
| [doc-selection-render.service.ts][doc-selection]：beforeinput / input / composition 与逻辑选区 | 输入事件具有源范围，输入桥不能只看最终字符串；缺失事件范围有额外处理 | 本地捕获 library-owned input / textarea 的 `selectionStart`、`selectionEnd`、inputType、data 和 composition；仅能推断的路径使用明确 origin，不冒称捕获了浏览器真实目标区间 |
| [doc-input.controller.ts][doc-input]：`getTextRunForSelection`、`InsertTextCommand`、`ReplaceSelectionCommand` | 输入按活动范围组织文本运行、customRanges 与 customDecorations，交给文档命令 | 本地按每次操作重放单格 `richText`，保留运行 ID、链接、空 marker 和扩展；没有复制上游 dataStream / customRanges 内部模型 |
| [paragraph.ts][paragraph]：`getTextRunForSelection`、`getTextRunAtPosition` | Docs 非折叠选区按 `startOffset + 1` 取样；`isCellEditor === true` 时按 `endOffset` 取样。普通折叠光标继承左邻字符，位置 0 使用首字符；公式单格另有样式分支 | 本地 replacement 继承首个被替换字符，边界 insertion 继承左侧非空运行。跨多个运行的 replacement 与上游单格分支存在差异；保留本地既有契约，尚未以真实 Excel 编辑核验该选择 |
| [doc-ime-input.controller.ts][doc-ime]、[ime-input.command.ts][ime-command]、[doc-ime-input-manager.service.ts][ime-manager] | composition ID、源范围、候选更新、取消恢复；输入命令按 composition 结束组织 history / undo / redo | 本地同一 composition 归属一个宿主，候选更新覆盖同一源范围操作；取消恢复进入 composition 前的 journal。尾随 input 和缺 compositionstart 路径有独立草稿；未证明各系统 IME 或浏览器 reconversion 已通过 |

`paragraph.ts` 与两个 input controller 本次通过固定提交的 `git show` 只读查看；其余来源沿用本批前序源码回读。没有检查或声称上游工作树未修改。Microsoft Office 的真实样本、系统输入法及跨运行选区样式仍是待核实证据。

## 2. 轨迹、权威源与命令

新增 `packages/office-studio/src/features/sheet-workbench/sheet-cell-input-trace.ts`：

```ts
interface SheetCellInputEditTrace {
  version: 1
  sessionId: string
  source: string
  operations: SheetCellTextEditOperation[]
}

interface SheetCellTextEditOperation {
  start: number
  end: number
  text: string
  deletedText: string
  inputType?: string
  origin?: 'beforeinput' | 'composition' | 'programmatic'
    | 'input-caret' | 'compatibility'
}
```

每次操作的坐标是操作前输入值的 UTF-16 偏移。命令边界只重建声明字段，校验范围、删除源文本、代理对边界、连续重放结果及 session ID；最终输入必须与轨迹重放一致。`projectSheetCellInputEditTrace` 逐操作把可选 apostrophe 转义映射回 raw string；文本格式 `@` 不剥离字符。这一步不根据最终公共前后段重新定位运行。

`sheet-cell-edit-session.ts` 的 factory 总会生成独立 session ID，并深克隆 `source.richText`。同文本的运行样式、ID、链接、空 marker 或扩展发生变化时返回 `rich-text` 冲突；基础单元格视觉样式和公式 cache 更新仍允许。旧会话没有运行快照时保留原兼容契约。

`SheetCellInputCommandInput` / `SheetCellInputOptions` 新增可选 `editTrace`；结果新增 `reason: 'edit-trace'` 与 `conflict: 'rich-text'`。planner 复查权威单元格编辑源，提交从最新 `previous.richText` 计算，`session.source.richText` 仅用作冲突依据，不能作为任意运行覆盖来源。输入范围无效、转义投影失败、源运行拼接不一致或 stop 验证失败均不提交内容，草稿保留供取消或重试。

有实际操作的同文本 trace 仍重放运行，避免删除后重输、同文字 replacement 被当作无变化。F2 的显式操作逐步保留未受影响的运行；直接覆盖输入及数字/公式类型转换清除旧运行。没有 trace 的旧编辑 API 保留最终字符串推断路径，因此不能把所有 API 调用都称为精确范围输入。

## 3. 操作归属与兼容边界

| origin / 路径 | 本地记录依据 | 边界 |
| --- | --- | --- |
| `beforeinput` | 操作前选区，实际 input 值；折叠删除结合 Backward / Forward 方向与长度变化 | 拼接必须可校验，选中文本被相同文字替换也记录；不能从 inputType 独立推出任意原生目标范围 |
| `composition` | composition 开始源范围及候选更新 | 可以用首次 composing beforeinput 修正 reconversion 范围；缺 compositionstart 可开组；没有可定位来源时仍标记 compatibility |
| `programmatic` | 明确 replacement 范围和插入值 | 画布 Alt+Enter、多行粘贴、AutoComplete 选项、公式 fx 的空输入；CRLF 规范为 LF |
| `input-caret` | 已记录的前置选区，或 input 后 caret / data 的可校验范围 | 属于推断定位，不能等同于 beforeinput 捕获；spellcheck / autofill 的复杂目标范围仍待浏览器核验 |
| `compatibility` | 单次前后值的公共前后段推断 | 重复文字有固有位置歧义；该路径保留在轨迹，不能冒称样式边界精确 |
| native undo / redo | 回访本会话已经记录的 journal 路径 | 拒绝旧宿主迟到历史与会话源之外的未知值；浏览器合并历史、相同字符串的历史分组及跨宿主原生栈仍需真实核验 |

空源 marker 保持顺序和身份，每个源运行最多输出一次；局部 splice 不复制源 ID。删除全部再输入可以恢复最后删除运行的输入样式及身份，并放回其原 donor 位置；这些属于本地运行模型契约，不冒称上游使用相同算法。

## 4. 输入控件、主题与生命周期

新增 `sheet-cell-native-input.ts` 围绕 `antdv-next` 已拥有的 Input / Input.TextArea 捕获原生事件，没有新建原生控件。WeakMap 绑定每个 DOM element 的 session ID 与 host ID；旧 element 不能自行绑定到另一个会话。同步 `inputContext` 在 Vue props 尚未刷新时读取父草稿；父上下文明确清空会话时不回退到旧 props 的 session ID。

`ExcelWorkbench.vue` 创建共享 recorder，确认时传 trace，Alt+Enter 调画布暴露的 `insertEditorText`。原生输入只经过 `input-event` 更新父草稿；overwrite 第一笔按完整源范围记录。重绑 studio 同步取消旧编辑；已结束或更换的单位/工作表/单格不接受旧事件。undo / redo emit 使用可选 `{ workbookId }`；公共 API、内容历史与生产 App 装配由主任务维护。

`SpreadsheetCanvas.vue` 捕获 beforeinput / input / composition / selection / focus / keydown，AutoComplete 按会话 key，Input → TextArea 切换及延迟 blur / focus 有会话保护。`SheetFormulaBar.vue` 过滤名称框，空闲 focus 开始首个会话时保留原 textarea 与 pointer / IME 选区；结束或替换已有会话后才换 key，不在 begin 的 nextTick 中覆盖 pointer 后续选区。

涉及的 AutoComplete、Input / TextArea、Button、Tooltip 和 Alert 都来自 `antdv-next`，保留统一 ConfigProvider / theme token。安装版为 1.8.2；组件与最小原生桥的事件依据来自安装版源码及 Context7 返回的 TextArea composition / nativeElement、Input 原生事件透传与 AutoComplete change / select 文档。Context7 可用的对应文档版本为 1.8.0，不能替代安装版或浏览器验证。未核验真实明暗主题、焦点恢复或系统输入法。

## 5. 已落盘文件与测试草稿

本专项代码路径均位于 `packages/office-studio/src/`：

- `features/sheet-workbench/sheet-cell-input-trace.ts`、`sheet-cell-native-input.ts`、`sheet-cell-rich-text-edit.ts`。
- `features/sheet-workbench/sheet-cell-edit-session.ts`、`sheet-cell-input-command.ts`。
- `components/ExcelWorkbench.vue`、`components/spreadsheet/SpreadsheetCanvas.vue`、`SheetFormulaBar.vue`。

以下计数仅是静态 `rg -c '^test\('` 的声明数量：

| `packages/office-studio/tests/` 下的文件 | 当前声明数 | 本专项新增声明 | 草稿场景 |
| --- | ---: | ---: | --- |
| `sheet-cell-input-trace.contract.test.mjs` | 19 | 19 | 重复插删、同文本 replacement、删除方向、多步输入、IME 候选/取消/尾输入/缺 start/reconversion、历史/宿主、Unicode、apostrophe 与程序化操作 |
| `sheet-cell-input-command.contract.test.mjs` | 19 | 8 | trace/session/runtime 校验、来源快照注入拒绝、权威运行、同文本运行改变、类型/覆盖及兼容 API |
| `sheet-cell-edit-session.contract.test.mjs` | 10 | 3 | 快照深克隆、同文本运行冲突及旧会话兼容 |
| `sheet-rich-text-input.browser.test.mjs` | 13 | 13 | 真实 library 控件与共享 planner、原生 input、合成 IME、draft/history、程序化多行/选项、取消重试、旧 DOM 的同任务迟到事件、standalone、主题/焦点及内部 snapshot 重开 |
| 合计 | 61 | 43 | 全部未运行 |

新增夹具为 `fixtures/sheet-rich-text-input.html`。浏览器草稿通过 Vite / Vue plugin 与 Playwright Edge 驱动真实组件，强制 Canvas2D；没有运行。这些合成 composition 事件不能证明 Windows 中文输入法、其他系统、其他浏览器或 WebGPU 行为。内部 snapshot 重开也不是真实 XLSX 导出和独立读取器重开。

## 6. 保留缺口

仍需执行有授权的统一验证，并核对真实系统 IME / reconversion、浏览器 compositionend 顺序与 native history、跨多个运行的 replacement 样式、复杂 RTF / HTML 局部粘贴、富文本局部格式工具、链接行为、Office 导出独立重开及完整编辑闭环。本次 trace 解决实现中的范围记录和校验，未交付或验收 SHEET-05 全量能力。

没有设备/浏览器/数据规模、首屏/编辑帧时间、长任务、Worker/WASM 传输与内存实测；不声称优化完成或性能达标。大样本、复杂 shaping、Worker / Rust / WASM 和全部 Sheets / Docs / Pro 目标继续保留。

[cell-edit]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/commands/operations/cell-edit.operation.ts
[editor-shortcut]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/shortcuts/editor.shortcut.ts
[editor-bridge]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/services/editor-bridge.service.ts
[editing-render]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-ui/src/controllers/editor/editing.render-controller.ts
[doc-selection]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/selection/doc-selection-render.service.ts
[doc-input]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/render-controllers/doc-input.controller.ts
[paragraph]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/basics/paragraph.ts
[doc-ime]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/controllers/render-controllers/doc-ime-input.controller.ts
[ime-command]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/commands/commands/ime-input.command.ts
[ime-manager]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-ui/src/services/doc-ime-input-manager.service.ts
