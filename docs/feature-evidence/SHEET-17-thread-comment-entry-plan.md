# SHEET-17 批注与线程评论接入方案

<a id="media-comment-entry-review-2026-11-27"></a>

本文件是较大改动前的静态方案，尚未新增 Sheets 批注或评论产品能力。当前 Sheets 没有普通 Note / ThreadComment 模型、对应内容命令、弹层或评论面板，也没有 Ribbon / 右键入口；已有 [Docs 消息结构](../../packages/office-studio/src/domain/model.ts) 和 [Docs 评论服务](../../packages/office-studio/src/features/document/comments/thread-comments.ts) 可复用作者、时间、消息语义，不应把 Docs 段落锚点套到 Sheets。

固定来源为 `E:\rust_test\univer`，版本 `v1.6.0`，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。范围遵循 [SHEET-17](../UNIVER_FEATURE_PARITY.md)：普通 Note 与 ThreadComment 分模型，单元格和绘图评论均纳入；本地快照保存不能替代真实多人持久同步。

<a id="note-entry-model-review-2026-11-27"></a>

## 固定版普通 Note 入口与内容契约

位置来自 [Note schema](../../../univer/packages/sheets-note-ui/src/menu/schema.ts)，状态来自 [note.menu.ts](../../../univer/packages/sheets-note-ui/src/menu/note.menu.ts)，中文标签来自 [Note zh-CN](../../../univer/packages/sheets-note-ui/src/locale/zh-CN.ts)。三个入口都在 MAIN_AREA / OTHERS，分组 order 0，条目 order 0；schema 按添加、删除、显示/隐藏声明，没有常驻 Ribbon 或 grid 位置。

| 标签 / title key | operation / command ID | 显隐、禁用与激活 |
| --- | --- | --- |
| 添加批注 / `rightClick.addNote` | `sheet.operation.add-note-popup` | Sheets 显示且活动格没有 Note 时显示；工作簿、工作表和选区需可编辑 |
| 删除批注 / `rightClick.deleteNote` | `sheet.command.delete-note` | 活动格已有 Note 时显示；工作簿、工作表和选区需可编辑 |
| 显示/隐藏批注 / `rightClick.toggleNote` | `sheet.command.toggle-note-popup` | 活动格已有 Note 时显示；`note.show === true` 激活；禁用只检查工作簿和工作表编辑权限，不套范围编辑权限 |

title key 的完整前缀为 `sheets-note-ui`。普通 Note 独立统计为 **0 常驻、3 上下文、0 浮动工具栏、0 快捷键**；Note UI 未声明扩展或溢出根。导出的 `SHEET_NOTE_CONTEXT_MENU_ID = 'sheet.menu.note'` 未用作该 schema 的菜单根，不另外计一个入口。与下文线程评论独立计数，不能把三个入口或命令数当成功能完成证明。

[AddNotePopupOperation](../../../univer/packages/sheets-note-ui/src/commands/operations/add-note-popup.operation.ts) 参数为 `{ trigger?: string }`，读取末选区 `primary.actualRow/actualColumn` 打开非临时 popup，默认 trigger 为 `add-note`。打开 popup 自身不写历史；[Note.tsx](../../../univer/packages/sheets-note-ui/src/views/Note.tsx) 在无 Note 且可编辑时准备空内容，经更新命令创建；文本和尺寸变更也经同一更新命令。

| 内容命令 | 参数与模型行为 | 撤销/重做 |
| --- | --- | --- |
| `sheet.command.update-note` | `{ unitId, sheetId, row, col, note: ISheetNote, silent?: boolean }`；新增和更新共用 | 已有 Note 恢复原对象；新 Note 撤销时移除 |
| `sheet.command.delete-note` | 可显式传 `{ unitId, sheetId, row, col }`，缺省取末选区的实际活动格 | redo 移除；undo 恢复完整 Note |
| `sheet.command.toggle-note-popup` | 无内容参数；末选区实际活动格的 `show` 取反 | 同一 toggle mutation 组成 undo/redo；显示状态属于可保存内容 |

上述三条命令来自 [note.command.ts](../../../univer/packages/sheets-note/src/commands/commands/note.command.ts)，都登记工作簿内容历史。四条 mutation 为 `sheet.mutation.update-note`、`sheet.mutation.remove-note`、`sheet.mutation.toggle-note-popup`、`sheet.mutation.update-note-position`，来源为 [note.mutation.ts](../../../univer/packages/sheets-note/src/commands/mutations/note.mutation.ts)；位置 mutation 使用 `{ unitId, sheetId, noteId?, row?, col?, newPosition: { row, col }, silent? }`，不把坐标变换当作重新创建 Note。

[ISheetNote / SheetsNoteModel](../../../univer/packages/sheets-note/src/models/sheets-note.model.ts) 保存 `id/row/col/width/height/note/show?`，其中 `note` 是字符串；固定模型没有作者、富文本 body 或回复。popup 宽高以 CSS pixel 应用，默认取配置或 160×72，并允许编辑者调整尺寸。下文拟模型补 `width/height`，`visible` 对应固定 `show`，需随内容历史和保存恢复；保留 Office 作者、富文本与 VML 属于本地保真扩展，不称为固定 DTO 的同形复制。

[SHEET_NOTE_EDIT_PERMISSIONS / TOGGLE_PERMISSIONS](../../../univer/packages/sheets-note/src/const.ts) 区分更新/删除的范围编辑与 toggle 的工作簿/工作表编辑；[权限 controller](../../../univer/packages/sheets-note/src/controllers/sheets-note-permission.controller.ts) 是命令执行边界。更新跨格时检查新旧位置，显式删除参数检查目标单格；无参数删除分支调用 permissionCheckWithoutRange，不能称为与显式单格检查完全相同。固定 mutation 刻意不在 undo/协作/ref 更新时再次检查权限；本地接入仍按既有历史恢复和当前权限契约验收。普通 Note 不复用线程评论的 workbook comment 或作者限定权限。

[Note 资源 controller](../../../univer/packages/sheets-note/src/controllers/sheets-note-resource.controller.ts) 以 `SHEET_NOTE_PLUGIN` 保存 `{ [sheetId]: { [row]: { [col]: ISheetNote } } }`，包含宽高和 show，加载时写回 NoteModel，卸载时清除该工作簿资源。工作表删除组织 remove/update undo，复制重新生成 Note ID；这些资源随宿主结构历史恢复。该 JSON 是插件快照，不是 Office comments/VML part。

公开 API 的 [FRange Note mixin](../../../univer/packages/sheets-note/src/facade/f-range.ts) 提供 `getNote/createOrUpdateNote/deleteNote`；创建/更新选该 range 起点，尺寸、文字和 show 由调用参数传入。[FWorksheet Note mixin](../../../univer/packages/sheets-note/src/facade/f-worksheet.ts) 提供该工作表全部 Note 查询。菜单取末选区实际活动格与 API 取 range 起点须分别验收。

## 固定版线程评论入口清点

入口位置来自 [schema.ts](../../../univer/packages/sheets-thread-comment-ui/src/menu/schema.ts)，状态与快捷键来自 [menu.ts](../../../univer/packages/sheets-thread-comment-ui/src/menu/menu.ts)，中文标签来自 [线程评论 zh-CN](../../../univer/packages/sheets-thread-comment-ui/src/locale/zh-CN.ts)。标题列同时保留真实语言包 key。

| 类别 | 标签 / title key | 位置与顺序 | operation ID | 状态 |
| --- | --- | --- | --- | --- |
| 常驻 | 打开评论 / `menu.openComments` | MEDIA 第 1 行、第 3 列；order 2 | `sheet.operation.toggle-comment-panel` | Sheets 显示；工作簿/工作表可查看；侧栏当前为评论时激活 |
| 常驻 | 添加评论 / `menu.addComment` | MEDIA 第 2 行、第 3 列；order 2.1 | `sheet.operation.show-comment-modal` | Sheets 显示；工作簿可评论、工作表和当前范围可查看 |
| 上下文 | 添加评论 / `menu.addComment` | MAIN_AREA / OTHERS；order 0 | `sheet.operation.show-comment-modal` | 同单元格添加入口 |
| 上下文 | 添加评论 / `menu.addComment` | DRAWING / OTHERS；order 0 | `sheet.operation.add-drawing-comment` | 工作簿可评论、工作表可查看；对象属于当前工作簿/活动表 |
| 浮动工具栏 | 添加评论 / `menu.addComment` | SHEET；order 10 | `sheet.operation.add-drawing-comment` | 同绘图入口，不构造单元格查看范围 |
| 快捷键 | 添加评论 | Ctrl/Cmd+Alt+M | `sheet.operation.show-comment-modal` | `whenSheetEditorFocused`，执行时仍检查权限 |

标题 key 的完整前缀为 `sheets-thread-comment-ui`。桌面线程评论独立统计为 **2 常驻、2 上下文、1 浮动工具栏、1 快捷键**；本插件 schema 未额外声明扩展或溢出根。4 个 operation 是打开/切换评论面板、打开单元格弹层、添加绘图评论，不能与 5 个内容命令或普通 Note 混算。classic 分组顺序沿用 schema 的 order；具体尺寸和最终布局需浏览器核实。

[mobile schema](../../../univer/packages/sheets-thread-comment-ui/src/menu/mobile-schema.ts) 使用相同常驻、上下文和浮动位置；mobile 添加 operation 共用 `sheet.operation.show-comment-modal` ID，不增加一个桌面入口或内容命令。[mobile menu](../../../univer/packages/sheets-thread-comment-ui/src/menu/mobile-menu.ts) 的单元格右键标签在该位置已有未解决线程时使用“打开评论”，否则为“添加评论”；移动交互仍未验收。

[const.ts](../../../univer/packages/sheets-thread-comment-ui/src/types/const.ts) 给出准确的 `sheet.operation.show-comment-modal`。另一个 operation 为 `sheet.operation.open-comment-panel`。这 4 个入口 operation 没有外部内容 DTO，读取当前上下文；[comment.operation.ts](../../../univer/packages/sheets-thread-comment-ui/src/commands/operations/comment.operation.ts) 的行为如下：

- 单元格弹层取末选区 `primary.startRow/startColumn`；已有未解决线程时激活模型按该位置返回的线程，不在选区每个单元格创建评论。
- 绘图入口取当前焦点的第一个对象，要求其 unit/subunit 等于当前工作簿和活动表；draft anchor 为 `{ kind: 'sheet-drawing', pageId: subUnitId, elementId: drawingId }`。
- 评论面板宽 360；可见、活动线程和 draft 是 UI 状态，不属于内容历史。

## 线程评论内容命令、模型和保存

准确来源为 [comment.command.ts](../../../univer/packages/thread-comment/src/commands/commands/comment.command.ts)；update payload 来自 [comment.mutation.ts](../../../univer/packages/thread-comment/src/commands/mutations/comment.mutation.ts)。

| 命令 ID | 完整参数形状 | 语义 |
| --- | --- | --- |
| `thread-comment.command.add-comment` | `{ unitId, subUnitId, comment: IThreadComment }` | 根评论和回复共用；回复 `parentId` 指向根消息 ID |
| `thread-comment.command.update-comment` | `{ unitId, subUnitId, payload: { commentId, text: IDocumentBody, attachments?: string[], updated?: boolean, updateT?: string } }` | 更新当前作者的消息内容 |
| `thread-comment.command.resolve-comment` | `{ unitId, subUnitId, commentId, resolved: boolean }` | 解决或重新打开线程 |
| `thread-comment.command.delete-comment` | `{ unitId, subUnitId, commentId }` | 删除当前作者的消息；根消息与回复的删除结果不同 |
| `thread-comment.command.delete-comment-tree` | `{ unitId, subUnitId, commentId }` | 找到所属根并删除完整线程 |

[IThreadComment](../../../univer/packages/thread-comment/src/types/interfaces/i-thread-comment.ts) 包含 `id/threadId/dT/personId/text/unitId/subUnitId/ref`，以及 `updateT/authorName/attachments/updated/mentions/parentId/resolved/children` 等扩展字段。单元格 `ref` 是 A1；绘图 `ref` 使用 [comment-anchor.ts](../../../univer/packages/thread-comment/src/types/comment-anchor.ts) 的版本化 `univer-comment-anchor:` 编码。

5 个 mutation 为 `thread-comment.mutation.add-comment`、`update-comment`、`update-comment-ref`、`resolve-comment`、`delete-comment`。其中 ref mutation 参数为 `{ unitId, subUnitId, payload: { commentId, ref }, silent?: boolean }`。泛用内容命令通过异步 datasource 执行后再发 mutation，本文件没有发现这些内容命令自己登记 undo/redo；update/resolve 在异步返回后比对最新对象，避免旧响应覆盖后续编辑。

[tc-resource.ts](../../../univer/packages/thread-comment/src/common/tc-resource.ts) 按 subunit 保存根与 children，经 datasource `saveToSnapshot` 写入资源；[Sheets 资源 controller](../../../univer/packages/sheets-thread-comment/src/controllers/sheets-thread-comment-resource.controller.ts) 把工作表删除/复制关联的 add/delete mutation 纳入结构命令历史。这些机制须分别追踪，不能把 resource snapshot 当作 Office 文件转换或远端服务。

公开 API 还包括 [FRange](../../../univer/packages/sheets-thread-comment/src/facade/f-range.ts) 的 `getComment/getComments/addCommentAsync/clearCommentAsync/clearCommentsAsync`，[FThreadComment](../../../univer/packages/sheets-thread-comment/src/facade/f-thread-comment.ts) 的 `getIsRoot/getCommentData/getReplies/getRange/getRichText/deleteAsync/updateAsync/resolveAsync/replyAsync`，以及 [FWorkbook](../../../univer/packages/sheets-thread-comment/src/facade/f-workbook.ts)、[FWorksheet](../../../univer/packages/sheets-thread-comment/src/facade/f-worksheet.ts) 的查询/清除和 [FUniver](../../../univer/packages/sheets-thread-comment/src/facade/f-univer.ts) 的 builder。绘图线程的 `getRange()` 不伪造单元格范围。

本续批回读本地 [SpreadsheetSheet](../../packages/office-studio/src/domain/model.ts)、[ExcelSheetView](../../packages/office-studio/src/services/excelViewAdapter.ts)、[OfficeStudio](../../packages/office-studio/src/api/office-studio.ts)、[Workbench](../../packages/office-studio/src/components/ExcelWorkbench.vue)、[Ribbon](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) 和 [右键菜单](../../packages/office-studio/src/components/spreadsheet/sheet-context-menu.ts)，确认两类资源均无领域字段、内容命令和界面接线；没有发现可复用现有 Sheets 评论命令修补的小型错接。现有 commitWorkbook / prepareWorkbookChange 可提供工作簿隔离候选、版本与内容历史，仍需下文权限和资源接入，不把 Docs 评论实现算作 Sheets 能力。

## 拟新增本地类型与边界

以下是待统一接入的完整模型形状，尚未修改共享文件。消息作者/时间字段复用 `DocumentCommentMessage`；富文本内容使用现有 `DocumentParagraph`，文字预览由 body 派生，避免同时维护 text/body 两个权威值。不会为此改动 Docs 现有消息类型或抽取通用评论框架。

```ts
import type { DocumentCommentMessage, DocumentParagraph } from '../../domain/model.ts'
import type { SheetDrawingRef } from '../../domain/sheet-drawings.ts'

export type SheetCellCommentAnchor = { kind: 'sheet-cell'; row: number; col: number }
export type SheetCommentAnchor = SheetCellCommentAnchor
  | { kind: 'sheet-drawing'; ref: SheetDrawingRef }

export interface SheetCommentMention {
  userId: string
  label: string
  startOffset: number
  endOffset: number
  sourceId?: string
}

export type SheetCommentMessage = Omit<DocumentCommentMessage, 'text' | 'mentions'> & {
  body: DocumentParagraph[]
  mentions?: SheetCommentMention[]
  attachments?: string[]
  metadata?: Record<string, unknown>
}

export interface SheetThreadComment {
  id: string
  anchor: SheetCommentAnchor
  messages: [SheetCommentMessage, ...SheetCommentMessage[]]
  resolved: boolean
  createdAt: string
  updatedAt: string
  metadata?: Record<string, unknown>
}

export interface SheetNote {
  id: string
  anchor: SheetCellCommentAnchor
  authorName: string
  body: DocumentParagraph[]
  width: number
  height: number
  visible: boolean
  metadata?: Record<string, unknown>
}

export interface WorkbookCommentPermission {
  defaultValue: boolean
  users: Record<string, boolean>
}

export interface SheetCommentActor {
  userId: string
  displayName: string
}

export interface SheetCommentScope { unitId: string; sheetId: string }
export interface SheetCommentDraft {
  body: DocumentParagraph[]
  mentions?: SheetCommentMention[]
  attachments?: string[]
}

export type SheetCommentContentCommand =
  | (SheetCommentScope & { id: 'sheet.command.add-thread-comment'; anchor: SheetCommentAnchor; draft: SheetCommentDraft })
  | (SheetCommentScope & { id: 'sheet.command.reply-thread-comment'; threadId: string; draft: SheetCommentDraft })
  | (SheetCommentScope & { id: 'sheet.command.update-thread-comment'; threadId: string; messageId: string; draft: SheetCommentDraft })
  | (SheetCommentScope & { id: 'sheet.command.resolve-thread-comment'; threadId: string; resolved: boolean })
  | (SheetCommentScope & { id: 'sheet.command.delete-thread-comment-message'; threadId: string; messageId: string })
  | (SheetCommentScope & { id: 'sheet.command.delete-thread-comment'; threadId: string })
```

`SpreadsheetSheet` 拟增加 `notes?: Record<string, SheetNote>` 和 `threadComments?: Record<string, SheetThreadComment>`；`WorkbookModel` 拟增加 `commentPermissions?: WorkbookCommentPermission`。权限未配置时本地默认允许评论，仍须有查看目标的权限。`SheetThreadComment.id` 是稳定线程身份，`messages[0].id` 是根消息身份；回复不再保存另一份 anchor，源 `parentId` 在边界映射到根消息。外部 thread/message/person ID、Office mentions/VML 和未解释 XML 必须按导入映射保留，不能重新生成后声称保真。

外部 commands、快照、Office 解析和远端响应各在真正的入口校验一次，内部依赖构造类型。命令请求不接受 actor、作者或时间冒充当前身份；Studio 从 `currentUserId` 和人员信息构造 actor。人物显示名称、外部 person ID 与本项目认证账号的映射仍需人员服务，不能仅比较显示名称。

菜单、快捷键、Canvas 标记、面板和 API 共用注册命令及能力判断。只读草稿、活动线程、面板开关、筛选留在 feature 的 Pinia UI 状态；评论内容唯一保存于 sheet 模型，并使用现有工作簿隔离历史。同步模型计算保持同步；文件解析使用现有 Worker/Rust 链路，远端 I/O 才使用异步。

## 权限与历史接入点

固定版 [Sheets authz](../../../univer/packages/sheets-thread-comment/src/services/sheets-thread-comment-authz.service.ts)、[泛用 authz](../../../univer/packages/thread-comment/src/services/tc-authz.service.ts) 区分查看、评论和作者操作：查看需要工作簿/工作表查看权限，单元格还需 range view；写入需要 workbook comment，而不是 worksheet/range edit。作者限定编辑/删除消息，所有可评论者可回复和解决；绘图不套 cell range。

本地 [OfficeStudio](../../packages/office-studio/src/api/office-studio.ts) 已有 `currentUserId`、`assertCanViewWorksheet/assertCanViewRange`、`commitWorkbook`、`getWorkbookRevision` 及历史 savepoint。新增评论注册可复用这些边界，但需要同时改以下真实接入点：

- [assertWorksheetProtectionChanges](../../packages/office-studio/src/features/sheet-protection/range-protection.ts) 当前比较除保护规则外的整个 sheet；直接增加 threadComments 后，只有评论权限的只读用户仍会被当作 sheet 编辑拒绝。必须分别比较编辑内容和评论内容，并对两类变化执行对应权限；不能跳过整张工作表保护。
- `prepareWorkbookChange`、operation 替换和历史恢复须检查评论变化及旧/新 anchor 的查看权限。手动改消息、根删除、结构性宿主删除应区分；结构删除由已授权的结构命令携带完整可撤销资源，不伪装成作者手工删除。
- 仅评论变化按现有 commit options 关闭无关公式重算；不为评论创建另一套历史。撤销/重做时按当前权限重新判定，避免通过历史修改其他作者或当前不可查看的资源。
- 当前本地工作簿没有独立的 workbook view 权限契约，不能据工作表 view 宣称远端工作簿授权完整。评论权限管理需限制为现有权限管理员，真实服务接入后由服务端鉴权，不能信任快照里的 grant。
- 远端提交复用 workbook revision/history generation 的发布约束；请求必须有持久 operation ID、当前版本和认证身份，旧响应不能覆盖本地后续编辑。服务端多人持久同步尚未实现，本模型类型和快照不能作为其验收证据。

## 锚点与资源生命周期

| 操作 | 已读本地接入点 | 拟行为与未核实项 |
| --- | --- | --- |
| 行列插入/删除/移动/复制 | [transformWorkbookAxisResources](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-resources.ts)、[轴命令](../../packages/office-studio/src/features/sheet-workbench/sheet-axis-structure-command.ts)、[轴变换](../../packages/office-studio/src/domain/sheet-axis-structure.ts) | 用同一 `SheetAxisEdit` 变换单格 anchor；删除 anchor 时同事务移除并可撤销恢复完整线程。轴复制不能沿用旧 ID；是否复制线程按已核实的复制规则单独处理 |
| 单元格插入/删除移位 | [transformWorkbookCellShiftResources](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-resources.ts)、[planWorkbookCellShiftCommand](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-shift-command.ts) | 复用现有 `mapSheetCellShiftPosition`；带内锚点跟随，带外不变，删除结果同内容候选发布 |
| 排序/筛选 | [sortRange/getSortedRowOrder/expandVisibleSortOrder](../../packages/office-studio/src/features/sheet-sort-filter/sort-filter-engine.ts)、Studio `sheet.command.sort-range` | 评论使用同一 `orderedIndices`，不再排序一次；header、隐藏行、筛选行保持当前映射，范围外列不移动。仅筛选/隐藏不删除评论；空值单元格也可能有线程，不能靠 cell 是否存在决定是否跟随。固定版排序共用 ReorderRange/ref-range 变换，实际浏览器和文件样本仍未验收 |
| 单元格复制/剪切/跨表粘贴 | [剪贴板 resources](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-workbook-resources.ts)、[clipboard types](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-types.ts) | 普通单元格复制在固定版不隐式克隆线程，也不删除目标既有线程；剪切保留线程/消息身份并迁移 anchor，跨表由同一候选移除源并添加目标。普通 Note 的复制与粘贴模式须独立验收，不能直接沿用线程规则 |
| 工作表复制/删除/改名 | [worksheet copy](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-copy.ts)、[worksheet lifecycle](../../packages/office-studio/src/features/sheet-workbench/sheet-worksheet-lifecycle.ts) | 删除随宿主 sheet 内容历史；改名不改基于 sheet ID 的身份。工作表复制在已有 `drawingIds` 映射处重绑绘图 anchor，重新分配线程/消息 ID、保留人员与内容。完整回复复制属于本地拟行为，见下文固定版差异 |
| 绘图删除/移动/复制/组合 | [reconcileWorkbookDrawingStates](../../packages/office-studio/src/features/sheet-drawings/sheet-drawing-model-lifecycle.ts)、[SheetDrawingRef](../../packages/office-studio/src/domain/sheet-drawings.ts) | 以 ref 身份跟随对象，不复制像素位置作为第二锚点。使用实际 owner 删除的 removedRefs，不把媒体缺失/加载失败当删除。拟在实际对象删除时移除其线程并由历史恢复；固定版此项及组合/解组后的结果仍未核实 |

固定版轴行为来自 [ref-range controller](../../../univer/packages/sheets-thread-comment/src/controllers/sheets-thread-comment-ref-range.controller.ts)：range 消失生成 delete/add undo，位置变化生成 update-ref undo/redo，绘图 ref 不登记为单元格范围。

固定版 [SortRangeCommand](../../../univer/packages/sheets-sort/src/commands/commands/sheets-sort.command.ts) 跳过 filtered rows，将 `order[目标绝对行] = 源绝对行` 交给 `ReorderRangeCommand`；评论的上述 controller 通过 [handleReorderRangeCommon](../../../univer/packages/sheets/src/services/ref-range/util.ts) 的范围映射跟随。本地 `orderedIndices` 是相对 data start 的偏移，接入时必须转换同一已排序映射，不能把两种行号协议混用。

普通复制和剪切差异有直接证据：[copy-paste controller](../../../univer/packages/sheets-thread-comment-ui/src/controllers/sheets-thread-comment-copy-paste.controller.ts) 与 [其测试](../../../univer/packages/sheets-thread-comment-ui/src/controllers/__tests__/sheets-thread-comment-copy-paste.controller.spec.ts) 明确普通 COPY 不隐式克隆；跨表 CUT 组织源 delete/目标 add，并按 datasource 同步策略处理 children。工作表复制的 [资源 controller](../../../univer/packages/sheets-thread-comment/src/controllers/sheets-thread-comment-resource.controller.ts) 和 [测试](../../../univer/packages/sheets-thread-comment/src/controllers/__tests__/sheets-thread-comment-resource.controller.spec.ts) 在关闭 mutation 同步时省略 children。本项目不应丢弃已经持有的回复；完整复制与该固定分支的差异需记录、样本验证，不能称为源码完全一致。

## Office 导入写出缺口

已读 [parse_workbook/sheet_payload](../../crates/office-parser/src/parser.rs)、[解析 DTO](../../crates/office-parser/src/model.rs)、[xlsx_parts](../../crates/office-parser/src/xlsx_parts.rs)、[workbookNormalizer](../../packages/office-studio/src/parser/workbookNormalizer.ts) 和 [XLSX writer](../../crates/office-writer/src/xlsx.rs)。当前 payload 未映射 Sheets Note/ThreadComment，没有评论 sidecar；writer 的 `SheetInput` 无评论字段，`add_unsupported_warnings` 将 `metadata.comments` 列为 `UNSUPPORTED_FEATURE`。归一化的 `metadata.source` 只保留已解析 raw payload，不含被解析链路遗漏的评论 part，不能据此声称原始 Office 评论已保存。这是当前项目链路的缺口，不代表 rxls 底层没有可扩展能力。

必要工作必须放在本项目内：

- Note 读取/写出 legacy comments 的 authors、comment ref 和富文本，并保留 VML 显示/尺寸/锚点及 worksheet relationships。Note 不转换成只有一条消息的 ThreadComment。
- ThreadComment 读取/写出 threadedComment parts 与人员 parts，保留根/回复 `id/parentId/personId`、时间、解决状态、mentions 及未知扩展。认证账号与 Office person ID 分开映射。
- 按 OPC relationships 定位 part，复用 `xlsx_parts`，必要时有依据地扩大其可复用接口；不得假定固定路径或仅按 `comments*.xml` 文件名判断。未知 XML/关系按已有 Office source 保真模式保存一次，不在每条消息重复复制整份 package。
- 行列/移位/排序/复制改变 anchor 或身份时，同步重写对应 Office source 绑定；不能内部位置变了而重新导出旧 ref。现代评论与可能伴随的 legacy Note/VML 需样本确认，不能按同坐标盲目合并或删除。
- Rust parser、共享 IO DTO、前端归一化、view adapter、快照和 writer 必须一起接入。只增加 `sheet.threadComments` 并在 JSON 中重开，不满足 Office 验收。
- XLS/XLSB/ODS、未知来源绘图评论、富文本/mentions 边界及附件导出仍需单独确认，未核实之前保持清单未验收；没有给它们添加空操作或永久禁用来替代实现。

## 文件归属与估算

以下均为拟改动，不是本批产品修改，也不是对较大实现的批准。

| 归属 | 拟文件/接入点 | 大致代码量 |
| --- | --- | --- |
| Sheets 专属 | 新 `features/sheet-comments/sheet-comment-types.ts`、`sheet-thread-comments.ts`、`sheet-comment-anchors.ts`；新 API 评论 command module | 450–650 行 |
| Sheets 专属 UI | 新评论面板和单元格 popup，feature UI store；控件统一 antdv-next Form/Input/Modal/Popover/Select/Button，Canvas 仅绘制标记和命中 | 300–450 行 |
| 根任务共享 | `model.ts`、`parser/workbookNormalizer.ts`、`excelViewAdapter.ts`、`model-utils.ts`、Studio 注册/权限/历史、ExcelWorkbench、Ribbon、右键菜单和语言包 | 150–250 行，由根任务接入 |
| 既有资源文件 | 轴、单元格移位、sort-filter、clipboard、worksheet copy、drawing lifecycle 与保护差异判断 | 150–250 行，先协调文件归属 |
| Rust/Office | 新 Sheets 评论 sidecar/writer 专属模块；`parser.rs/model.rs`、共享 IO DTO、`xlsx_parts.rs`、`xlsx.rs` 接入 | 700–1100 行，须独立确认模块归属 |
| 验证 | 专属模型/权限/命令/资源合同、Office fixtures、浏览器输入和 Worker/WASM 边界测试 | 350–550 行，只测试真实回归与闭环 |

本地/Office 接入估计 2100–3250 行；真实持久评论服务、认证/人员与多人版本协议需要另行确认服务文件和规模，仍在交付范围内。本续批仅修改本证据，不修改上述共享文件。`workbook-outline-resources.ts` 已核对，仅导入 outline resource，不是通用评论钩子；不向该文件塞入评论业务。

## 必要验收与当前状态

后续单元/契约应验证两类 anchor、作者限定、只读可评论者、不可查看目标、解决/重开、根删除/回复删除、富文本/mentions、旧异步结果及文档隔离历史；资源合同覆盖轴/移位/排序/稀疏空格/隐藏行/跨表剪切/工作表复制/真实绘图删除和对应撤销。

普通 Note 独立验证三个右键入口、新增/更新/删除、末选区实际活动格与 API range 起点、宽高调整、show 持久状态和对应内容撤销/重做；覆盖不可编辑目标、跨格新旧位置权限，以及工作表复制/删除后尺寸、文字和 show 恢复。Note 的 Office 作者/富文本/VML 保真仍须真实文件往返，不能用线程评论消息测试替代。

浏览器必须覆盖两个常驻根、两类上下文、浮动工具栏、快捷键、单元格 popup 和面板闭环，检查 antdv-next 主题/禁用/焦点恢复/IME；Office 样本必须经过导入→编辑→导出→重新导入验证 Note、完整线程、人员和锚点，不能仅比较 JSON。真实多人验收需要两个认证主体、服务重启后持久恢复及版本冲突，不能用本地数组或 BroadcastChannel 证明完成。

本续批只补本文的普通 Note 入口、DTO/尺寸、权限、内容历史和资源保存证据，以及两类独立分类口径；未修改生产代码、领域模型或测试草稿。测试、类型检查、构建、Rust/WASM、浏览器、文件样本和性能基准全部后置，未启动 dev。本方案中的 UI、Office 与多人能力均为未实现/未验收，原 2100–3250 行及 Office/持久服务方案仍待确认，不能标为 SHEET-17 完成。
