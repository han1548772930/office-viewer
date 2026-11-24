# DOC-11 内容控件：公开入口、模型契约与本批实现证据

状态：**未验收**。本记录覆盖内容控件 Pro 插件的公开入口和本批实现；不表示 DOC-11、Docs 工作台或 Pro 功能整体完成。本轮按用户要求先实现、后统一测试，未执行测试、类型检查、构建、Cargo/WASM 检查、dev、浏览器或基准。

第二批更新日期：2026-11-24。临时卸壳、普通内容编辑同步服务、IME 候选预览和焦点 token 已落盘；共享文本命令的最终接线与卸壳后的光标恢复见下文接线状态，尚不能据此标记普通输入闭环完成。

关联：[功能对齐清单](../UNIVER_FEATURE_PARITY.md)、[引用规则](../UNIVER_REFERENCE_RULES.md)、[执行标准](../../PROJECT_EXECUTION_STANDARD.md)、[渲染与计算标准](../RENDERING_AND_COMPUTE_STANDARD.md)。

## 来源与版本

| 编号 | 实际查阅来源 | 可复查锚点与用途 |
| --- | --- | --- |
| S1 | `E:/rust_test/univer`，Univer v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546` | `packages/core/src/types/interfaces/i-document-data.ts:592–672`：`SdtKind`、`SdtLock`、`SdtPlacement`、`ISdtRangeProperties`、`ISdtCustomRange`。公开模型包含 11 种类型，直接插入排除 `repeatingSectionItem`。 |
| S2 | `.research.local/univer-1.6.0/packages/docs-sdt-ui/package.json` | 包名 `@univerjs-pro/docs-sdt-ui`、版本 `1.6.0`；公开发行包的构建提交与固定 OSS 提交是否一致未核实。此缓存用于查阅，没有把商业 SDK 添加到本项目依赖。 |
| S3 | `.research.local/univer-1.6.0/packages/docs-sdt-ui/lib/es/index.js` | `docs-sdt-ui.menu.insert`、插入数组 `ln`、schema 对象 `Tn` / `En`、`InsertSdtMenuFactory` 和上下文菜单工厂对应导出、13 个选中控件 UI 命令。文件为压缩发行产物；按字符串锚点复查，不把压缩内部变量当稳定 API。 |
| S4 | 同包 `lib/types/menu/{schema,menu}.d.ts`、`lib/types/commands/{insert-sdt,sdt-ribbon}.command.d.ts`、`lib/es/locale/zh-CN.js` | schema 名称、`InsertableSdtKind`、选中控件命令参数，以及下表中文标签。 |
| S5 | `.research.local/univer-1.6.0/packages/docs-sdt/package.json`、`lib/types/commands/{sdt-properties,update-sdt-value,repeating-sdt}.command.d.ts`、`lib/es/index.js` | `@univerjs-pro/docs-sdt@1.6.0`；`rangeId` / `unitId` / `segmentId`、属性和值修改、重复节命令、绑定同步接口。`doc.command.update-sdt-properties` 与 `doc.command.remove-sdt` 的发行代码证明 wrapper lock 在 remove 时检查，不统一禁止属性修改。 |
| S6 | [Univer 内容控件公开规格](https://docs.univer.ai/zh-CN/guides/docs/features/content-controls) | 功能清单已有的公开规格链接。最新网页与 v1.6.0 的差异、真实公开演示和 Microsoft Office 文件样本本轮尚未重新核对，不能作为已经完成交互验收的证据。 |
| S7 | 已读取的 `antdv-next@1.8.2` 本地 DatePicker / Form / Checkbox 类型与实现；Context7 的 Dayjs 文档 `/iamkun/dayjs` | DatePicker 实际接收 Dayjs，change 第二参提供日期字符串；`valueFormat` 不保证对传入 ISO 字符串做适配。Dayjs 的非严格解析会归一化日期，因此保留独立的严格 ISO 校验，不通过 `toISOString()` 转换本地日期。 |
| S8 | [Microsoft Open XML `TemporarySdt`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.temporarysdt?view=openxml-3.0.1)，2026-11-24 已读取 | 规格说明 parent structured document tag 在 contents modified 后被移除；卸壳保留内部内容。固定 S1 存在 `temporary` 字段，S3 公开临时开关；本轮未在 S5 发行代码找到自动卸壳运行时，不能推断其内部算法。项目以实际文本或显式语义值变化触发，no-op 和纯格式候选保持壳；此选择尚未用 Word 行为验收。 |
| S9 | [Microsoft Open XML `SdtContentDropDownList`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sdtcontentdropdownlist?view=openxml-3.0.1)，2026-11-24 已读取 | 下拉内容不可直接编辑，无论 locking；选择已有列表项允许。普通输入候选拒绝改变下拉正文，显式选择命令校验保存值与显示标签。 |
| S10 | [Microsoft Open XML `SdtContentDate`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sdtcontentdate?view=openxml-3.0.1)，2026-11-24 已读取 | date 可为单 run / paragraph / cell，日期格式应用于其内容；FullDate 的语义为 Last Known Date。手输无效文本的具体 Word 状态未实测，项目保留原文字并单独保存最后有效 ISO 日期，属于实现策略。 |
| S11 | [Microsoft Open XML `SdtContentComboBox`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sdtcontentcombobox?view=openxml-3.0.1)，2026-11-24 已读取 | 可选择列表项或输入自由文字。项目普通输入保存实际文字，不从重复显示标签猜保存值；显式选择仍保留项的保存值。 |
| S12 | [Microsoft Open XML Office2010 `SdtContentCheckBox`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.sdtcontentcheckbox?view=openxml-3.0.1)，2026-11-24 已读取 | 单字符内容、checked 状态与 checked / unchecked glyph。Microsoft 规格的缺省字体是 MS Gothic，本项目按 S3 四预设使用 Segoe UI Symbol；自定义字符和字体可保存，真实 Windows 字体行为未验收。 |
| S13 | Context7 官方 Vitest `/vitest-dev/vitest` 与 Playwright `/microsoft/playwright` 文档，2026-11-24 已读取 | Vitest beforeAll / afterAll 支持 awaited 清理，hook timeout 为数字；浏览器草稿使用 `60_000`。Playwright `keyboard.insertText` 只产生 input，合成 `dispatchEvent` 是 untrusted，不能作为真实操作系统 IME 验收。 |

S1–S5 用于查阅公开架构、入口、交互与类型契约。本项目按 Vue / Canvas / 项目服务模型重新实现，不依赖 Univer Server、Exchange 或其私有协议。

## 清点口径

只统计 DOC-11 插件的公开入口，不混入整个 Docs / Sheets Ribbon，也不把浮层内的取消、保存、日期格子或列表编辑按钮算作功能根。来源 S3 / S4 的根与菜单叶采用不同维度，不能相加后作为功能完成证明。

| 类别 | 根数 | 菜单叶数 | 说明 |
| --- | ---: | ---: | --- |
| 基础常驻 DOC-11 入口 | 0 | 0 | 内容控件入口来自 Pro 扩展。 |
| Pro 扩展常驻入口 | 1 | 10 | Docs INSERT / CELL 中的内容控件插入根与种类选择。 |
| Pro 扩展上下文入口 | 13 | 0 | 选中内容控件时的 3 个上下文组。具体外观、锁定、日期格式和复选样式的选项另列。 |
| 专用溢出根 | 0 | 0 | 此插件 schema 未定义独立 overflow 根；窄视口由工作台布局处理，未进行浏览器核对。 |

可复查结果为 **1 个插入根、10 个插入叶、13 个上下文根**。UI 操作根对应 14 个命令 ID：1 个插入命令按 `kind` 区分 10 种类型，另有 13 个选中控件命令。

## 常驻插入入口

根：`docs-sdt-ui.menu.insert`，中文“内容控件”，位置 INSERT / CELL，order `7.1`，grid `row:2,column:1,showLabel:true`。非 Docs 文档隐藏。子项共用上游 `doc.command.insert-sdt`，参数 `{ kind }`；本项目 UI 的 `doc.operation.insert-content-control` 最终调用注册的 `doc.command.add-content-control`，使用 `{ unitId, control, host? }`。

classic 布局按 INSERT / CELL 分组和下列数组顺序排列；发行 schema 没有给各菜单叶指定独立 grid 单元格，不能凭空补一个位置。创建条件是可编辑的有效选择及对应宿主；实际禁止范围、权限和锁检查由本项目模型/命令再次执行。

| 顺序 | 上游中文标签 | `kind` | 观察到的模型/后续链路 | 实现边界与状态 |
| --- | --- | --- | --- | --- |
| 1 | 富文本 | `richText` | UTF-16 范围、占位、格式保留、锁与临时属性 | 单段落/单 cell 值与占位链路、实际内容修改后的 temporary 卸壳服务及临时跨段落拆分已有；持久跨结构容器仍需实现，共享普通命令接线待补。未验收。 |
| 2 | 纯文本 | `text` | `text.multiLine`、占位、锁与临时属性 | 值命令支持多行开关与 LF 归一化；普通输入同步、首次占位替换、temporary/no-op 服务已落盘，共享提交路径和结构模型仍需补齐。未验收。 |
| 3 | 图片 | `picture` | 持久化 image block、尺寸、原锚点、占位恢复 | 既有图片替换服务与本批 image 几何接通；Office 样本与结构嵌套未验收。 |
| 4 | 复选框 | `checkbox` | checked、两种 Unicode 字符与字体、正文替换 | 本批创建/切换/样式更新真实写入正文与 checked/value/checkbox 状态。未验收。 |
| 5 | 组合框 | `comboBox` | 列表项保存值与显示文字；允许自由值 | antd AutoComplete 浮层和列表属性编辑已有实现。未验收。 |
| 6 | 下拉列表 | `dropDownList` | 列表值必须属于项；显示名与保存值分离 | antd Select 浮层、增删排序/重复值校验、当前值删除拒绝已有实现。未验收。 |
| 7 | 日期选取器 | `date` | fullDate、显示格式/区域/日历/存储格式 | 本批 Dayjs DatePicker 适配、严格 ISO、格式化/区域与日历映射已有实现。非 Gregorian 细节和文件样本未验收。 |
| 8 | 重复节 | `repeatingSection` | 初始 item、容器/后代 ID、复制删除、范围变换 | 既有 inline 单段落路径可复用；完整 block/table/row/cell 重复容器仍受模型限制。未验收。 |
| 9 | 组合 | `group` | 显式 parentId；嵌套范围与独立编辑 | 单段落嵌套、父子锁与最内层选择已有；完整跨结构组合未完成。未验收。 |
| 10 | 构建基块 | `docPartObj` | gallery/category/unique 与占位 | 属性保存已有；部件库选择、部件复用/唯一语义和真实插入闭环尚未实现。未验收。 |

本项目部分已有译名使用“分组”“文档部件”“日期”，与 S4 的“组合”“构建基块”“日期选取器”有差异；这不改变类型 ID。是否统一最终用户译名仍需在完整 Docs 清点中核对。

## 上下文入口

tab `docs-sdt-ui.ribbon.contentControl`，order `110`，仅选中内容控件时出现。classic 布局使用组序和条目序；grid 使用表内 r/c。菜单工厂的显隐/禁用是上游观察值，本项目仍需补充文档权限、宿主和输入 session 判断。

| 组（order） | 条目（order；grid） | 上游标签 / schema key 尾部 | 上游命令 ID | 显隐 / 禁用 | 本项目最终注册链路 |
| --- | --- | --- | --- | --- | --- |
| value（0） | 0；r1c1 | 复选框样式 / `checkbox-style` | `doc.command.set-selected-sdt-checkbox-style` | 仅 checkbox，content lock 时隐藏 | `doc.command.update-content-control`，正文字符/字体随 checkbox 属性更新。 |
| value（0） | 1；r1c1 | 更换图片 / `change-picture` | `doc.command.change-selected-sdt-picture` | 仅 picture | `doc.command.replace-picture-content-control-image`，项目本地上传与 image block 服务。 |
| value（0） | 2；r1c1 | 日期格式 / `date-format` | `doc.command.set-selected-sdt-date-format` | 仅 date | `doc.command.update-content-control`，非占位日期会重写显示正文，内容锁拒绝。 |
| value（0） | 3；r1c2 | 恢复占位符 / `reset-placeholder` | `doc.command.reset-selected-sdt-placeholder` | 有 placeholder 才显示；已显示占位或 content lock 时禁用 | `doc.command.reset-content-control-placeholder`。 |
| value（0） | 4；r1c1 | 允许换行 / `multiline` | `doc.command.toggle-selected-sdt-multiline` | 仅 text；activated 跟随 multiLine | `doc.command.update-content-control`。值命令校验已有，跨段落普通输入语义尚未完整。 |
| value（0） | 5；r2c1 | 编辑后移除 / `temporary` | `doc.command.toggle-selected-sdt-temporary` | 仅 text/richText；activated 跟随 temporary | 属性命令及实际编辑卸壳服务已落盘；共享普通提交与卸壳后 caret 待补，不能标为闭环完成。 |
| behavior（1） | 0；r1c1 | 外观 / `appearance-control` | `doc.command.set-selected-sdt-appearance` | 选中控件；value 跟随 appearance | `doc.command.update-content-control`；本批 Canvas 支持 boundingBox/tags/hidden。 |
| behavior（1） | 1；r2c1 | 锁定 / `lock-control` | `doc.command.set-selected-sdt-lock` | 选中控件；value 跟随 lock | `doc.command.update-content-control`；内容锁/控件锁独立，所有实际修改再次检查。 |
| behavior（1） | 2；r1c2 | 属性 / `properties` | `doc.command.open-selected-sdt-properties` | 选中控件 | antd Modal / Form；差异 patch 最终执行 `doc.command.update-content-control`。 |
| behavior（1） | 3；r2c2 | 删除内容控件 / `remove` | `doc.command.remove-selected-sdt` | wrapper lock 时禁用 | `doc.command.remove-content-control`；保留正文并删除壳，锁/后代检查由服务执行。 |
| actions（2） | -1；r1c2 | 选择上层控件 / `select-parent` | `doc.operation.select-parent-sdt` | 无父控件禁用 | 只改变选区，不进入内容历史；完整宿主与表格位置由 Workbench 接线核对。 |
| actions（2） | 0；r1c1 | 复制 / `duplicate-item` | `doc.command.duplicate-selected-repeating-sdt-item` | 仅 repeatingSectionItem；row/cell/table 不显示 | `doc.command.duplicate-repeating-section-item`；既有 inline 路径，完整结构能力仍需实现。 |
| actions（2） | 1；r2c1 | 删除项目 / `delete-item` | `doc.command.delete-selected-repeating-sdt-item` | 仅 repeatingSectionItem；cell/table 不显示 | `doc.command.delete-repeating-section-item`；既有 inline 路径，完整结构能力仍需实现。 |

schema key 的完整前缀为 `docs-sdt-ui.ribbon.`。三组分别是 `docs-sdt-ui.ribbon.value`、`.behavior`、`.actions`。全部 13 根均纳入清点；上游当前隐藏的 row/cell/table 类型同样属于本项目完整目标，不能据此删掉最终能力。

选择项：

- 复选预设：Cross `2612 / 2610`、Check `2611 / 2610`、Minimal `2713 / 25A1`、Round `25CF / 25CB`，字体都是 `Segoe UI Symbol`。
- 日期格式：`yyyy-MM-dd`、`yyyy/MM/dd`、`M/d/yyyy`、`dd MMM yyyy`、`MMMM d, yyyy`。
- 外观：`boundingBox`、`tags`、`hidden`；标签分别为边界框、开始/结束标记、不显示标记。
- 锁：`unlocked`、`sdtLocked`、`contentLocked`、`sdtContentLocked`；项目还兼容旧 `controlLocked` 别名。

属性表单包含 type / OOXML ID 的只读显示、alias / tag / appearance / lock / color / placeholder，text 的 multiLine / temporary，date 的 format / locale / calendar / storage，列表项，repeatingSection 的 sectionTitle / doNotAllowInsertDelete，docPart 的 gallery / category / unique，已有 dataBinding 的 xpath / storeItemID。本批在保留这些字段的基础上加入列表逐项编辑、添加、删除、排序和自定义复选字符/字体；没有把 XML dataBinding 属性保存等同绑定求值完成。

## 本批实现及边界协议

| 层 | 本批文件 | 责任与实际实现 |
| --- | --- | --- |
| 标量领域服务 | `packages/office-studio/src/features/document/services/content-controls.ts` | create/update/setValue/setChecked 在 detached candidate 中校验，完成正文、字段与范围变化后发布；值与复选命令接入编辑同步服务，完全同值不重写正文/dirty/temporary。列表保存值改变即使显示相同仍为实际编辑。替换复用共享 run splitter，保留边界零长实体。两种范围 transformer 可卸掉实际跨段落的 unlocked temporary 并重接后代；persistent、wrapper-locked 和不可换行类型仍受模型限制。 |
| 普通输入候选服务 | `packages/office-studio/src/features/document/services/content-control-editing.ts` | prepare 最深占位控件的完整逻辑选区；按 body / table cell / header / footer / note 段落 ID 和 UTF-16 同步 value、checked、dateValue、placeholder 和 binding dirty；约束下拉、组合、复选及图片直接文字编辑；temporary 实际内容变化卸壳、wrapper lock 保壳、后代接最近存续 parent。控制计划与字体/host 兼容投影先全部成功再发布。 |
| 表示与属性验证 | `packages/office-studio/src/features/document/services/content-control-presentation.ts`、`packages/office-studio/src/features/document/services/content-control-properties.ts` | Unicode 码点/字体、列表唯一值、严格 ISO、日期 token/引号/转义/短月名/weekday、Intl 日历映射；明确四位年 Gregorian mask 反解析和格式化回验，支持 locale 月名/星期大小写，拒绝猜 yy 世纪或非 Gregorian 显示输入。runtime 属性 DTO 只允许可保存数据；属性差异提交和字段/位置冲突检查。 |
| 命令契约 | `packages/office-studio/src/features/document/commands/content-control-command-types.ts`、`packages/office-studio/src/features/document/commands/content-control-command-state.ts` | 复选/日期/列表浮层共用注册值命令；能力判断读 `permissionRules` 并检查文档/段落/表格/宿主权限和嵌套内容锁。 |
| 几何 | `packages/office-studio/src/features/document/layout/content-control-geometry.ts` | 使用段落或 cell ID、UTF-16、已有 glyph/line DTO；换行、空范围、页复制、裁剪、图片块边界和 parent depth 排序。完全空裁剪不产生可见 fragment，反向命中选择最内层。 |
| Canvas scene | `packages/office-studio/src/features/document/rendering/content-control-scene.ts` | boundingBox 与真实开始/结束括号，hidden 保留模型/命中但不画边界；使用现有 RenderScene 供项目 Canvas 渲染器处理缩放/滚动/裁剪。复选符号作为正文绘制。 |
| 交互控制器 | `packages/office-studio/src/features/document/controllers/content-control-interaction.ts` | delayed request 校验完整源模型、宿主和输入会话。value edit selection helper 对存续控件返回合法范围，对已卸壳 temporary 按原段落范围及前后长度差返回 collapsed end caret，核对未改动前缀/后缀和 UTF-16 / table / host target。同步命令返回后克隆有效会话为 focus token；下一 tick 核对 unit/host/revision/generation/当前选区，再只聚焦当前输入桥。invalid / 多次 revision 推进 / 宿主删除 / 失活拒绝。 |
| 可替换 UI | `packages/office-studio/src/components/document/DocumentContentControlOverlay.vue`、`packages/office-studio/src/components/document/DocumentContentControlPanel.vue` | Checkbox / DatePicker / Select / AutoComplete / Form / FormItem / Button / Tooltip / Popover / Modal / Input / ColorPicker 均使用 antdv-next；控件配色取统一 ConfigProvider token。属性 Panel 接收可选 `readOnly`，Form disabled、显式子控件 disabled、草稿修改处理器与保存前后守卫共用这个上下文；异步验证后再次检查 IME 合成状态，取消保持可用。没有新增原生输入桥。 |
| 画布装配 | `packages/office-studio/src/components/document/DocumentCanvasViewport.vue` | geometry/scene/overlay 接入；IME preview 使用已有 executeDocumentCommand 和同步服务，准备完整 placeholder 选区并检查内容锁/权限。捕获 explicit activeControlId、source selection、host/session，预览卸壳不改变真实控件身份，compositionend 仍走共享命令。focus token 在命令返回后被动读取会话，不为旧 source 续期。 |

组件不解析或写出 Office 文件；标量模型服务不依赖 DOM。几何复用布局 DTO，不另建权威布局引擎。重解析/写出和大布局的 Rust/Worker/WASM 职责沿用项目架构，本批不增加新的 Worker 二进制或 WASM 内存协议。

Canvas 新增 `activeContentControlId?: string` 和两个事件：

```ts
'content-control-selection': DocumentContentControlSelectionRequest
'content-control-command': DocumentContentControlCommandRequest

interface DocumentContentControlSelectionRequest {
  controlId: string
  selection: DocumentSelection
  source: DocumentModel                 // 完整源模型；不是 host scoped view
  host?: DocumentAuxiliaryBlockHost     // undefined 明确为 body
  inputSession?: DocumentInputSession
}
```

命令 request 在以上字段之外携带 typed `command`，只允许 `doc.command.set-content-control-checked`、`doc.command.set-content-control-value` 或 `doc.command.select-content-control-item`。先发 selection，等待 Vue 更新，再重新捕获输入 session；不把选择之前的 generation 用于确认。初次 selection 可以没有输入 token，有 token 时仍核对；有 API 的 command 必须有有效 token。选择、源模型或 host 变化、控件删除、权限拒绝、输入不可用、正文/宿主合成输入、失活和弹层锚点离开可视范围都会取消 pending editor；浮层内部 IME 合成期间保留草稿并禁止确认。确认后由 Canvas 和 Workbench 再检查源模型、宿主、权限和 session，补 `unitId` 后进入项目注册命令。Workbench 提交完成后再检查预期 revision 和重新捕获的 generation；观察者重入改变选区或再提交内容时，不恢复旧请求的选区。

属性面板保存读取最新 API 控件，只有用户真正改变的字段才进入 patch。当前值改变但未改动的属性不会被旧表单覆盖；位置/类型或所提交字段已变化会拒绝保存。候选式发布与全模型深克隆增加复制成本；尚无大样本测量，不能称为性能优化或内存达标。

普通输入候选接口：

```ts
interface DocumentContentControlEditOptions {
  origin: 'text-input' | 'control-value'
  controlId?: string
  selection?: DocumentSelection       // prepare 后、正文修改前的逻辑选区
  activeControlId?: string            // 显式 inside affinity，不从非空首尾位置猜测
}
declare function prepareDocumentContentControlTextInputSelection(
  model: DocumentModel, selection: DocumentSelection, options?: { controlId?: string },
): DocumentSelection
declare function synchronizeDocumentContentControlsAfterEdit(
  before: DocumentModel, draft: DocumentModel, options: DocumentContentControlEditOptions,
): DocumentContentControlEditResult
// 返回 changedControlIds / removedTemporaryControlIds / retainedLockedTemporaryControlIds
```

`documentContentControlCommandIsTextInput` 明确分类 insert-text、delete、paste-blocks、table-set-cell-text、table-set-range-text、replace-range 六种普通内容输入；format、属性命令、字段缓存刷新不自动卸壳。普通非空范围首尾插入按 half-open 范围视为外部；显式合法控件 ID 可指定内部插入。空控件必须携带源 selection，避免把全部同位置控制区都当成输入目标。祖先一起增长；临时祖先卸壳后存续子区的身份和未知 metadata 保留。

两种底层范围 transformer 的末尾新增向后兼容 `ContentControlRangeEditOptions = { activeControlId?: string }`：block 变体为第 8 参，body range 变体为第 7 参。自身及祖先获得内部首尾 insertion affinity；完整删除 active 内容保留空范围供 synchronize 判断 persistent / temporary / wrapper lock，单行类型换行仍拒绝。无 ID 保留 half-open 行为，非法 ID / 范围 / 父级关系提前拒绝；已经被普通结构转换移掉的 temporary 也重新检查源内容锁和祖先锁，不能因候选内已无外壳而绕过检查。共享调用端仍需透传同一合法 ID，不能只在最终同步服务传入后补救早拒。

普通日期输入成功反解析时保存 ISO。无效输入保留正文到 `value`，清掉 `dateValue`，把最后有效 ISO 留在 `metadata.contentControlLastKnownDate`；清空日期也保留最后有效值。这不是已观测的 Word 行为，S10 仅提供 Last Known Date 语义依据。严格 ISO 可直接接受，即使配置非 Gregorian 显示；非 Gregorian 反解析、时代和二位年份仍待实现/验收。显式日期命令必须提供 ISO 或空值，不能用任意 raw 文本伪造日期。

复选字体变化计划会先克隆受影响辅助内容区并用已有 `projectDocumentStoryBlocks` 准备 runs / paragraphs / pageFields 兼容投影；所有检查成功后原子更新候选权威 blocks 和投影。主段落、table cell、辅助 host 不维护第二份控件引擎。这里存在额外候选/host 克隆，复制次数与字节尚未测量。

本记录写入时的接线状态：显式 setValue / setChecked 和 Canvas IME 候选预览已调用新服务；`api/office-studio.ts` / `api/document-editor-commands.ts` 的共享普通文本命令尚未接入 prepare / synchronize，相关集成草稿不能视为可通过。`DocumentWorkbench.vue` 显式命令后只恢复仍存在的控件范围，temporary 卸壳后的 collapsed caret 尚待共享层在 revision/generation guard 后调用 `documentContentControlSelectionAfterValueEdit(request, currentModel)` 并发布选区；Canvas 已能接受有效结果会话的 focus token，但不会替共享层创建光标。非空边界内部换行和完整删除内部内容的 range options 透传仍待协调。这些缺口保留在完整目标内。

persistent 控件完成内部输入后，其新 caret 可正好位于 endOffset；后续输入应延续本次合法内部编辑身份。共享层需要在 checked commit 后保留实际存续控制区的 active ID，temporary 已移除则清掉该 ID。`onSelection` 会清旧 active ID，不能因此把下一次内部输入误算到范围外；也不能从任意新 caret 位于边界就猜内部身份。该连续输入与完整删除后再输入已有浏览器草稿，尚未接线/执行。

## 历史与持久化

最终业务命令通过 `OfficeStudio` 的文档提交和按文档隔离的内容历史执行，支持 BeforeCommandExecute 取消。普通选区、展开浮层与取消草稿不写内容历史。服务的正文/属性/相邻范围在同一个候选内提交；表格内控件从权威 model 重读 paragraph，不继续使用发布前的旧 table/block 引用。

`FDocument.save()` / 重新构造 `OfficeStudio` 是内部模型快照闭环，**不是 DOCX 文件往返验证**。本次主任务另行扩展项目 Rust parser/writer 的 SDT、story 和 note 数据链；本记录不能把这些尚未运行检查和文件样本的在途改动标为已经保真。`rawSdtPr` / `rawSdtEndPr`、未知 metadata、dataBinding 和内容控件范围保留于项目模型，XML 未知语义和跨结构写出仍须用 Office 样本验收。

## 未运行的测试草稿

| 文件 | 有意义的断言范围 | 状态 |
| --- | --- | --- |
| `packages/office-studio/tests/content-control-presentation.contract.test.mjs` | 复选字符/字体和相邻范围、补充平面 Unicode、四个预设、wrapper/content/祖先锁、LF/UTF-16、日期名称/字面量/非法参数、列表 CRUD/顺序/当前值删除、combo 自由值、调用者隔离、runtime DTO 原子拒绝、属性冲突和 permissionRules、table cell 真实正文 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-geometry.contract.test.mjs` | 换行/空范围/end caret、emoji、同范围嵌套、cell 目标、空/部分裁剪、重复页 affinity、image 几何、hidden/tags 和 Canvas 缩放/滚动 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-editing.contract.test.mjs` | temporary/no-op/纯格式、祖先/内容/wrapper 锁和重接、预先删壳不能绕内容锁、完整 placeholder 首输、empty/首尾 inside affinity、两个 planner 的非空边界内部/外部换行与完整删除空壳语义、UTF-16、checkbox glyph/font 及后续非法控件的原子拒绝、Gregorian 反解析/Last Known Date、combo 自由值/下拉拒绝、table cell/range/matrix、body/header/footer/note 隔离和兼容字体投影、跨 host/重复 ID、边界零长实体 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-interaction.contract.test.mjs` | 卸壳后带前后正文的 emoji end caret、存续/no-op 范围、周围文字改变拒绝、结果会话复制隔离、no-op/cancel、观察者选择当前光标、陈旧 revision/generation、输入无效、UTF-16 半个 surrogate、table 坐标、文档/host/disabled/辅助块删除拒绝 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-history.contract.test.mjs` | 真实 Facade 注册命令与 OfficeStudio CommandExecuted 事件、单事务、撤销重做/内部保存重开、普通 temporary/placeholder/date/combo、same-value revision/savepoint/redo、不合法下拉/取消/段落与 note entity 权限、历史隔离、header/footnote 普通输入及字体兼容投影 | 草稿；未运行，普通输入部分依赖共享接线。 |
| `packages/office-studio/tests/document-content-control-workbench.browser.test.mjs` | 真实 Workbench + OfficeStudio + antd 复选/日期/列表/组合输入、temporary 属性/普通输入/卸壳后继续输入和 emoji caret、persistent 合法内部连续输入/完整删除后再输入、same-value date、完整 placeholder 合成 IME/重复 input 抑制/取消/旧 DOM 节点、普通 header/note 输入、键盘/焦点、弹层过期、观察者重入、权限/属性只读、自定义复选、wrapper lock、classic/grid 明暗主题、内部保存重开；模型同步采用 Vue sync watcher，命令观察使用 CommandExecuted，额外断言 listener errors 为空 | 草稿；未运行，普通输入和卸壳 caret 部分依赖共享接线。 |
| 既有 `content-controls.contract.test.mjs`、`content-control-boundaries.contract.test.mjs` | 修正复选插入会缩为持久化单字符后的 fixture offsets；table 引用从权威模型重读；祖先锁夹具在创建 checkbox 后锁定 | 修改草稿；未运行。 |

实际执行仅有文件读取、搜索、公开规格/工具文档查阅、改动读取和编辑落盘。没有运行测试草稿、类型或构建；没有浏览器、系统剪贴板、Windows 真实输入法、DOCX/Word 样本、两渲染后端或性能证据。

## 剩余缺口与完整目标

- temporary 实际内容编辑自动卸壳服务和显式值路径已落盘；共享普通文本提交、显式卸壳后的 collapsed caret 尚待接好，完整浏览器/历史闭环未验收。
- 日期、复选、组合和下拉普通输入约束及 value / checked / dateValue / showingPlaceholder / dirty 同步已有候选服务；共享命令、边界内部换行/完整删除的 affinity 与取消重入仍需完成并逐路径验证。IME 预览不是持久提交验收。
- 当前范围模型以单段落或单 cell 的 UTF-16 范围为主。允许换行且 unlocked temporary 的跨段落修改可先卸壳；persistent / wrapper-locked 跨段落、block/table/row/cell 结构内容控件、完整重复节/组合仍有模型和命令限制，属于必须继续实现的目标。
- docPartObj 的真实部件库与唯一行为、XML dataBinding 求值及绑定控件同步没有完成；保存属性/dirty 不等于可观察闭环。
- 内容控件与 native fields、嵌套图片、linked story、权限变更/历史跳转、取消后的焦点等需要真实集成场景；本批草稿的覆盖不代表这些场景已通过。
- 非 Gregorian 日历、地区历法/时代格式、全部 Office 日期格式与真实 Windows 字体/IME 未验收。
- 大量控件的几何扫描、候选复制、header 重复页和图片内存未测量。未记录首屏、滚动、编辑、缩放、长任务、Worker/WASM 复制字节/次数、峰值内存等指标，性能状态保持未验收。
- DOCX 读写、Word 文件重开、未知 SDT XML 保真和 Pro 对照演示均未执行。本记录不缩减完整功能范围，不把缺口改成永久禁用、提示消息或展示入口。
