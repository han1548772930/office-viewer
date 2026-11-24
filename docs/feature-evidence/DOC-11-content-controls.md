# DOC-11 内容控件：公开入口、模型契约与本批实现证据

状态：**未验收**。本记录覆盖内容控件 Pro 插件的公开入口和本批实现；不表示 DOC-11、Docs 工作台或 Pro 功能整体完成。本轮按用户要求先实现、后统一测试，未执行测试、类型检查、构建、Cargo/WASM 检查、dev、浏览器或基准。

第四批更新日期：2026-11-25。共享普通文本提交、内部边界 affinity、连续输入、temporary 卸壳后的光标恢复及 custom XML 标量同步已接线。完整 block range 与连续 table row range 的插入、重复项复制/删除、group 结构锁和结果选区已落盘；复制同时重映射表格/续格、字段、图片锚点、内容控件、注释内容与批注身份，删除同事务清理被移除资源。inline 注释编号长度变化使用真实 note/run trace，不以单纯文本差异猜新选区。图片空输入锚点标记、派生 line 和短图 caret 裁剪已接入。父级导航、图片上传、重复项结果均保留 host/revision/generation 守卫；同范围无显式 parentId 的上游父级顺序仍有缺口。代码与测试草稿均未运行，不能据此标记结构、引用、焦点、绑定、图片或 DOC-11 闭环已验收。

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
| S14 | S1 固定仓库 `packages/core/src/docs/sdt-binding.ts`、`packages/core/src/docs/__tests__/sdt-binding.spec.ts` | `getSdtBindingKey` 仅对 storeItemID 忽略大小写，XPath 和原始 prefixMappings 均参与 key；`getSdtBindingValue` 区分占位空值、布尔状态、ISO 日期、列表保存值和文字。单元草稿未在本轮运行。 |
| S15 | S5 的 `lib/types/commands/update-sdt-value.command.d.ts` 及 `lib/es/index.js` 中 `function V(e,t)`、`function H(e)`；S3 的 `doc.command.reset-selected-sdt-placeholder` | 公开 `synchronizeSdtBindings` 要求 peers 合入原始编辑、caret 和 undo 数据；发行代码遍历 body/header/footer/note，使用来源标量 Map、placeholder fallback 与文字选区变换。Map 的后遍历源会覆盖同 key 先前源；复选接受 `true` / `1`，其他字符串置 false。`H` 采用 nullish fallback，显式空 prompt 不被替换；选中控件 UI reset 则要求 truthy `properties.placeholder`，已 showingPlaceholder 时返回 false。本项目保留 UI reset 的默认限制，仅绑定 fallback 允许有效空 prompt。只读取有界片段，不依赖压缩变量作为 API。 |
| S16 | Context7 `/vuejs/core` 的 scheduler / apiWatch、Vitest hooks 文档，2026-11-24 已读取 | Vue `nextTick` 在当前 DOM/job flush 后运行，`flush: 'sync'` 监听立即执行；这些只支撑草稿的模型同步与结果焦点守卫，不是已运行的交互证据。文档返回 main 来源，本地版本以 pnpm 锁与源码为准。 |
| S17 | 本地 `node_modules/typescript/lib/lib.dom.d.ts`，TypeScript `6.0.3`，2026-11-25 静态读取 | `DOMParser.parseFromString`、`XMLSerializer.serializeToString`、`XPathEvaluatorBase.evaluate`、`XPathResult` 的实际声明。用于核对平台适配的 XML/XPath 接口；类型存在不证明浏览器行为通过。XML 节点身份与 scalar alias 规则是本项目实现选择，未冒称上游内部算法。 |
| S18 | Context7 `/websites/vuejs_guide` 返回的 [Vue 模板引用](https://vuejs.org/guide/essentials/template-refs)与 [DOM 更新时序](https://vuejs.org/guide/essentials/reactivity-fundamentals#dom-update-timing)，2026-11-25 已读取 | `script setup` 默认私有，父组件使用的公开方法须由 `defineExpose` 暴露；`nextTick` 等待 DOM 更新。用于核对 Ribbon → Canvas 公开方法及延迟 focus 的接线，版本以项目 pnpm 锁为准，文档查阅不构成浏览器验证。 |
| S19 | S3 的 `lib/es/index.js` 中 `function Wt(e)` / 公开导出 `selectedParentRange`、`doc.operation.select-parent-sdt`，以及 S4 的 `sdt-ribbon.command.d.ts`，2026-11-25 有界读取 | 上游按最小包含范围选择父区；严格包含或同范围且 customRanges 数组索引更后才算父区。操作保留 `segmentId` / `segmentPage`，替换选区、关闭原浮层并选择父区的 `properties` 模式。本项目显式 parentId 权威，保留 host 与真实 table 坐标；local create 的 block/start/id 排序不保留上游 customRanges 的同范围顺序，不能以本地 ID 顺序伪造父关系。上游 properties 浮层模式尚未复现。 |
| S20 | S1 固定仓库 `packages/docs-drawing-ui/src/commands/commands/insert-image.command.ts:27`、`src/controllers/render-controllers/doc-drawing-update.render-controller.ts:160` 与 S3 的 `doc.command.change-selected-sdt-picture`，2026-11-25 有界读取 | 选中图片 SDT 的 UI 命令保留 segmentId，将范围交给 insert-float-image；公开 insertDocImage 捕获原始范围，在 file picker、ImageIo 和解码之后重复合法性/disposed 检查。本项目使用本地单图上传和持久 image block + paragraph anchor，逐阶段检查请求/API/模型/host/选区/会话；multiple、完整 floating/drawing/Office 环绕仍未验收。 |
| S21 | 本地 `antdv-next@1.8.2` 的 `dist/modal/interface.d.ts`、`Modal.js`；Context7 `/antdv-next/antdv-next` 返回的 [Modal API](https://github.com/antdv-next/antdv-next/blob/main/docs/src/pages/components/modal/index.en-US.md)，2026-11-25 已读取 | `afterClose` 是完全关闭后调用的 callback prop，Modal emit 不包含 afterClose；`focusTriggerAfterClose` 可关闭默认 trigger 归还。本项目使用 `:after-close` 和已有结果 focus token。请求传入版本 1.8.2，文档索引显示 1.8.0、返回 main，实际接口以本地 1.8.2 类型与实现为准；未做真实焦点/动画验证。 |
| S22 | S5 的 `lib/types/commands/repeating-sdt.command.d.ts`、`lib/es/index.js` 中 `buildDuplicateRepeatingItemUpdate` / `buildDeleteRepeatingItemUpdate` / `buildDeleteRepeatingRowsUpdate` 对应有界片段；S3 selected repeating 命令和菜单工厂，2026-11-25 已读取 | 普通项目复制生成新的 customRange/ooxml/paragraph ID，更新包含范围和 caret；删除 row 路径用真实 table rows JSONX mutation、保留 segmentId，并在删整表时按剩余长度限制 caret。上游普通复制排除 row/cell/table，row 仅走结构删除。本项目在同一注册命令内独立实现完整块与连续行结构复制/删除，row 复制入口是明确实现差异；结果使用提交后的实际块/行/字符身份。上游分支和本地草稿均未执行。 |

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
| 1 | 富文本 | `richText` | UTF-16 范围、占位、格式保留、锁与临时属性 | 单段落/单 cell 与完整 block range 插入、共享输入、占位及 temporary 卸壳已落盘；完整块包含表格/图片时保留结构，跨 cell 或部分跨段选择拒绝而不丢内容。同绑定标量随同事务同步，完整结构/Office 样本未验收。 |
| 2 | 纯文本 | `text` | `text.multiLine`、占位、锁与临时属性 | 值命令支持多行开关与 LF 归一化；共享提交已接 prepare / finalize / bindings，首次占位替换、内部边界、temporary/no-op 和连续输入已有接线；持久跨结构模型仍需补齐。未验收。 |
| 3 | 图片 | `picture` | 持久化 image block、尺寸、原锚点、占位恢复 | 服务、image 几何和真实 antd 单图上传已接注册命令、延迟请求与结果选区守卫；读取/解码/历史/Office 样本与结构嵌套均未验收。 |
| 4 | 复选框 | `checkbox` | checked、两种 Unicode 字符与字体、正文替换 | 本批创建/切换/样式更新真实写入正文与 checked/value/checkbox 状态。未验收。 |
| 5 | 组合框 | `comboBox` | 列表项保存值与显示文字；允许自由值 | antd AutoComplete 浮层和列表属性编辑已有实现。未验收。 |
| 6 | 下拉列表 | `dropDownList` | 列表值必须属于项；显示名与保存值分离 | antd Select 浮层、增删排序/重复值校验、当前值删除拒绝已有实现。未验收。 |
| 7 | 日期选取器 | `date` | fullDate、显示格式/区域/日历/存储格式 | 本批 Dayjs DatePicker 适配、严格 ISO、格式化/区域与日历映射已有实现。非 Gregorian 细节和文件样本未验收。 |
| 8 | 重复节 | `repeatingSection` | 初始 item、容器/后代 ID、复制删除、范围变换 | inline、完整 block range 和连续 table row range 的初始 item、结构复制/删除、资源重映射与真实结果选区已落盘；独立 whole-cell/whole-table 重复容器仍有缺口。所有路径未验收。 |
| 9 | 组合 | `group` | 显式 parentId；嵌套范围与独立编辑 | 单段落、完整块及连续行 group 已落盘；内容锁只允许真实 child control 内编辑，mutation 后核对全部 outside signature，包含普通 cell/续格输入与资源。全部运行时场景未验收。 |
| 10 | 构建基块 | `docPartObj` | gallery/category/unique 与占位 | 属性保存已有；部件库选择、部件复用/唯一语义和真实插入闭环尚未实现。未验收。 |

本项目部分已有译名使用“分组”“文档部件”“日期”，与 S4 的“组合”“构建基块”“日期选取器”有差异；这不改变类型 ID。是否统一最终用户译名仍需在完整 Docs 清点中核对。

## 上下文入口

tab `docs-sdt-ui.ribbon.contentControl`，order `110`，仅选中内容控件时出现。classic 布局使用组序和条目序；grid 使用表内 r/c。菜单工厂的显隐/禁用是上游观察值，本项目仍需补充文档权限、宿主和输入 session 判断。

| 组（order） | 条目（order；grid） | 上游标签 / schema key 尾部 | 上游命令 ID | 显隐 / 禁用 | 本项目最终注册链路 |
| --- | --- | --- | --- | --- | --- |
| value（0） | 0；r1c1 | 复选框样式 / `checkbox-style` | `doc.command.set-selected-sdt-checkbox-style` | 仅 checkbox，content lock 时隐藏 | `doc.command.update-content-control`，正文字符/字体随 checkbox 属性更新。 |
| value（0） | 1；r1c1 | 更换图片 / `change-picture` | `doc.command.change-selected-sdt-picture` | 仅 picture | `doc.command.replace-picture-content-control-image`，项目本地上传与 image block 服务。 |
| value（0） | 2；r1c1 | 日期格式 / `date-format` | `doc.command.set-selected-sdt-date-format` | 仅 date | `doc.command.update-content-control`，非占位日期会重写显示正文，内容锁拒绝。 |
| value（0） | 3；r1c2 | 恢复占位符 / `reset-placeholder` | `doc.command.reset-selected-sdt-placeholder` | 有 placeholder 才显示；已显示占位或 content lock 时禁用 | `doc.command.reset-content-control-placeholder`；Ribbon、Canvas 与实际值命令共用 capability，另核对支持类型、祖先锁、权限、宿主和输入状态。 |
| value（0） | 4；r1c1 | 允许换行 / `multiline` | `doc.command.toggle-selected-sdt-multiline` | 仅 text；activated 跟随 multiLine | `doc.command.update-content-control`。值命令校验已有，跨段落普通输入语义尚未完整。 |
| value（0） | 5；r2c1 | 编辑后移除 / `temporary` | `doc.command.toggle-selected-sdt-temporary` | 仅 text/richText；activated 跟随 temporary | 属性、共享普通提交和实际编辑卸壳已接线；显式卸壳后使用 revision / generation 守卫恢复 caret，复杂嵌套绑定级联的 caret 映射仍有边界。未验收。 |
| behavior（1） | 0；r1c1 | 外观 / `appearance-control` | `doc.command.set-selected-sdt-appearance` | 选中控件；value 跟随 appearance | `doc.command.update-content-control`；本批 Canvas 支持 boundingBox/tags/hidden。 |
| behavior（1） | 1；r2c1 | 锁定 / `lock-control` | `doc.command.set-selected-sdt-lock` | 选中控件；value 跟随 lock | `doc.command.update-content-control`；内容锁/控件锁独立，所有实际修改再次检查。 |
| behavior（1） | 2；r1c2 | 属性 / `properties` | `doc.command.open-selected-sdt-properties` | 选中控件 | antd Modal / Form；差异 patch 最终执行 `doc.command.update-content-control`。 |
| behavior（1） | 3；r2c2 | 删除内容控件 / `remove` | `doc.command.remove-selected-sdt` | wrapper lock 时禁用 | `doc.command.remove-content-control`；保留正文并删除壳，锁/后代检查由服务执行。 |
| actions（2） | -1；r1c2 | 选择上层控件 / `select-parent` | `doc.operation.select-parent-sdt` | 无父控件禁用 | 显式 parentId 或 legacy 严格包含范围经 helper 解析完整表格位置；Workbench 核对 host/revision/generation 并复用输入焦点恢复，不进入内容历史。无效目标、输入失活或 IME 禁用；同范围隐式顺序与上游 properties 浮层仍未对齐。未验收。 |
| actions（2） | 0；r1c1 | 复制 / `duplicate-item` | `doc.command.duplicate-selected-repeating-sdt-item` | 上游仅 repeatingSectionItem；row/cell/table 不显示 | `doc.command.duplicate-repeating-section-item`；本项目 inline/完整块/连续行共用候选与注册命令，row 显示是独立增强；cell/table 尚未实现。未验收。 |
| actions（2） | 1；r2c1 | 删除项目 / `delete-item` | `doc.command.delete-selected-repeating-sdt-item` | 上游仅 repeatingSectionItem；cell/table 不显示 | `doc.command.delete-repeating-section-item`；inline/完整块/连续行删除真实结构和资源，保留合法 caret。未验收。 |

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
| 标量领域服务 | `packages/office-studio/src/features/document/services/content-controls.ts` | create/update/setValue/setChecked/reset 在 detached candidate 中校验，完成正文、字段与范围变化后发布；图片 reset 的恢复、尾段合并或后续检查失败不会泄漏结构改动。值、复选与 reset 接入编辑同步，完全同值不重写正文/dirty/temporary；已经显示 prompt 的 reset 保持 no-op。列表保存值改变即使显示相同仍为实际编辑。替换复用共享 run splitter，保留边界零长实体。两种范围 transformer 可卸掉实际跨段落的 unlocked temporary 并重接后代；persistent、wrapper-locked 和不可换行类型仍受模型限制。 |
| 普通输入候选服务 | `packages/office-studio/src/features/document/services/content-control-editing.ts` | prepare 最深占位控件的完整逻辑选区；按 body / table cell / header / footer / note 段落 ID 和 UTF-16 同步 value、checked、dateValue、placeholder 和 binding dirty；约束下拉、组合、复选及图片直接文字编辑。显式 reset 的 `placeholderText` 校验 prompt/value/正文一致，不把日期或列表 prompt 当作保存值；日期清 `dateValue` 并留最后有效 ISO。temporary 实际内容或占位状态变化卸壳、wrapper lock 保壳、后代接最近存续 parent。控制计划与字体/host 兼容投影先全部成功再发布。 |
| 标量绑定编排 | `packages/office-studio/src/features/document/services/content-control-bindings.ts` | 来源 DTO 保留 detached control、raw key 和语义 scalar；包括预先卸壳 temporary、表格实际来源和显式 value/reset。raw-key legacy peers 与实际 XML node peers 共用候选同步、锁/权限、级联、选区变换与缓存发布；冲突拒绝整批。直接 peer echo 不再次写 XML，其他实际变化才可级联。所有检查结束后才 commit XML session 并发布候选。显式 temporary caret 使用存续 peer 的实际显示/range，可核对同段落不相交 alias；已卸壳的异 key alias 和复杂嵌套仍有边界。 |
| 表格来源投影与协议 | `packages/office-studio/src/features/document/model/document-table-text-origin.ts`、`packages/office-studio/src/features/document/commands/table-mutations.ts` | DFS 投影包括 object-only/空 cell 与 stored continuation containers；每格来源记录 table/cell 身份、anchor row/column、order 和语义 value。新增文本段落的结构 LF 仅进入 text edit，不混入 temporary scalar。binder 校验原始/候选格身份、顺序、值、段落数、UTF-16、分隔符和完整宿主文字；同 offset 不同空格插入按存储顺序核对，重叠普通范围拒绝。 |
| custom XML 模型/会话/平台 | `packages/office-studio/src/domain/document-custom-xml.ts`、`packages/office-studio/src/features/document/services/document-custom-xml-bindings.ts`、`packages/office-studio/src/platform/dom/custom-xml-runtime.ts` | 共享层已有 `DocumentCustomXmlPart` 结构校验和 detached XML/XPath runtime；会话以实际节点 key 合并 aliases，写前捕获已匹配 peer，值谓词失配时保留本事务原节点，新匹配胜出。storeItemID 与 XPath 都非空才激活绑定；缺少其中一项保留 dormant 配置，不以 cachedResult 伪造 XML 源。serialize/cache patch 全部准备成功才发布，失败/释放会话不能再提交。DOM handle 不进入领域快照。当前使用有界主线程浏览器 XPath 1.0；未实现或证明 Worker/WASM 的 XML 路径与性能。 |
| 共享提交接线 | `packages/office-studio/src/api/document-content-control-commands.ts`、`packages/office-studio/src/api/document-editor-commands.ts`、`packages/office-studio/src/api/office-studio.ts` | 已静态核对普通输入 prepare → mutation → control finalize → bindings → scoped permission → 发布；表格 whole-cell 输入拒绝内部 active ID。显式 value/checked/list/reset 与属性更新使用注册命令和同一候选历史。内容历史重放在发布前复查当前权限并克隆候选；XML-only 与绑定属性变化不能仅依赖正文差异判断权限。 |
| 数据源公开扩展与包验证 | `packages/office-studio/src/api/document-custom-xml-commands.ts`、`packages/office-studio/src/api/facade.ts`、`packages/office-studio/src/features/document/services/document-custom-xml-validation.ts`、`document-custom-xml-candidate.ts` | `FDocument.getCustomXmlParts()` 返回 detached 部件；`setCustomXmlParts(parts, options)` 返回 `this`，注册 `doc.command.set-custom-xml-parts` 返回 boolean。完整数组、XML/properties/OPC 关系、revision/history generation、编辑权限及激活绑定引用的 store 均校验。constructor 与通用候选边界同样校验；任意 mutate 改 XML 或 store/XPath/prefixMappings 时从 XML 刷新 peers，不能绕过映射校验。显式 setter 可强制刷新同值 XML，XML/缓存/关联内容仍合入同一权限与内容历史候选；尚未运行验证。 |
| 表示与属性验证 | `packages/office-studio/src/features/document/services/content-control-presentation.ts`、`packages/office-studio/src/features/document/services/content-control-properties.ts` | Unicode 码点/字体、列表唯一值、严格 ISO、日期 token/引号/转义/短月名/weekday、Intl 日历映射；明确四位年 Gregorian mask 反解析和格式化回验，支持 locale 月名/星期大小写，拒绝猜 yy 世纪或非 Gregorian 显示输入。runtime 属性 DTO 只允许可保存数据；属性差异提交和字段/位置冲突检查。 |
| 命令契约 | `packages/office-studio/src/features/document/commands/content-control-command-types.ts`、`packages/office-studio/src/features/document/commands/content-control-command-state.ts` | 复选/日期/列表浮层和 Ribbon reset 共用 typed 值命令；picture 与 repeatingItem state 使用同一宿主/权限上下文及真实服务的锁、section flag 与范围判断，返回 available/reason。普通 reset 只允许 text/richText/date/dropDownList/comboBox/picture 的非空 prompt，已经 showingPlaceholder 时不可再执行；wrapper lock 仍允许值/reset/picture 修改并保壳。复制/删除项目的自身 wrapper lock 则按既有服务拒绝，删除还检查后代锁。 |
| 几何 | `packages/office-studio/src/features/document/layout/content-control-geometry.ts` | 使用段落或 cell ID、UTF-16、已有 glyph/line DTO；换行、空范围、页复制、裁剪、图片块边界和 parent depth 排序。完全空裁剪不产生可见 fragment，反向命中选择最内层。 |
| Canvas scene | `packages/office-studio/src/features/document/rendering/content-control-scene.ts` | boundingBox 与真实开始/结束括号，hidden 保留模型/命中但不画边界；使用现有 RenderScene 供项目 Canvas 渲染器处理缩放/滚动/裁剪。复选符号作为正文绘制。 |
| 交互控制器 | `packages/office-studio/src/features/document/controllers/content-control-interaction.ts` | delayed request 校验完整源模型、宿主和输入会话。SelectionForId / ParentSelection 从 textTargets 解析完整 table 点并校验 UTF-16；显式 parentId 权威，缺 ID 可解析最小严格包含范围，排除 synthetic story lock。value edit selection helper 复用合法范围 helper；已卸壳 temporary 先核对完整段落的绑定 peer 编辑，再按可核对单段落差异恢复 collapsed end caret。reset 的 XML scalar 为 `''`，来源正文则使用实际 prompt，避免把 prompt 当成绑定值。同步命令返回后克隆有效会话为 focus token；下一 tick 核对 unit/host/revision/generation/当前选区，再只聚焦当前输入桥。invalid / 多次 revision 推进 / 宿主删除 / 失活拒绝。 |
| 可替换 UI | `packages/office-studio/src/components/document/DocumentContentControlOverlay.vue`、`packages/office-studio/src/components/document/DocumentContentControlPanel.vue` | Checkbox / DatePicker / Select / AutoComplete / Form / FormItem / Button / Tooltip / Popover / Modal / Input / ColorPicker 均使用 antdv-next；控件配色取统一 ConfigProvider token。属性 Panel 接收可选 `readOnly`，Form disabled、显式子控件 disabled、草稿修改处理器与保存前后守卫共用这个上下文；异步验证后再次检查 IME 合成状态，取消保持可用。没有新增原生输入桥。 |
| 画布装配 | `packages/office-studio/src/components/document/DocumentCanvasViewport.vue` | geometry/scene/overlay 接入；IME preview 使用已有 executeDocumentCommand 和同步服务，准备完整 placeholder 选区并检查内容锁/权限。捕获 explicit activeControlId、source selection、host/session，预览卸壳不改变真实控件身份，compositionend 仍走共享命令。公开 `executeContentControlCommand(request)` 复用 overlay dispatch，并在 nextTick 核对结果 focus token；token 在命令返回后被动读取会话，不为旧 source 续期。 |
| Ribbon / Workbench 接线 | `packages/office-studio/src/components/DocumentWorkbench.vue`、`packages/office-studio/src/components/document/document-ribbon.ts` | reset 经 Canvas 同一事件进入注册命令；“选择上层控件”使用同一 host 的完整 parent range，版本不变且 generation 只推进一次才设 active ID/滚动。图片弹窗捕获 API/source/selection/host/session，各上传阶段和提交时复查，changed 后解析新合法锚点；no-op/取消不重新选择。普通重复项直接调用既有注册命令，携带 expectedRevision/inputGeneration，按提交后新增 ID 或合法 caret 恢复。所有结果发布后复查 API/revision/generation/host/选区；观察者较新状态不被旧 active ID 或滚动覆盖。现有 focus token 和输入监听负责延迟焦点，不给陈旧会话续期。插图弹窗使用 antd Form/Input/Upload/Button/Alert/Modal，afterClose 绑定为 prop；主题和真实 IME/焦点未验证。 |

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

命令 request 在以上字段之外携带 typed `command`，允许 `doc.command.set-content-control-checked`、`doc.command.set-content-control-value`、`doc.command.select-content-control-item` 和 `doc.command.reset-content-control-placeholder`。浮层先发 selection，等待 Vue 更新，再重新捕获输入 session；Ribbon reset 则在真实按钮动作开始时捕获当前 source/session，并调用 Canvas 暴露的 `executeContentControlCommand(request)`。不把选择之前的 generation 用于确认。初次 selection 可以没有输入 token，有 token 时仍核对；有 API 的 command 必须有有效 token。选择、源模型或 host 变化、控件删除、权限拒绝、输入不可用、正文/宿主合成输入、失活和弹层锚点离开可视范围都会取消 pending editor；浮层内部 IME 合成期间保留草稿并禁止确认。确认后由 Canvas 和 Workbench 再检查源模型、宿主、权限和 session，补 `unitId` 后进入项目注册命令。Workbench 提交完成后再检查预期 revision 和重新捕获的 generation；观察者重入改变选区或再提交内容时，不恢复旧请求的选区。`DocumentRibbonContext.selectedContentControl` 新增可选 `resetPlaceholderDisabledReason`，来自共享能力的拒绝原因，不是另一个状态引擎。

属性面板保存读取最新 API 控件，只有用户真正改变的字段才进入 patch。当前值改变但未改动的属性不会被旧表单覆盖；位置/类型或所提交字段已变化会拒绝保存。候选式发布与全模型深克隆增加复制成本；尚无大样本测量，不能称为性能优化或内存达标。

父区导航的内部 helper：

```ts
declare function documentContentControlSelectionForId(
  model: DocumentModel, controlId: string, host?: DocumentAuxiliaryBlockHost,
): DocumentSelection | undefined
declare function documentContentControlParentSelection(
  model: DocumentModel, controlId: string, host?: DocumentAuxiliaryBlockHost,
): { controlId: string; selection: DocumentSelection } | undefined
```

helper 只返回逻辑点，不把 host 写进 DocumentSelection；调用端在输入会话中保留 host。显式父必须位于同一段落并包含子范围，失效 ID、跨内容区、字符中间或越界范围返回 undefined，不向其他候选偷偷回退。选择无内容提交，复用 `onSelection` 与 `applyInputState`；提交后被动读取会话，不为观察者的陈旧/失效状态重新续期。本项目保持 Canvas 输入焦点，属性由 Ribbon“属性”入口打开；S19 的 properties 浮层模式和同范围隐式顺序仍为必须继续对齐的缺口。

普通输入候选接口：

```ts
interface DocumentContentControlEditOptions {
  origin: 'text-input' | 'control-value'
  controlId?: string
  placeholderText?: string           // 仅显式 reset；允许绑定 fallback 的空 prompt
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

本记录写入时已静态核对：显式 setValue / setChecked、Canvas IME 候选预览和共享普通输入均调用准备/同步服务。`api/document-editor-commands.ts` 已在独立候选中串联 prepare、正文 mutation、control finalize、bindings 与全部变化目标权限；`api/office-studio.ts` 提交该结果选区并使用统一内容历史。`DocumentWorkbench.vue` 在预期 revision/generation 与 host guard 后调用 `documentContentControlSelectionAfterValueEdit(request, currentModel)`，既可选择存续控件也可发布已卸壳 temporary 的可核对 collapsed caret；Canvas 使用结果会话的 focus token。非空边界 inside affinity 和完整删除的合法 active ID 已透传；所有路径仍未运行。

persistent 控件完成内部输入后，其新 caret 可正好位于 endOffset；Workbench 已在 checked commit 后核对返回范围并保留实际存续控制区的 active ID，temporary 已移除则清掉 ID。普通 `onSelection` 清旧 ID，不能从任意新 caret 位于边界就猜内部身份。该连续输入与完整删除后再输入已有浏览器草稿，接线已落盘但未执行。

绑定来源与事务协议：

```ts
interface DocumentContentControlBindingOriginTextEdit {
  startOffset: number
  endOffset: number
  text: string
  tableCell?: { tableId: string; cellId: string; row: number; column: number; order: number; value: string }
}
interface DocumentContentControlBindingSource {
  readonly control: DocumentContentControl // detached snapshot，不保留 DOM/旧候选引用
  readonly key: string
  readonly value: string
}
interface DocumentCustomXmlBindingSession {
  applySources(sources: readonly DocumentContentControlBindingSource[]): readonly DocumentContentControlBindingSource[]
  bindings(): readonly DocumentCustomXmlControlBinding[]
  refresh(): readonly DocumentCustomXmlControlBinding[]
  commit(): void
  dispose(): void
}
```

这些是 feature 内部协议，没有新增公共 index 导出。`originTextEdits` 一旦提供始终校验，即使此次无需恢复已消失的 source。表格 DTO 的 `value` 是真实格内容，`text` 可包含 materialize 段落需要的 LF；用 UTF-16 edit range 重建完整宿主，不依赖重复文本的 longest-prefix diff。已经预先卸壳的格内 temporary source 使用对应 cell 语义值；显式 scalar 命令可传 `sourceControlId/sourceValue`，普通输入可传 finalizer 的 detached removed controls。

有 `customXmlParts` 时，会话按大小写折叠 store ID 找实际部件，并用当前 XPath 1.0 首个匹配节点作为标量目标；单个 leaf text/CDATA 与其唯一父叶元素共用 node key。storeItemID 和 XPath 均非空的绑定才激活；不完整配置保留 dormant，不因为其 cachedResult 声称已更新原生数据源。激活 source 的来源冲突、相互覆盖节点、无效 XML/XPath、缺 store/源 node、后续锁/权限或无效列表值全部拒绝候选。原本未匹配的非 source 保留内容和缓存，不造节点，相关 Office 行为未核实；已匹配后失配的别名仅在本事务保留写前节点。尚无 `customXmlParts` 的 legacy snapshot 只执行严格 raw-key peer 同步和 `dirty:true`，不能宣称已经写 XML；显式配置空部件列表时仍激活的 source 因缺 store 拒绝。constructor/通用 mutate/setter 共用候选验证；XML 或映射字段变化先刷新关联内容，再以同一权限和历史原子提交。

本项目列表保存值权威、显示标签不反推 ID；ISO 日期独立于 mask，checkbox 只接受 `true/1/false/0/空`，保留节点原有 `1` 字符串而不经 peer echo 改为 `true`。S15 的来源 Map 覆盖与宽松布尔 fallback 未照搬，本项目拒绝冲突和非法 scalar，属于可观察差异，尚待真实对照。reset 第三参仍默认 false；仅绑定空值恢复使用 true，按 S15 nullish fallback 保留 `''`，日期/列表 prompt 不被当成 ISO/项值，reset 的实际变化遵守 temporary 与 wrapper lock。

`refreshCustomXml:true` 忽略 UI sources，以替换后的 XML 刷新当前映射；缓存及 XML-only 改动也核对所属段落/story/entity 权限。普通 editor 使用 binder 的精确 `textEdits` 变换原始选区；显式 temporary value 命令保持既有 boolean 签名，snapshot helper 校验完整段落和存续 peer 的实际坐标后恢复 caret。helper 内部第六参 `sourceText?: string` 仅提供 reset 来源实际 prompt，关联 scalar 仍为 `''`。异 key 的 temporary peer 已无最终锚点、交叉或多层级联无法完全重建时不恢复旧焦点；完整精确 trace 通道仍需实现。

数据源扩展的实际签名已静态核对：

```ts
FDocument.getCustomXmlParts(): DocumentCustomXmlPart[]
FDocument.setCustomXmlParts(parts: readonly DocumentCustomXmlPart[], options?: {
  expectedRevision?: number
  expectedHistoryGeneration?: number
}): this
// doc.command.set-custom-xml-parts: { unitId, parts, expectedRevision?, expectedHistoryGeneration? } → boolean
```

setter 默认捕获当前 revision/history generation；getter 先检查查看权限再返回 detached 部件。fluent 返回值不能用作 changed 标志，草稿通过模型版本和内容历史检查实际变更与同值 no-op；没有向浏览器 evaluate 结果返回 Facade 对象。该命令是本项目本地公开扩展，不增加固定版 Univer 的 Ribbon 根入口统计。

## 历史与持久化

最终业务命令通过 `OfficeStudio` 的文档提交和按文档隔离的内容历史执行，支持 BeforeCommandExecute 取消。普通选区、展开浮层与取消草稿不写内容历史。服务的正文/属性/相邻范围在同一个候选内提交；表格内控件从权威 model 重读 paragraph，不继续使用发布前的旧 table/block 引用。

`FDocument.save()` / 重新构造 `OfficeStudio` 是内部模型快照闭环，**不是 DOCX 文件往返验证**。本次主任务另行扩展项目 Rust parser/writer 的 SDT、story 和 note 数据链；本记录不能把这些尚未运行检查和文件样本的在途改动标为已经保真。`rawSdtPr` / `rawSdtEndPr`、未知 metadata、dataBinding 和内容控件范围保留于项目模型，XML 未知语义和跨结构写出仍须用 Office 样本验收。

## 未运行的测试草稿

| 文件 | 有意义的断言范围 | 状态 |
| --- | --- | --- |
| `packages/office-studio/tests/content-control-presentation.contract.test.mjs` | 复选字符/字体和相邻范围、补充平面 Unicode、四个预设、wrapper/content/祖先锁、LF/UTF-16、日期名称/字面量/非法参数、列表 CRUD/顺序/当前值删除、combo 自由值、调用者隔离、runtime DTO 原子拒绝、属性冲突和 permissionRules、table cell 真实正文 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-geometry.contract.test.mjs` | 换行/空范围/end caret、emoji、同范围嵌套、cell 目标、空/部分裁剪、重复页 affinity、image 几何、hidden/tags 和 Canvas 缩放/滚动 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-editing.contract.test.mjs` | temporary/no-op/纯格式、祖先/内容/wrapper 锁和重接、预先删壳不能绕内容锁、完整 placeholder 首输、empty/首尾 inside affinity、两个 planner 的非空边界内部/外部换行与完整删除空壳语义、UTF-16、checkbox glyph/font 及后续非法控件的原子拒绝、Gregorian 反解析/Last Known Date、combo 自由值/下拉拒绝、table cell/range/matrix、body/header/footer/note 隔离和兼容字体投影、跨 host/重复 ID、边界零长实体 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-interaction.contract.test.mjs` | table 的 child/parent 完整坐标与 detached 结果、显式父优先/失效不回退、legacy 最小严格包含、同范围显式关系边界、跨段落/host/越界/UTF-16/synthetic story lock 拒绝和 header/note 内 table；reset 支持类型/非空 prompt/已显示占位、内容/祖先/wrapper 锁/权限/宿主；temporary prompt caret 与存续 wrapper-locked 范围；卸壳后带前后正文的 emoji end caret、存续/no-op 范围、周围文字改变拒绝、结果会话复制隔离、no-op/cancel、观察者选择当前光标、陈旧 revision/generation、输入无效、文档/host/disabled/辅助块删除拒绝 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-history.contract.test.mjs` | 真实 Facade 注册命令与 OfficeStudio CommandExecuted 事件、单事务、撤销重做/内部保存重开、普通 temporary/placeholder/date/combo、same-value revision/savepoint/redo、不合法下拉/取消/段落与 note entity 权限、历史隔离、header/footnote 普通输入及字体兼容投影 | 草稿；未运行，不能用静态接线推断通过。 |
| `packages/office-studio/tests/content-control-command-integration.contract.test.mjs` | 普通 command prepare/finalize、真实段落/cell 与辅助 host、active ID/边界、temporary 与 placeholder、下拉/锁拒绝、取消和单候选历史 | 草稿；未运行。 |
| `packages/office-studio/tests/content-control-bindings.contract.test.mjs` | raw key/语义 scalar、六类 host、单历史/内部保存重开、same-value redo、冲突/后续权限回滚、替换与代码源编辑、temporary source、nullish 空 prompt、date/list reset 和最后有效 ISO、图片 reset 后续失败不泄漏恢复；cell/range/matrix 的真实来源、空/对象格/相同 offset 的 DFS 顺序/分隔符/负面 DTO、materialize cell 的字段缓存与 locked 回滚；显式 caret 校验实际存续 peer 显示/坐标 | 草稿；未运行。 |
| `packages/office-studio/tests/document-custom-xml-bindings.browser.test.mjs` | 浏览器真实 XPath、命名空间/leaf text alias、header/temporary、单历史/内部保存重开、布尔 lexical `1`、空 XML/prompt、无效源/冲突/后续锁/权限/列表回滚、cache-only 权限、值谓词失配/重匹配、不匹配非 source 保留、显式 temporary alias/不同节点显示/清空及 reset 来源 prompt caret、XML-source refresh、fluent setter、revision/history no-op 与 getter 隔离 | 草稿；未运行。已静态核对最终 setter/getter 签名；这里未启动浏览器。 |
| `packages/office-studio/tests/document-content-control-workbench.browser.test.mjs` | 真实 Workbench + OfficeStudio + antd 复选/日期/列表/组合输入、table 上层选区完整坐标/模型历史不变/header 同范围父 host/焦点恢复/DocumentInputStateChanged 观察者改写保留新 generation；temporary 属性/普通输入/卸壳后继续输入和 emoji caret、Ribbon reset 与同段落 XML alias 的 combo/reset/取消/观察者选区及内容重入/能力状态、persistent 合法内部连续输入/完整删除后再输入、same-value date、完整 placeholder 合成 IME/重复 input 抑制/取消/旧 DOM 节点、普通 header/note 输入、键盘/焦点、弹层过期、观察者重入、权限/属性只读、自定义复选、wrapper lock、classic/grid 明暗主题、内部保存重开；模型同步采用 Vue sync watcher，命令观察使用 CommandExecuted，额外断言 listener errors 为空 | 草稿；未运行。 |
| 既有 `content-controls.contract.test.mjs`、`content-control-boundaries.contract.test.mjs` | 修正复选插入会缩为持久化单字符后的 fixture offsets；table 引用从权威模型重读；祖先锁夹具在创建 checkbox 后锁定 | 修改草稿；未运行。 |

实际执行仅有文件读取、搜索、公开规格/工具文档查阅、改动读取和编辑落盘。没有运行测试草稿、类型或构建；没有浏览器、系统剪贴板、Windows 真实输入法、DOCX/Word 样本、两渲染后端或性能证据。

## 剩余缺口与完整目标

- temporary 实际内容编辑、reset、共享普通文本提交、显式卸壳后的可核对 collapsed caret 已接线；复杂多 key/nested、异 key alias peer 同时卸壳及精确 edit trace 仍需实现，完整浏览器/历史闭环未验收。
- 日期、复选、组合和下拉的共享普通输入、value / checked / dateValue / showingPlaceholder / dirty、边界 inside affinity 和合法 active ID 连续输入已有接线；取消/观察者重入、真实 IME 和全部输入路径仍未验证。
- 范围模型已有单段落 UTF-16、完整 block range 和连续 table row range。块/行 richText、repeatingSection/item、group 的结构插入、复制/删除和普通输入约束已落盘；部分跨段选择、跨 cell 控件、独立 whole-cell/whole-table 重复容器、全部结构组合与 Office 样本仍有必须继续处理的边界。新增草稿不表示结构功能已验收。
- “选择上层控件”已接完整 table 点、辅助 host 和输入版本守卫；本地控件排序不保留 S19 同范围 customRanges 原顺序，无 parentId 的同范围导航和父区 properties 浮层模式仍未实现。table/header/观察者改写草稿未运行，滚动定位、焦点与真实浏览器交互未验收。
- docPartObj 的真实部件库/部件复用/唯一行为尚未实现。基础 scalar custom XML 读写、XPath 1.0 aliases、缓存/权限/整批候选同步及独立数据源 getter / fluent setter / 命令已落盘；完整 Office 关系/来源保真、多节点重复绑定、跨结构控件和所有闭环仍未验收，不能将属性保存或本轮基础实现当作全部能力完成。
- 内容控件与 native fields、嵌套图片、linked story、权限变更/历史跳转、取消后的焦点等需要真实集成场景；本批草稿的覆盖不代表这些场景已通过。
- 非 Gregorian 日历、地区历法/时代格式、全部 Office 日期格式与真实 Windows 字体/IME 未验收。
- 大量控件的几何扫描、候选复制、header 重复页、图片内存和 XML 主线程任务未测量。XML baseline 是有界主线程平台适配，尚未实现或基准验证 Worker/Rust/WASM 路径；额外遍历、clone、序列化和整段文字重建都不能称为优化。未记录首屏、滚动、编辑、缩放、长任务、Worker/WASM 复制字节/次数、峰值内存等指标，性能状态保持未验收。
- DOCX 读写、Word 文件重开、未知 SDT XML 保真和 Pro 对照演示均未执行。本记录不缩减完整功能范围，不把缺口改成永久禁用、提示消息或展示入口。
