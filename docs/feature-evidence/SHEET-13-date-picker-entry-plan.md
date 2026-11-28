# SHEET-13：日期控件实施与实现记录

日期：2026-11-28。状态：日期、列表控件及规则面板的生产实现、接线和测试草稿已写入，未验收；详细实际接口见 [完整控件记录](SHEET-13-validation-cell-input-2026-11-28.md)。原方案的逐批文件/行数确认约束已由当前 AGENTS.md 的持续授权替代，本批按完整范围直接推进。以下历史方案保留设计依据，当前文件归属与接口以实际记录为准。

<a id="date-validation-cell-input-2026-11-28"></a>

目标是在现有 Sheets Canvas 选中/编辑主格上提供真实日期和时间选择，复用本项目 serial、输入会话、权限、数据验证和内容历史。快捷入口仍将规则写入全部选区；日期控件只编辑其绑定的目标格。

## 固定来源与可观察行为

固定版本：Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

- [DateValidatorView](../../../univer/packages/sheets-data-validation-ui/src/views/validator-views/date-validator-view.ts) 将 DATE 映射到日期 dropdown，并注册 [DateShowTimeOption](../../../univer/packages/sheets-data-validation-ui/src/views/components/DateShowTimeOption.tsx)。后者只更新 `bizInfo.showTime`，保留其他扩展项。
- [dropdown-manager.service.ts](../../../univer/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts) 的日期保存处理写入数值 serial、NUMBER 类型和日期格式，清除原公式/富文本；STOP 验证失败不写入、不关闭。纯日期选择保留已有隐藏时间，空格从午夜开始；只改时间保留 serial 0 / 60 的伪日期锚点；超出可表示范围的原 serial 仍可作为数字编辑。
- [DateDropdown.tsx](../../../univer/packages/sheets-ui/src/views/dropdown/DateDropdown.tsx) 管理日期/时间草稿、dirty 与明确确认，`onChange` 返回 false 时保持打开，无实际改动时保留原值。
- [DateValidator](../../../univer/packages/sheets-data-validation/src/validators/date-validator.ts) 使用 workbook date system / locale；无 operator 表示任意合法日期，showTime 参与公式归一化。

## 现有接口依据

- [serial-date.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/serial-date.ts) 已有 `IExcelDateTimeParts`、`excelSerialToDateTimeParts`、`excelDateTimePartsToSerial`、`getDateSystemMaxSerial` 和 `EXCEL_1904_OFFSET = 1462`。1900 serial 0 是 1900-01-00，60 是虚构 1900-02-29；1904 serial 0 是 1904-01-01。转换保留 fractionalSecond，使用 Gregorian 算术；不借 native Date 时间戳差计算 serial。
- [sheet-number-format-edit.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-number-format-edit.ts) 已有 `sheetNumberEditText` / `parseSheetNumberEdit`、原值未变时保留精度、隐藏日期/时间处理；[date-time-edit-pattern.ts](../../packages/office-studio/src/features/sheet-workbench/number-format-engine/date-time-edit-pattern.ts) 已有编辑展示模式。复用这些入口，不再实现通用日期文本 parser。picker 的数值提交直接使用 parts/serial；输入展示模式不等同于保存格式选择。
- [sheet-cell-edit-session.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-edit-session.ts) 已有 `createSheetCellEditSession`、`getSheetCellEditConflict` 和 date1904 解析。会话绑定 sessionId、文档、工作表、主格、日期系统、region、currentYear 和源内容/格式/富文本；日期控件沿用这份绑定。
- [sheet-cell-input-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-input-command.ts) 已有公共目标/会话输入 reader、冲突消息与 `SheetCellInputCommandResult`（accepted / changed / issues / reason）。[sheet-cell-input-commands.ts](../../packages/office-studio/src/api/sheet-cell-input-commands.ts) 已有 `commitWorkbook`、`preserveStyleInheritance`、视图发布和 `SheetValueChanged` 事件契约。
- [sheet-cell-native-input.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-native-input.ts) 仅在组件库拥有的输入上捕获原生事件，按 sessionId / hostId 绑定；不建立第二个 DOM 输入。日期选择继续由 Ant 控件提供输入。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 已有 `cellInputContext`、workbenchActive / IME 状态、`ribbonEditorOwner`、`keepCellEditFocus` 和 `restoreCellInputFocus`。`keepCellEditFocus` 当前识别所属 Ribbon 控件及 Select / Popover / Dropdown / ColorPicker 弹层；新增日期控件需把所属 picker 弹层纳入同一焦点判断。
- [antdv-locales.ts](../../packages/office-studio/src/i18n/antdv-locales.ts) 已配置 Dayjs / Ant 日期 locale；[package.json](../../packages/office-studio/package.json) 已声明 dayjs。新控件复用统一 ConfigProvider、主题 token 和翻译入口。

## 历史生产文件方案

以下为实现前的路径与行数估计，仅保留历史设计依据。实际已拆为共用 widget planner、日期/列表 planner、两个 owner hook、统一命令适配、Canvas 标签/定位及 Ant 控件；日期 API 的实际文件为 `api/sheet-validation-input-commands.ts`，规则事务为 `api/sheet-data-validation-commands.ts`。当前改动清单见完整控件记录；OfficeStudio、根导出和 locale 注册由主任务协调。

| 文件 | 预估行数 | 具体改动 |
| --- | --- | --- |
| 新增 `features/sheet-validation/validation-date-input.ts` | 190–260 | 同步日期草稿和数值编辑计划；复用现有 serial / 日期文本显示；处理 date/time 分别修改、0/60、1904、隐藏时间、fractionalSecond、未改动和不可表示 serial；复用既有模型、权限与数据验证函数 |
| 新增 `api/sheet-date-input-commands.ts` | 100–145 | 注册日期提交内容命令；在命令入口复用公共目标/会话 reader，读取日期专用 change 参数；通过既有冲突与验证结果形成一次 `commitWorkbook`，同时提交值/类型/日期格式并发布事件 |
| 新增 `components/spreadsheet/SheetDateValidationPicker.vue` | 120–170 | 使用 Ant Form / DatePicker / TimePicker / InputNumber / Button；受控草稿和明确确认；特殊日锚点与不可表示 serial 通过数值输入保留；失败保留弹层和草稿；使用统一 token 与 locale |
| [components/spreadsheet/SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) | 45–70 | 日期规则触发命中和受控覆盖层；按既有单元格布局定位；滚动、缩放、冻结变化更新位置；事件绑定当前 session 和 editor owner |
| [components/ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) | 45–70 | 查询当前编辑/选中主格的 date 规则、原 serial、日期系统和格式；创建/复用输入会话；路由统一日期提交；accepted 后结束该会话并恢复画布焦点；切工作台/目标/文档取消所属弹层 |
| [api/office-studio.ts](../../packages/office-studio/src/api/office-studio.ts) | 5–10 | 注册并随 studio 释放日期内容命令，不迁移其他业务 |
| [components/spreadsheet/SheetDataValidationPanel.vue](../../packages/office-studio/src/components/spreadsheet/SheetDataValidationPanel.vue) | 20–40 | 日期类型增加 Ant Checkbox 显示时间选项，读取/更新 `metadata.showTime` 并保留其他 metadata；继续使用原 replace-data-validation 命令和历史 |

## 日期提交协议与边界

命令参数绑定 unitId / sheetId / row / column 和现有 SheetCellEditSession；日期修改使用有区分的 change：unchanged、date parts、time parts、numeric serial。日期系统取会话及当前 workbook，不从控件显示用的 Dayjs 推断。公共目标/会话字段在入口由既有 reader 读取一次，日期 change 同样只在入口读取一次；内部计算接受已确定类型。

UI 只持有草稿和控件适配值。权威值仍是 SpreadsheetCell 的数值 serial；特殊日期的近似日历锚点只用于控件显示，时间修改和原值未改确认保留真实 serial。新计划复用现有日历算术、冲突、保护、验证、格式和内容历史，不引入第二套日期引擎。单格日期换算为 O(1) 同步计算，不新增 Worker/WASM 任务或跨线程缓冲区。

成功提交一次发布数值、类型与格式；accepted false 时不产生内容历史并保持草稿。accepted true / changed false 的未改确认关闭所属弹层，保留原值及格式，不再次进行文本写入。撤销/重做还原同一内容事务；原生保存/重开保留 serial、numberFormat 和 showTime metadata。

## 后置验证范围与实际草稿

- 新单元/命令契约草稿：1900 serial 0 / 59 / 60 / 61、1904 epoch / 1462 差值、纯日期保留隐藏时间、时间修改保留伪日期、fractionalSecond、空格午夜、no-op、不可表示 serial、STOP 拒绝、受保护目标、会话冲突、一次历史及原生重开。
- 已补 [真实 Canvas/Studio 浏览器草稿](../../packages/office-studio/tests/sheet-validation-cell.browser.test.mjs) 与 [fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html)，包含日期/时间/数字/duration、面板 showTime、日期系统、冻结缩放、关闭回焦、主题和合成 composition；这些草稿尚未运行，不能仅用 metadata 或合成事件冒充日期控件/Windows IME 验收。
- XLSX 文件样本和浏览器运行、类型、前端、Rust/WASM 及性能检查统一后置。取得实际证据后更新 SHEET-13 状态；当前全部保持未验收。
