# SHEET-13：下拉搜索/编辑配置、虚拟列表焦点与输入宿主释放

记录日期：2026-11-28，Asia/Shanghai。固定参考为 Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本批接续原线程四份在途文件，生产接线与草稿已落盘，功能状态仍为未验收。

## 入口和固定来源

| 类别 | 来源与契约 | 本地实现 |
| --- | --- | --- |
| UI 配置 | [config.ts](../../../univer/packages/sheets-data-validation-ui/src/config/config.ts)，`showEditOnDropdown` / `showSearchOnDropdown` 均默认 true | `ExcelWorkbench.validationDropdownUi` 传到 Canvas 单元格控件；参数仅改变列表 UI，不修改规则、权限或文档快照 |
| 搜索/编辑入口 | [dropdown-manager.service.ts](../../../univer/packages/sheets-data-validation-ui/src/services/dropdown-manager.service.ts)，创建列表时读取两个配置值；[ListDropDown.tsx](../../../univer/packages/sheets-ui/src/views/dropdown/ListDropDown.tsx)，条件显示搜索框、权限下的编辑入口，搜索显示后聚焦 | 原 Ant Form/Input/编辑 Button 条件呈现；编辑入口仍经过当前 owner 与实际权限；同目标重复 show 保留搜索与输入草稿 |
| 显示/隐藏 operation | [data-validation.operation.ts](../../../univer/packages/sheets-data-validation-ui/src/commands/operations/data-validation.operation.ts)，两个固定 ID、相同位置重复显示幂等 | 复用原 `sheet.operation.show-data-validation-dropdown` / `sheet.operation.hide-data-validation-dropdown`、注册/释放、完整 merge View/Edit 和 owner 逻辑 |
| 单/多选列表 | [ListDropDown.tsx](../../../univer/packages/sheets-ui/src/views/dropdown/ListDropDown.tsx)，按原 options 顺序重新收集选中项 | 复用唯一候选、CSV、内容命令与历史；Ant Checkbox 的 Enter 转到其真实 click/change，Space 沿组件原事件，不另设一条值提交路径 |

分类口径沿用[已有清点](SHEET-13-checkbox-feedback-dropdown-2026-11-28.md)：数据验证四个常驻根、管理两个子项；show/hide 为两个 operation，不是 Ribbon 根。本批常驻、插件扩展、溢出、上下文根增量均为 0；两个配置字段、N 个列表候选和内部按钮不充当功能完成数。

## 实际改动与接口

- [validation-dropdown-ui.ts](../../packages/office-studio/src/features/sheet-validation/validation-dropdown-ui.ts) 定义 `SheetValidationDropdownUiOptions` / `SheetValidationDropdownUiState` 与 `resolveSheetValidationDropdownUi`；未提供的字段默认 true，拒绝对象/布尔字段类型错误，不把字符串 `false` 强转为权限状态。根公共导出由主任务接入。
- [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) 和 [SpreadsheetCanvas.vue](../../packages/office-studio/src/components/spreadsheet/SpreadsheetCanvas.vue) 保留原在途传递，未新增常规根。
- [SheetValidationCellControl.vue](../../packages/office-studio/src/components/spreadsheet/SheetValidationCellControl.vue) 在搜索关闭且列表异步尚未挂载时保留初次焦点请求，实际候选到达后用已有逻辑索引导航聚焦虚拟行；空列表/失败聚焦合法关闭/重试按钮。切换搜索显隐、owner、query、revision 和 IME 会取消旧焦点请求。重新显示搜索按固定源码语义聚焦 Input。
- 搜索输入正在 composition 时，配置改为隐藏仍保留该 Input；composition 结束后清空隐藏 query，恢复完整候选及合法焦点。编辑入口隐藏也不会覆盖搜索草稿；设置和焦点均不进入内容历史。
- 虚拟导航的 `requestAnimationFrame` 和取消使用实际 popup `ownerDocument.defaultView`，以独立 frame 记录保存窗口/id/resolve，避免不同窗口数值 ID 冲突和卸载后未释放等待。
- [use-sheet-validation-cell.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-validation-cell.ts) 为内部 popup 增加 `detach()`。Canvas 控件被裁切、重建或卸载时，销毁已经不存在的输入宿主租约，阻止旧提交、关闭和回焦；普通 hide 在 composition 中继续拒绝。detach 不提交内容、不结束原单元格输入会话、不创建历史，也不能从旧 owner 关闭新的 popup。

所有可替换 UI 保持 `antdv-next`：Form/FormItem、Input、Checkbox、Listy、Popover、Tooltip、Button、Alert、Tag 和进度反馈。仅操作这些组件的真实输入/焦点 DOM 和 Canvas 宿主，不新增原生表单控件或第二份内容模型。主题沿统一 ConfigProvider、既有 theme token。持久模型、命令 ID、历史来源和 Worker/WASM 协议未新增。

本地多选列表沿前段选择后保持打开的交互，固定 `ListDropDown` 会在成功提交后隐藏；该既有差异继续保留，不能宣称内部实现完全相同。虚拟行键盘导航、Enter 和隐藏搜索后的合法焦点属于本项目可访问交互实现，固定源码只确认 options、搜索和点击选择契约。

## 草稿与只读检查

新增草稿均未运行：

| 文件 | 本批新增声明 / 计划 | 当前声明 / 计划 | 场景 |
| --- | ---: | ---: | --- |
| [sheet-validation-dropdown.contract.test.mjs](../../packages/office-studio/tests/sheet-validation-dropdown.contract.test.mjs) | 3 / 3 | 18 / 20 | 默认/非法配置、组合输入宿主 detach 保留内容和输入会话、旧 detach 不影响新目标 |
| [sheet-validation-cell.browser.test.mjs](../../packages/office-studio/tests/sheet-validation-cell.browser.test.mjs) | 4 / 8 | 27 / 53 | classic/grid 的真实 Rust 大列表初次焦点/End/Home/Enter、显隐配置与 composition、单选历史与重开、scroll 裁切旧宿主和程序化重开 |
| [fixture](../../packages/office-studio/tests/fixtures/sheet-validation-cell.html) | — | — | UI ref 配置传入真实工作台；不替换 query、命令、历史或 Worker 算法 |

本批实际只读静态快照：8 个文件、8 段 TS/JS/夹具脚本、3 个 Vue 模板、3 个样式块、241 处本地引用通过；两份 `.mjs` 的 `node --check` 通过，限定 diff 空白检查通过。仅使用 TypeScript parseDiagnostics、Vue 模板解析、PostCSS 语法解析、来源/引用及 AST 草稿计数，没有运行测试 runner、类型检查、构建、dev 或浏览器。

后续 Markdown 本地链接/固定来源检查单独执行，不把前段静态快照用于未覆盖的文件。测试、类型、前端构建、Rust/WASM、浏览器、Windows IME、两种渲染后端、Office 往返和大样本性能仍按 AGENTS.md 后置；真实执行次数为 0。本批不提升 SHEET-13 或其它 65 个功能组的验收状态。
