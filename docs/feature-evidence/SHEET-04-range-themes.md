# SHEET-04 范围主题、目录与区域样式来源

记录日期：2026-11-26。固定参考 `E:\rust_test\univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**前端实现与草稿已落盘，未验收**。续接[默认样式与稀疏来源](SHEET-04-default-style-inheritance.md)和[公开样式读回](SHEET-04-public-style-readback.md)。按用户安排测试后置，本批没有运行测试、类型检查、构建、浏览器/dev、Cargo/Rust/WASM 或性能基准。

## 固定依据与入口清点

以下路径均相对固定参考根，已静态读取；没有引用不存在的 `controllers/range-theme.controller.ts`。

| 来源 | 核实的契约 |
| --- | --- |
| `packages/sheets/src/models/range-theme-util.ts`：`RangeThemeStyle` | 九区段作者配置、getter/setter、JSON、按行列偏移与末行/末列/条纹状态组合。有效空 mask 返回空样式，无 mask 返回 null。 |
| `packages/sheets/src/models/range-themes/default.ts`、`build-in-theme.factory.ts` | default、default-last-row 与 30 个颜色配置，共 32 个内置主题；最后一行边框和各区段保持独立。 |
| `packages/sheets/src/models/range-theme-model.ts`：`SheetRangeThemeModel` | 内置/自定义目录、RTree 规则、隐藏/零高度/过滤行条纹、样式拦截与 `SHEET_RANGE_THEME_MODEL_PLUGIN` 资源。 |
| `packages/sheets/src/services/range-theme.service.ts` | 目录注册、规则应用/移除、已应用主题及内置名称查询。 |
| `packages/sheets/src/commands/commands/add-worksheet-range-theme.command.ts`、`delete-worksheet-range-theme.command.ts` | set/remove 固定命令与按工作簿的撤销重做。移除命中规则的整个应用范围。 |
| `packages/sheets/src/commands/commands/register-range-theme.command.ts`、`unregister-range-theme.command.ts` | 注册/取消注册命令、作者 JSON 与内容历史。固定 unregister handler 调用 Register mutation 的代码事实单独保留，本地按取消注册语义实现，不复制该调用。 |
| `packages/sheets/src/facade/f-workbook.ts`、`f-range.ts` | 目录、builder、注册、`useThemeStyle`、`removeThemeStyle` 与 `getUsedThemeStyle`；null/undefined 走移除。 |
| `packages/sheets/src/models/__tests__/range-theme-util.spec.ts`、`range-theme-model.spec.ts`、`services/__tests__/range-theme.service.spec.ts` | 区段组合、目录、范围和可见行契约参考；本批没有运行上游测试。 |

固定版该能力以 API 为主要入口。本批新增本地常规 Ribbon 扩展 **1**，固定新增常规根 **0**、专用上下文根 **0**；主题 Select 的 32 个内置项及自定义项、弹窗三个操作按钮均不计作常规根。More 复用同一根，不另增加计数。

| 本地入口 | 标签/组与顺序 | classic/grid 位置 | 状态及命令 |
| --- | --- | --- | --- |
| `sheet.operation.open-range-theme` | 开始 → 工具，group order 40，第 3 项 | classic 在范围保护、清除内容之后；grid row 1、column 3、rowSpan 2，显示标签 | 无工作表、选区不可编辑、输入法组合中或已有范围主题弹窗时禁用；打开真实 Ant Modal，应用/移除共用下列固定命令。 |

当前本地 Sheets 常规根从历史 **68（29/6/11/17/5）** 增至 **69（30/6/11/17/5）**；Home 分组为 clipboard 5、font 11、layout 6、number 5、tools 3。固定 OSS 仍为 63（30/9/12/6/6）。新增范围主题入口不替代 SHEET-18 的结构化表格，也不能用总数相同证明分组或功能完整。

## 模型、来源与真实行为

[domain/model.ts](../../packages/office-studio/src/domain/model.ts) 保存 workbook `rangeThemes`、worksheet `rangeThemeRules`（稳定 ID、themeName、物理 CellRange 与扩展）。九区段为 whole、first/second/last row、first/second/last column、header row/column；[sheet-range-theme-style.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-style.ts) 负责稀疏短字段/本地样式归一化、builder 与脱离模型 DTO。[sheet-range-theme-presets.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-presets.ts) 保存冻结的 32 个固定配置。

[sheet-range-theme-resolver.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-resolver.ts) 按真实物理范围与可见行解析。组合顺序为默认/行列 → 范围主题 → cell theme → 直接引用/reset/patch；边框逐边合并，显式 false/null 不被较低层覆盖。手动隐藏、零高度、筛选、表筛选及大纲折叠合并成紧凑 spans，不为百万行隐藏范围逐行分配集合或单元格。末行样式跟随最后可见行；隐藏表头不迁移到下一行，列条纹仍按物理列。

[sheet-range-theme-model.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-model.ts) 在赋值前校验全部目录和规则；非法后续项、重复规则 ID、缺失目标工作表或目录名字不一致时不留下部分导入。公开资源导入转换 `rangeInfo.subUnitId` 和 start/endColumn，原始 rule 扩展保存在来源 metadata。特殊名称如 `__proto__` 使用 own key，目录更新保留已有顺序。

行列插删/移动/复制与工作表复制已接范围规则变换；删除耗尽规则，拆分和复制使用独立新 ID。区域样式快照保存已解析主题、原始坐标和可见行相位；缺目录的上层规则保存 `theme:null` blocker，源关闭或注销后复制读取不漏出较低层主题。现有剪贴板、格式刷和填充通过[sheet-style-region.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-style-region.ts) 使用该来源；仍需专门证明各跨模块操作及 Office 重开，不将接线等同全部闭环验收。

## 命令、Facade、权限和控件

[sheet-range-theme-command.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-command.ts) 提供 DTO 与共享 planner；[sheet-range-theme-commands.ts](../../packages/office-studio/src/api/sheet-range-theme-commands.ts) 注册四个固定命令：

| 命令 ID | 持久内容与权限 |
| --- | --- |
| `sheet.command.register-worksheet-range-theme-style` | 写入脱离 builder 的目录配置；检查所有使用该名字的工作表/规则范围权限，捕获 model 身份和 revision，拒绝序列化重入。 |
| `sheet.command.unregister-worksheet-range-theme-style` | 移除自定义目录，保留 dangling 作者规则；无自定义项为 no-op。重新注册可恢复其显示。 |
| `sheet.command.set-worksheet-range-theme-style` | 应用/替换范围规则；同范围同主题为 no-op；null/undefined 移除当前命中规则。 |
| `sheet.command.remove-worksheet-range-theme-style` | 按名字移除每个选区命中的一条规则；权限覆盖被移除规则的整个应用范围。 |

实际内容修改进入同一 workbook 内容事务、dirty 与撤销重做，关闭无关公式重算；只读/范围保护、无效参数、缺主题、陈旧 revision 均在发布前拒绝。Facade 提供固定目录、create/get builder、注册/注销以及范围 use/remove/read，内部保存/重开保留作者目录与规则。register 替换已有目录的撤销恢复旧配置，这是本地内容历史实现。

[use-sheet-range-theme.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-range-theme.ts) 同时支持 Studio 和独立工作台，捕获目标/内容版本或独立 view signature；准备、发布和回焦重新判断当前宿主。切工作表、文档、Studio 或停用工作台取消旧弹窗；生命周期 generation 防止迟到关闭回调在同 ID 的新激活会话夺焦点。错误保留弹窗及选择；移除时说明整个规则范围，保护尾部不可编辑时保持模型和历史。

[SheetRangeThemeDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetRangeThemeDialog.vue) 的基础批使用 `antdv-next` Modal、Form/FormItem、Select、Button、Alert 和统一主题 token；Canvas/公式输入继续使用现有 Ant 输入桥。弹窗 composition 状态阻止组合输入中提交；关闭后按当前性恢复焦点，其他异步粘贴/填充/格式刷、Ribbon 错误和弹层回焦路径增加范围主题 pending 判断。新增中英资源，其他语言使用英文 fallback。基础批时只有选择器；同日后续[九区段编辑](SHEET-04-range-theme-editor.md)已落盘 Ant 字段、原子保存/应用及 45 个新草稿，仍未验收。

## 差异与未核实项

- 固定 RTree 使用第一个搜索结果，具体可见重叠顺序未核实；本地选择后应用规则优先。exact range 换主题替换旧规则并保留 ID。
- remove/null 每个选区移除一条相交规则的整个范围；多选可以依次命中多条。不是只清选中子范围，也不承诺一次清掉所有重叠规则。
- 固定默认目录优先于同名自定义；本地保持该查找顺序。unregister 保留 dangling 规则，本地不渲染，重新注册恢复。
- builder getters/JSON 返回脱离模型 DTO；`ol` 等未建模字段能存/读，不等于已绘制或行为完整。
- 基础公开 snapshot resource 已导入；全套固定 snapshot 输出、XLSX 默认/范围主题与 theme/tint 映射、Office 独立重开尚未验收。
- [局部移位/剪切续作](SHEET-04-range-theme-structure.md)已接物理/逻辑片段、内部 cut/all 作者规则与稀疏来源，新增 32 个未运行草稿；[共享移位/合并确认](SHEET-02-cell-shift-transactions.md)与[九区段编辑/历史及宿主 guards](SHEET-04-range-theme-editor.md)也已落盘。复杂轴组合和全部复制/格式刷/填充跨模块闭环、动态权限及完整混合状态仍需实施或验证。
- 当前解析/样式命令属于 Sheets feature，Canvas 读取同一派生样式。未改 Worker 控制 DTO、WASM lease 或网络协议；大量规则索引、投影与布局 Worker/Rust/WASM 及取消路径仍待实现/验收。无真实大样本或复制/内存测量，不能声称已优化。

## 草稿与验收状态

| 文件 | 静态场景数 | 计划覆盖 |
| --- | ---: | --- |
| [sheet-range-theme.contract.test.mjs](../../packages/office-studio/tests/sheet-range-theme.contract.test.mjs) | 33（25 纯契约 + 8 Facade） | 目录/九区段/空 mask、稀疏继承、隐藏条纹/footer、重叠、blocker、原子导入、轴/工作表复制、权限、重入、历史和内部重开。 |
| [sheet-range-theme-host.contract.test.mjs](../../packages/office-studio/tests/sheet-range-theme-host.contract.test.mjs) | 7 | 独立工作台一次发布、陈旧内容、切宿主/迟到关闭、不完整规则权限、dangling 移除与准备/发布重入。 |
| [sheet-range-theme.browser.test.mjs](../../packages/office-studio/tests/sheet-range-theme.browser.test.mjs) + [fixture](../../packages/office-studio/tests/fixtures/sheet-range-theme.html) | 6 | 真实 Ribbon/Ant Modal/Canvas、内置/自定义、历史和内部重开、子范围移除、保护/陈旧目标、grid/暗色/Input 与 Escape。 |

合计 **46 个未运行草稿**。静态阅读和链接/计数核对不提供测试通过结论；浏览器/真实中文 IME、两渲染后端、Office 样本、权限组合、完整保存及性能均保持未验收。全部 Sheets / Docs / Pro 范围保留，服务与文件协议增量见[后端需求](../BACKEND_REQUIREMENTS.md)。
