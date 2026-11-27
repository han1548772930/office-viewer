# SHEET-04 边框命令、合并边段与笔状态（待统一验证）

记录日期：2026-11-23。参考根为 `E:\rust_test\univer`，固定版本 Univer v1.6.0、提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。当前记录仅证明源码核对、实现和测试用例已经落盘；本批未运行测试、类型检查、构建、Rust/WASM 或浏览器，功能状态保持“未验收”。字体和颜色阶段见 [SHEET-04 颜色记录](SHEET-04-color-reset.md)，总清点见 [顶部工作区对照](../UNIVER_TOOLBAR_PARITY.md)。

## 固定源码链

以下路径均相对上面的参考根，不依赖 Univer Server 或商业 SDK。

| 来源 | 已核对的可观察契约 |
| --- | --- |
| `packages/sheets-ui/src/menu/schema.ts`、`menu/border.menu.ts` | START / FORMAT 的边框分体根，`SetBorderBasicCommand`；禁用包含工作簿编辑、工作表编辑、设置样式和范围 EditPoint；Sheets 外或普通绘图上下文隐藏。 |
| `packages/sheets-ui/src/views/border-panel/interface.ts`、`BorderPanel.tsx` | 15 个位置预设、固定顺序、5 列排列；12 个菜单线型；颜色子面板和笔状态订阅。 |
| `packages/core/src/types/enum/border-style-types.ts` | 公开位置别名及 0–13 数字线型；NONE、SLANT_DASH_DOT 不增加菜单线型数量。 |
| `packages/sheets/src/services/border-style-manager.service.ts` | 默认 ALL / THIN / `#000000` / inactive；位置/basic 激活；inactive 的颜色/线型命令只改变笔配置。 |
| `packages/sheets/src/commands/commands/set-border.command.ts` | 外边设置清邻格反向边；不跳过筛除或手动隐藏行；斜线写合并主格，cardinal 边保留物理格边段；当前固定实现只处理 `selections[0]`。 |
| `packages/sheets/src/commands/mutations/set-range-values.mutation.ts`、`commands/commands/__tests__/set-border.command.spec.ts` | 正逆 mutation、按 unitID 隔离历史；合并 `B1:B3` 与邻格 `A2` 的局部边段回归。 |
| `packages/sheets/src/facade/f-range.ts` 的 `setBorder` | 完整范围一次传入 `set-border-basic`；`setBorder(type, style, color?)` 支持链式调用。 |

## 入口清点

常驻根入口 1 个：`sheet.menu.borders`，Home / 字体组 classic 第 11 项、grid 第 2 行第 5 列。窄视口移入 overflow 是同一入口。主区使用最近已接受的位置和笔；面板位置键如下，固定版对应位置命令为 `sheet.command.set-border-position`，本项目面板保留兼容 `set-borders` 别名，右键使用固定位置命令。

| 面板行/列 | 标签 | 本项目位置 | 固定版 BorderType |
| --- | --- | --- | --- |
| 1/1 | 上框线 | top | TOP |
| 1/2 | 下框线 | bottom | BOTTOM |
| 1/3 | 左框线 | left | LEFT |
| 1/4 | 右框线 | right | RIGHT |
| 1/5 | 无框线 | none | NONE |
| 2/1 | 所有边框 | all | ALL |
| 2/2 | 外侧框线 | outside | OUTSIDE |
| 2/3 | 内部框线 | inside | INSIDE |
| 2/4 | 内侧横线 | horizontal | HORIZONTAL |
| 2/5 | 内侧竖线 | vertical | VERTICAL |
| 3/1 | 右下斜线 | diagonal-down | TLBR / tlbr |
| 3/2 | 右下双斜线 | diagonal-down-pair | TLBC_TLMR / tlbc_tlmr |
| 3/3 | 右下三斜线 | diagonal-down-triple | TLBR_TLBC_TLMR / tlbr_tlbc_tlmr |
| 3/4 | 右上斜线 | diagonal-up | BLTR / bl_tr |
| 3/5 | 右上双斜线 | diagonal-up-pair | MLTR_BCTR / mltr_bctr |

菜单线型顺序：THIN(1)、HAIR(2)、DOTTED(3)、DASHED(4)、DASH_DOT(5)、DASH_DOT_DOT(6)、MEDIUM(8)、MEDIUM_DASHED(9)、MEDIUM_DASH_DOT(10)、MEDIUM_DASH_DOT_DOT(11)、THICK(13)、DOUBLE(7)。API 另接受 NONE(0)、SLANT_DASH_DOT(12)，面板通过 `antdv-next` Select.labelRender 显示这两种当前笔而不增加菜单选项。两种显示标签有独立的 19 语言字典；现有颜色字典仍为 13 个阶段键。

颜色面板为 5×9 的 45 色、1 个“更多颜色”阶段入口，以及自定义谱/色相/通道输入与确认/取消；草稿阶段不计作额外 Ribbon 根或新的 Office 功能。已有菜单回退树中的 15 位置、12 线型、45 色是相同语义的替代呈现，不能与面板重复计算完成数。

本项目新增上下文扩展根 1 个、子位置 15 个，适用于 cell / row / column / all 命中；复用相同位置命令和选区能力判断。固定 schema 中的边框入口位于 Ribbon 格式组，本批没有把新右键分组声称为固定版现有上下文入口，也未把该扩展视为新的边框能力完成证明。绘图/局部富文本上下文仍未核实。

## 模型、命令与状态

`sheet-cell-border.ts` 先构建不可变 patch plan，按多个选区的并集判断内外边，去重后不填充 bounding-box 间隙；部分选区触及合并格时扩展到该逻辑合并区域。合并格不新增内部 cardinal 线，斜线只写原主格；均匀 perimeter 在主格保留摘要，以兼容现有快照和 Office 写出链。实际物理格有自身边属性时优先使用它，显式 null 阻断主格回退。清除主格所代表的一个段之前，先物化其他继承段，避免邻格清除使整条合并边一起消失。

外边设置保留邻格其他边与扩展 metadata，只清其相向边。权限检查覆盖输入选区、扩展合并区域及所有真正变化的 patch，包括邻格和物化继承段；任一拒绝时不应用任何 patch。邻格没有需要改变的相向边时不扩大权限目标。2026-11-25 已接[有效默认样式](SHEET-04-default-style-inheritance.md)：无继承边框时 `none` 保持既有稀疏扫描；有继承边框时按需投影选区和邻格，明确 author border group 以阻断原边。公式、值、缓存及无关样式来源保留。

共享命令 ID 为 `set-border`、`set-border-basic`、`set-border-position`、`set-border-color`、`set-border-style` 和兼容 `set-borders`，均带 `sheet.command.` 前缀。`sheet-border-format-command.ts` 提供 action staging、两种行列坐标参数归一化及 model/view 适配；`sheet-toolbar-format.ts` 保留四参 `applySheetBorders` 的 void 契约，内部调用相同 plan。兼容 helper 没有合并上下文，合并调用应使用工作表 model/view 命令。

Studio 按 workbook + worksheet 保存 `{ type, style, color, active }` 笔状态。`getSheetBorderFormatState` 返回副本并检查查看权限。inactive 的直接颜色/线型 API 返回 false，仅配置笔；专门 `SheetBorderFormatStateChanged` 事件同步 UI。通用执行器对成功返回 true/false 的 handler 都发 `CommandExecuted`，因此 inactive 配置同样有真实命令事件，但没有内容历史或 revision 变化。取消、非法或受保护的请求先失败，不保存新笔，也没有成功命令/状态事件。

UI 颜色/线型选择显式激活当前 preset；这保留原面板“选色立即应用”的体验，区别于 inactive 直接 API。控件由接受后的 props 驱动，失败时不提前记忆新色或线型。设置相同内容仍报告已接受，但无模型变化不增加历史。笔状态作为会话工具状态，不随内容撤销或存入文档快照；关闭工作簿或 Studio 时清理，快照重开恢复默认笔。四类入口经同一命令和工作表保护判断，独立工作簿只读/设置样式权限仍随整体权限能力保留“未验收”。

Facade 新增 `FOffice.Enum.BorderType`、`FOffice.Enum.BorderStyleTypes` 和 `FRange.setBorder(type, style, color?)`，一次传完整范围、共用注册命令和按工作簿隔离的历史；`src/index.ts` 追加边框枚举、类型和 state/input 类型导出。内容提交关闭公式重算。默认 Studio 笔为黑色；旧三参 helper 的默认 `#1f2937` 为兼容保留，不把两者混写为相同默认。

## Canvas 与 UI

`sheetMergedBorderSegments(range, readBorder, visibleRange?)` 默认保留全量段契约；第三参只读取原合并四侧与可见行列区间的交集，回退始终读原 start，不生成虚假的 viewport 外框。渲染接入由主任务在 `scene.ts` 处理，并将边框放到 pane 背景之后；该接口和调用仍须统一渲染验证。读取计数用例不等同于耗时、内存或帧时间的实测基准。

`SheetBorderSelector.vue` 使用 `antdv-next` Button、Tooltip、Form/Form.Item、Select 和共享颜色控件；SVG 仅作为按钮内预览图。右键的 Menu/Button 与容器颜色、边框、阴影、悬浮、禁用和焦点状态使用 ConfigProvider token；Teleport 宿主通过共享 OfficeDirectionKey 设置 dir，弹出子菜单继续继承 provider。新增入口没有引入原生表单控件；字体浏览器输入定位也兼容库内 Input/TextArea。明暗主题、RTL、overflow、焦点和真实 IME 尚未执行浏览器验收。

## 已写用例与剩余缺口

新增 [sheet-border-format.contract.test.mjs](../../packages/office-studio/tests/sheet-border-format.contract.test.mjs) 38 组：固定枚举/预设/线型、非法参数、staging 和 inactive/UI activation、cardinal/diagonal golden mask、并集与重叠、不可变 plan、实际邻格及合并物化权限、局部合并边段/null、NONE 笔、Office 边界、筛除/隐藏行、model/view/legacy 保真、no-op、百万行合并的可见读取计数、19 语言、Ribbon/右键/Facade、历史、隔离和快照重开。

原 [sheet-border-format.browser.test.mjs](../../packages/office-studio/tests/sheet-border-format.browser.test.mjs) 有 14 个声明、15 个计划用例：真实 Canvas 保留、15 预设/12 线型、完整颜色草稿、主区复用、右键范围、合并段的撤销重做/快照、Facade 同步、inactive API-only 笔标签、取消、邻格保护、入口禁用、隐藏行、活动输入草稿、工作簿笔隔离和主题/RTL/grid overflow。2026-11-27 续核后为 15 个声明、19 个计划用例，见下节。全部仅已编写，未执行；浏览器用例尚不能作为渲染或功能通过证据。

集中验证仍需执行新旧相关契约、浏览器、类型、前端及 Rust/WASM 检查；确认真实 Canvas 的合并分段、背景覆盖顺序及各种线型可见结果。Office 文件仍走本项目解析与写出，SLANT_DASH_DOT 的 XLSX 线型映射由主任务补齐；XLSX 导出/重导入、独立 Excel 重开、主题色/tint、原始扩展边数据保真、局部富文本/绘图格式均未验收。

非清除大范围设置当前仍需 Worker/批量模型路径和取消契约；主格回退局部清除可能物化整侧继承段。继承空白边框清除新增临时同步投影预算 100,000 个位置，超预算的分块 Worker 路径尚未实现；该预算不是性能测量或大范围能力证明。没有首屏、滚动、编辑、缩放、Worker/WASM 传输、内存及同数据集前后基准，不声明大样本性能达标。匿名 standalone 仍沿用现有按名称标识会话的边界，重名匿名工作簿隔离尚未核实。

<a id="border-input-session-review"></a>

## 2026-11-27 边框输入会话与绘图显隐续核

本批属于 SHEET-04 的 Ribbon 状态与 Sheets 输入控制器；只传递现有会话能力并拒绝不适用的 UI 命令。边框 plan、领域/API、权限、内容历史、模型/保存和 Worker/WASM 协议没有新增实现；生产两文件增加 8 行、替换前删除 4 行，不增加 Ribbon 根、菜单子项或 UI handler。

| 固定版来源 | 核实结果与本地差异 |
| --- | --- |
| `packages/sheets-ui/src/menu/border.menu.ts:36` 的 `CellBorderSelectorMenuItemFactory` | `getCurrentRangeDisable$` 未传 `supportCellEdit=true`；根保留 `SetBorderBasicCommand`、笔订阅及 BUTTON_SELECTOR。原本地根仅用 `fontDisabled`，普通单元格或公式栏编辑尚能执行边框。 |
| `packages/sheets-ui/src/menu/menu-util.ts:163`、`:191`、`:192`、`:225` | 默认 `supportCellEdit=false`，当前单元格编辑器可见或公式栏焦点时禁用；普通绘图焦点时也禁用。本地复用既有 `canQuickFormatSelection`，包含工作台活动、选区权限、输入、组字及绘图选区；字体的局部输入格式能力保持独立。 |
| `packages/sheets-ui/src/menu/border.menu.ts:84` 的 `hidden$` | Sheets 外或 `FOCUSING_COMMON_DRAWINGS` 时隐藏。本地原边框根没有绘图 `hidden`，本批补 `hidden: state.drawingFocused`；已有字体、布局和合并显隐不重复修改。 |

`ExcelWorkbench.vue` 将现有能力传为可选 `SheetRibbonState.borderEditable`；`sheet-ribbon.ts` 用 `fontDisabled || state.borderEditable === false` 同步分体主按钮、下拉触发、15 位置、线型/颜色父菜单及其 12/45 子项。未传字段保持原调用方行为，classic 第 11 项、grid 第 2 行第 5 列、接受后的 preset/笔参数及面板/回退菜单计数不变。原 `RibbonItem.vue` 向 `SheetBorderSelector` 传递 disabled 并在禁用或隐藏时关闭面板；位置 Button、Form/Select 与颜色控件已承载该状态，无新原生控件或第二套颜色输入。

原 `executeRibbon` 会先 `finishEdit`，导致迟到的边框 UI 回调提交输入草稿。现在在该步骤前检查既有 `SHEET_BORDER_FORMAT_COMMAND_IDS`：`sheet.command.set-border`、`set-border-basic`、`set-border-position`、`set-border-color`、`set-border-style` 与 `set-borders` 共六个已定义 ID。此为既有分发的会话守卫，API 注册 ID 不计作新增 UI handler；没有新增别名。原面板 `changeBorderOptions` 继续经 basic 命令进入同一守卫。外部 API 在输入期间修改边框后继续编辑的现有行为与草稿保留，右键和独立 view 适配不改动。

[原 Ribbon 合同草稿](../../packages/office-studio/tests/sheet-ui-ribbon.contract.test.mjs) 新增 1 个声明/1 个计划用例，核对未传/可用/不可用输入能力、字体权限、无工作表、完整回退树、真实 invocation、位置参数和绘图显隐；当前 28 个声明/28 个计划用例。[原边框浏览器草稿](../../packages/office-studio/tests/sheet-border-format.browser.test.mjs) 新增 1 个声明，按 cell/formula × classic/grid 展开 4 个计划用例；夹具只新增真实 revision 观察，沿原 Studio、Canvas 和 ConfigProvider。普通输入、合成 composition 与明转暗期间验证两按钮禁用、六个合法参数的迟到 execute 不提交草稿，不改模型/revision/历史/笔/成功命令或笔事件；结束组字后仍编辑则继续禁用，Escape 后实际右框线、邻格反向边、原值/公式、一次历史、撤销/重做及内部 JSON 重开形成同一草稿闭环。

本批只读取源码和 diff、编写草稿及核对声明数量；运行测试、类型、构建、浏览器、Rust/WASM、Office 样本和基准均为 0，统一验证仍后置。全部能力保持未验收。绘图显隐只有 schema 合同草稿，尚未用真实绘图焦点浏览器验证；合成 composition 不替代 Windows 原生 IME，主题/焦点、独立 view 跨模块输入与重开、Office 边数据往返和真实性能缺口继续保留。
