# SHEET-06 地区货币、会计与格式预览

2026-11-26，参考 Univer v1.6.0 固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。按 [顶部入口清单](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)与 [Sheets 源码清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)复核「开始 → 数字」。

| Classic 顺序 | Grid 位置 | Univer 命令 | 本项目入口 |
| --- | --- | --- | --- |
| 数字格式 | r1c1 跨 4 列，160px | `sheet.operation.open.numfmt.panel`，选择预设执行 `sheet.command.numfmt.set.numfmt` | `sheet.command.set-number-format`，更多格式打开 `sheet.operation.open-custom-number-format` |
| 百分比 | r2c1 | `sheet.command.numfmt.set.percent` | `sheet.command.set-percent-format` |
| 货币 | r2c2 | `sheet.command.numfmt.set.currency` | `sheet.command.set-currency-format`，接受上游命令别名 |
| 增加小数位 | r2c3 | `sheet.command.numfmt.add.decimal.command` | `sheet.command.increase-decimals` |
| 减少小数位 | r2c4 | `sheet.command.numfmt.subtract.decimal.command` | `sheet.command.decrease-decimals` |

固定来源为 `packages/sheets-numfmt-ui/src/menu/schema.ts`、`menu.ts` 与 `number-format-options.ts`。数字组 5 个常驻根，格式菜单 13 个预设、4 个分隔、1 个面板入口。Classic 和 Grid 复用入口，More 收纳同组。扩展、独立溢出、上下文和常驻根的本批增量均为 0。固定版仅在 Sheets 显示，禁用条件读取 workbook 可编辑、worksheet 可编辑及设置样式权限和范围编辑权限。本项目保留 Sheets 工作台、工作表/范围保护、编辑会话状态和混合格式的现有判断。

## 区域、草稿与持久化

- 固定版 `packages/core/src/services/locale/currency-symbol.ts` 列 19 个区域及固定符号。`packages/sheets-numfmt/src/commands/commands/set-currency.command.ts` 在执行时读 RegionService，按所有当前选区提交两位小数、千分位和红色括号负数格式。`packages/sheets-numfmt-ui/src/menu/number-format-options.ts` 同样用区域符号生成货币、会计和整数货币三个菜单格式。本项目沿用这些模板。精确格式码保存在 `CellStyle.numberFormat`，切换区域后更新快捷按钮的图标和预览的分隔符，现有单元格保留保存的符号与数值。工作台使用实际区域符号作为图标，例如 NT$ 或 Rp。固定版使用各地区分组的货币图标。
- `packages/sheets-numfmt-ui/src/views/components/{Currency,Accounting}.tsx` 从当前模式和使用习惯初始化，再用 `packages/sheets-numfmt/src/base/const/formatdetail.ts` 的六种负数模式或会计填充模式。固定版高级面板缺省取最近使用或第一个通用符号。本项目设计为当前格式的符号优先，其次最近使用，二者缺省时用当前区域。面板仍保留原有 27 个符号，并加入 6 个地区符号（NT$、HK$、R$、Rp、zł、﷼），共 33 个去重选项。选择来源里已有的 OOXML 方括号符号也会临时列入。选项及已确认偏好共用同一目录。
- `sheet-currency.ts` 从正值段的符号文本或 OOXML `[$符号-地区]` 识别货币，忽略填充、留位和其他分段里的同形字符。多字符符号优先识别完整值。本地 `sheet-number-format.ts` 的会计分类与高级面板共用，覆盖固定版以 `_(` 开头的会计格式，以及本地已有的菜单会计格式和带 `*` 填充的会计格式。只改小数位时保留原始分段、负值颜色、符号位置和地区标记。显式选择货币或负数样式仍切换到该选择对应的模板。
- `CustomNumberFormatPanel.vue` 负数选项调用与网格同一格式渲染器，按当前区域、草稿小数位显示文本和模式指定颜色。面板沿用 [本地组件文档](../llm.md)中的 antdv-next 1.8.2 Form/FormItem、Select、InputNumber、Menu 和反馈控件，以及现有 ConfigProvider 主题与焦点恢复。预览仅更新 Canvas 的格式覆写，取消丢弃覆盖，确认后才更新内容及偏好。
- 工作台既有 `runNumberFormatCommand` 和 `FRange.setNumberFormat` / `setNumberFormats` 共用 Studio 数字格式命令及计划。单次格式提交只写选区的显示格式并进一次工作簿历史，数值、公式文本及已知缓存类型继续由原模型保留。重复应用现有格式不新增历史。内部快照与 `crates/office-writer/src/xlsx.rs` 的 numFmt 注册读取同一格式码。面板草稿创建参数现分别接收界面语言、区域和最近使用符号，公开 Facade 签名与 WASM 协议不变。

## 验证状态

已静态核对固定源码、地区目录及现有命令、快照和写出路径。给 `sheet-number-format.contract.test.mjs` 和 `sheet-number-format.browser.test.mjs` 各补两条草稿，涵盖 19 个地区模板、符号识别、会计分段保持、快捷按钮激活、负数地区预览、Facade、撤销重做及内部快照重开。

测试、类型检查、构建、浏览器及 Rust/WASM 留待统一阶段。真实 XLSX 导出与独立 Office 重开、全部地区实际预览、主题和焦点的真实复核，以及大范围格式应用性能仍为未验收。
