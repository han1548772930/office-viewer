# SHEET-04 公开样式读回 DTO

记录日期：2026-11-25。固定参考 `E:\rust_test\univer`，v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。状态：**前端已补代码，未验收**。续接[默认样式/引用/reset](SHEET-04-default-style-inheritance.md)，不新增 Ribbon、插件、溢出或上下文根；没有运行测试、类型、构建、浏览器/dev、Rust/WASM、服务或性能检查。

## 固定来源与实际差异

| 来源（相对固定根） | 契约 |
| --- | --- |
| `packages/core/src/types/interfaces/i-style-data.ts` | 字体 `ff/fs/bl/it/ul/st`、颜色 `cl/bg`、`n`、`tr/stf/ht/vt/tb` 和 `bd` 短字段；装饰/颜色/边框为对象，false 为 0，unset 与明确 null 分开。 |
| `packages/core/src/types/enum/text-style.ts`、`border-style-types.ts` | 水平 0–6、垂直 0–3、换行 0–3、边框 0–13；middle=2、overflow=1，不能把数组第一个同义值当实际默认。 |
| `packages/sheets/src/facade/f-range.ts`：`getCellStyleData/getFontFamily/getFontSize` | row/col 组合优先级、cell 只读直接；返回 `IStyleData`，字体读取使用 `ff/fs`。 |
| `packages/core/src/sheets/styles.ts` | 目录 ID 从实际样式目录解析；源数据扩展和直接作者来源保持。 |

此前本地 `getCellStyleData()` 只有 nested `font` 等字段，短字段消费者无法按固定 Facade 读取。本批提供短字段并保留 nested 作为本地兼容扩展，尚未输出全套固定工作簿 snapshot 或宣称全部未建模属性生效。

## 前端实现

- [sheet-public-style.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-public-style.ts) 提供 `sheetStyleToPublicData()`（短字段）与 `sheetStyleDataWithPublicFields()`（短字段+本地 nested），类型为 `SheetPublicStyleData/SheetApiStyleData`，已加入 index 导出。
- 只输出存在的作者属性；font false→0，颜色 null→`{rgb:null}`，装饰保留 t/c/cl/未知扩展而更新 s，数字格式和旋转保留无关来源字段。清掉装饰/数字格式/旋转/对齐时，不从来源 metadata 复活这些属性。
- `bd` 四边/六斜线逆映射回固定键；14 个线型共用 `sheetBorderStyleNumber()`，本地线型/颜色变化覆盖固定 s/cl，null、未知边和扩展保留。空 border group 输出 `{}`，不复活旧边。
- `normalizeSheetStyleInput()` 保存 ht/vt/tb 来源枚举；读回保留仍与当前属性兼容的 0/5/6 等同义值，后续真实属性改变输出新枚举。默认 middle/overflow 明确输出 2/1。
- [facade.ts](../../packages/office-studio/src/api/facade.ts) 的 `getCellStyleData('row'|'col'|'cell')` 接入该投影，`getFontFamily/getFontSize` 改读 ff/fs。cell 模式依然从已解析引用减 reset mask 的直接 patch 投影，无直接字号就没有 fs。
- 返回 DTO 和未知扩展均脱离模型；读取和调用方修改结果不添加内容历史或 revision。读回 DTO 可再作为 `cell.s` 输入，沿既有命令、权限、历史和快照链路。
- `resolveSheetStyleSource()` 核对目录 own key；普通对象原型的 constructor/toString 不会误认成目录项，确实存在的 own 同名 ID 仍可用。

这是领域样式→API DTO 的适配层，没有新增解析/公式引擎或第二套权威模型，不改变 Worker/WASM 内存协议。UI 没有新增控件，Canvas/Ribbon 输入继续使用现有 `antdv-next` 和统一 ConfigProvider。

## 草稿与剩余范围

[sheet-public-style.contract.test.mjs](../../packages/office-studio/tests/sheet-public-style.contract.test.mjs) 静态计 **13 个草稿**：10 个纯契约覆盖 sparse/null/false、默认枚举、来源扩展、属性更新、全部线型/边、清除不复活、返回隔离与目录 own key；3 个 Facade 场景覆盖 row/col/direct、reset→读回再写入→历史/重开，以及读取无内容副作用。全部未运行；没有新增浏览器用例，公开 API 与实际 Canvas 输入组合需统一回归。

基础[公开轴尺寸/自动模式/隐藏](SHEET-03-public-axis-import.md)已续作，另有 17 个未运行草稿。尚未完成或未验收：完整 eastAsia/复杂脚本、上标/下标、全部装饰、theme/tint/td/pd 的行为和每属性来源保真；完整自动行高组合、轴当前值读回、目录编辑 API、全套固定 snapshot DTO；真实 XLSX/独立 Office 往返；所有浏览器/主题/焦点/IME、Worker/WASM、内存与性能。未知字段可读/可存不能算相应产品能力实现。

对服务/文件而言，短字段是边界投影，模型中的 sparse/ref/reset/默认来源仍为权威；后续需求见[后端需求](../BACKEND_REQUIREMENTS.md)，完整 Sheets / Docs / Pro 范围保持。

2026-11-26 [范围主题续作](SHEET-04-range-themes.md)已将同一短字段投影用于九区段 builder/目录/API，加入默认/行列后的范围主题层，并保留缺目录规则的区域快照 blocker。该切片新增本地 Ribbon 根 1 和 46 个未运行草稿；本读回切片的根增量仍为 0、草稿数仍为 13。未知字段可读/可存不等于 Canvas 已实现相应行为，完整 fixed snapshot、Office 和主题/IME 仍未验收。
