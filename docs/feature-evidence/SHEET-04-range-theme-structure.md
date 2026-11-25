# SHEET-04 范围主题局部移位与剪切规则

记录日期：2026-11-26。固定参考 Univer v1.6.0，`E:\rust_test\univer`，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。续接[范围主题基础批次](SHEET-04-range-themes.md)。状态：**实现与草稿已落盘，未验收**。按用户安排没有执行测试、类型检查、构建、浏览器/dev、Rust/WASM 或基准。

## 固定契约与入口

| 固定根下来源 | 核对内容 |
| --- | --- |
| `packages/sheets/src/commands/commands/insert-range-move-right.command.ts`、`insert-range-move-down.command.ts` | 当前范围/最后选区、重叠插入判断、局部横纵带、样式保留、尾部扩展、资源 interceptor 与同一内容历史。 |
| `packages/sheets/src/commands/commands/delete-range-move-left.command.ts`、`delete-range-move-up.command.ts` | 删除带与后续内容移位、选择跟随、资源 interceptor 与撤销。 |
| `packages/sheets/src/commands/commands/move-range.command.ts`、`commands/utils/handle-range.mutation.ts` | 来源先读取、目标覆盖、移动与资源更新同一操作。 |
| `packages/sheets/src/controllers/zebar-crossing.controller.ts`、`models/range-theme-model.ts` | 范围主题规则、范围资源变换及样式读取。具体局部拆分后的表头/隐藏行相位未做上游运行验证。 |
| `packages/sheets/src/facade/f-range.ts`：`insertCells` / `deleteCells`；`packages/sheets-ui/src/menu/schema.ts` | 四个局部移位命令的公开和上下文入口。 |

本批常规 Ribbon 根增量 **0**、新增上下文根 **0**。已有右键“插入/删除”子项分别指向 `sheet.command.insert-range-move-right/down` 与 `sheet.command.delete-range-move-left/up`；剪切/粘贴复用已有共享命令。本批只接主题规则的结构行为，不用子项数量证明完整功能。

## 模型与实现

[domain/model.ts](../../packages/office-studio/src/domain/model.ts) 的 `SheetRangeThemeRule` 增加可选 `styleRange`。`range` 表示真实物理成员，`styleRange ?? range` 表示逻辑表头、条纹和末行/末列起点；拆分或部分移动后保留逻辑范围，避免每个片段重新生成表头。移动到 A1 的片段可有负逻辑起点，逻辑范围必须包含真实成员，允许行坐标 `[-1_048_576, 2_097_151]`、列坐标 `[-16_384, 32_767]`，物理范围仍遵守 Office 上限。此字段是本地协议扩展，不能冒称上游原样字段。

[sheet-range-theme-structure.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-structure.ts) 提供局部带、整轴和剪切 planner。局部移位先沿垂直方向分片，只变换选中带；整轴插删/移动/复制也变换逻辑几何。完整跨表移动可保留 ID，片段和冲突 ID 重映射，metadata 脱离原作者对象。插入会把已有物理成员移出 Office 上限时拒绝整次准备，不静默截断规则；未触及的带保持原规则。源数组不原地修改。

剪切先缓存原始来源，裁掉源与被覆盖的目标，再追加来源规则；同表重叠遵守同样次序。未注册的后层主题继续作为 blocker。无主题的来源会清掉目标相应主题成员。保留片段仍沿原逻辑位置解析，显式同名重新应用可保留 ID 并重新设置表头。

[sheet-range-theme-model.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-model.ts) 校验/归一化与轴入口使用该 planner；[sheet-range-theme-resolver.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-range-theme-resolver.ts) 的 live、区域快照和 public style 读回统一读取逻辑范围。内部保存/重开保留该字段。

[sheet-cell-range-move.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-cell-range-move.ts) 已在后续批改为共享 planner 的兼容适配，工作台/API 直接使用完整候选，详见[单元格移位事务](SHEET-02-cell-shift-transactions.md)。[sheet-clipboard.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard.ts) 仅在内部 cut + all 时移动作者规则；[sheet-clipboard-cell.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-cell.ts) 保留对应 cell 的 raw direct/reference/reset/theme 来源，虚拟主题空白不创建 cell。普通 copy 仍固化有效格式，values-only 不移动主题规则。

[sheet-clipboard-data.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-data.ts) 的 cut fingerprint 纳入相交原始规则，metadata/ID 的变化也使 pending cut 过期。[sheet-clipboard-changes.ts](../../packages/office-studio/src/features/sheet-workbench/sheet-clipboard-changes.ts) 将实际主题成员/逻辑几何差异纳入权限与事件范围；只因分裂重映射 ID 而未改样式的保留片段不扩大权限范围。

算法按规则矩形工作，不遍历主题空白格。当前同步保护为 100,000 个片段与 5,000,000 次差异工作量，属于拒绝预算，不是测量阈值或性能达标证据；大量规则索引/Worker/Rust/WASM 和取消仍待实施。

## 草稿与 UI

| 文件 | 未运行草稿 | 计划验证 |
| --- | ---: | --- |
| [sheet-range-theme-structure.contract.test.mjs](../../packages/office-studio/tests/sheet-range-theme-structure.contract.test.mjs) | 26（21 纯契约 + 5 Facade） | 四方向/真实路径、边界与 malformed 逻辑几何、部分/重叠/跨表 cut、负起点、blocker、后续轴编辑、ID/来源、稀疏空白、fingerprint、权限、历史和内部重开。 |
| [sheet-range-theme-structure.browser.test.mjs](../../packages/office-studio/tests/sheet-range-theme-structure.browser.test.mjs) + [fixture](../../packages/office-studio/tests/fixtures/sheet-range-theme.html) | 6 | 真实 Ant 右键下移、Canvas cut/paste、稀疏作者格、历史、跨表 A1/内部重开、重叠、目标保护、暗色/grid 下动态主题与 Ant 输入回焦。 |

本批 **32** 个草稿，基础批次的 **46** 保持独立记录，合计 **78** 个未运行场景。未新增可替换 UI 控件；右键和输入复用现有 `antdv-next` 控件与 ConfigProvider。中文 IME、明暗主题、焦点与两渲染后端尚未实际验证。

## 未完成与未核实

- 后续[共享 cell shift/确认批](SHEET-02-cell-shift-transactions.md)已接四核心+四确认命令、Facade、完整候选/权限/验证、内容与选择历史、资源/事件及 revision/宿主保护；本批不再由工作台独立处理。56 个新草稿未运行，完整 drawing/media、所有资源组合与 Office/性能仍未验收，不能据接线宣称整个移位能力已闭环。
- `styleRange` 为本地扩展；隐藏行条纹、复杂多次拆分、逻辑起点经过所有轴组合及动态权限需要统一运行验证。原始规则顺序按本地后应用优先，不宣称等同上游 RTree 动态结果。
- 全套固定公开资源输出、XLSX 样式/default/theme/tint 映射及真实 Office 独立重开仍有缺口。内部 JSON 重开不代替 Office 验收。
- 普通 copy、填充和格式刷继续使用既有来源投影；所有跨模块组合、大样本传输/内存和性能仍未验收。后端与文件协议登记见[后端需求](../BACKEND_REQUIREMENTS.md)。
