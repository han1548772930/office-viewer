# SHEET-01 工作表生命周期与标签交互（2026-11-25，未验收）

固定来源：`E:/rust_test/univer`，Univer v1.6.0，提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。本记录延续[工作表复制证据](SHEET-01-worksheet-copy.md)，不新增或合并顶部入口计数。

## 固定源码契约

| 行为 | 命令 ID | 固定源码 |
| --- | --- | --- |
| 插入 | `sheet.command.insert-sheet` | `packages/sheets/src/commands/commands/insert-sheet.command.ts`；`packages/core/src/sheets/workbook.ts::uniqueSheetName` |
| 删除 | `sheet.command.remove-sheet` | `packages/sheets/src/commands/commands/remove-sheet.command.ts` |
| 重命名 | `sheet.command.set-worksheet-name` | `packages/sheets/src/commands/commands/set-worksheet-name.command.ts` |
| 隐藏 | `sheet.command.set-worksheet-hidden` | `packages/sheets/src/commands/commands/set-worksheet-hide.command.ts` |
| 显示 | `sheet.command.set-worksheet-show` | `packages/sheets/src/commands/commands/set-worksheet-show.command.ts` |
| 重排 | `sheet.command.set-worksheet-order` | `packages/sheets/src/commands/commands/set-worksheet-order.command.ts` |
| 标签颜色 | `sheet.command.set-tab-color` | `packages/sheets/src/commands/commands/set-tab-color.command.ts` |

`packages/core/src/sheets/typedef.ts::WorksheetHiddenState` 的三个状态为 VISIBLE=0、HIDDEN=1、VERY_HIDDEN=2。veryHidden 不列入普通取消隐藏入口，但 API 可以显示；它本身不代表查看权限。

引用变更依据 `packages/sheets-formula/src/controllers/__tests__/update-formula.controller.spec.ts`：重命名更新跨表公式，删除引用得到 `=#REF!`。链接生命周期参考 `packages/sheets-hyper-link/src/controllers/remove-sheet.controller.ts`。标签重排参考 `packages/sheets-ui/src/views/sheet-bar/sheet-bar-tabs/utils/sheet-tab-drag-sort.ts`，可见标签的位置必须映射回完整持久顺序。

## 已落盘的前端行为

- `sheet-worksheet-lifecycle.ts` 在独立候选上规划七类生命周期操作，校验名称、唯一 ID、位置、权限与最后可见表，保存 veryHidden 状态并在隐藏／删除活动表时选择可访问的可见邻表。显式插入名称冲突时生成唯一名称，不能仅凭大小写创建重名。
- `sheet-lifecycle-commands.ts` 统一 UI 和 Facade 的注册入口，使用工作簿内容事务、revision/history generation 与宿主校验。重命名前后事件在共享命令内发出；取消或重入产生的过期候选不发布。旧 `delete-sheet` 调用转到固定 `remove-sheet` 命令。
- `sheet-worksheet-references.ts` 更新单元格公式、限定名称、3D 端点、完整行列引用、迷你图、验证公式、条件格式和定义名称。外部引用、字符串和结构化列文字保持原样。局部名称的 `localSheetId` 通过稳定 sheet ID 在插入、删除、重排时重绑定。
- 删除时移除源表宿主的 tables/pivots/charts/slicers，清理绑定到删除 table/pivot 的切片器和透视图。其他表上的透视／图表保留显式失效源 ID，供刷新和源错误处理使用，不能静默改用宿主数据。
- `SheetFooter.vue` 保留隐藏表参与完整重排顺序，处理 before/after 拖放、顺序变化拒绝与边缘滚动。标签颜色实际绘制；底栏状态使用统一主题 token。
- `SheetTabColorDialog.vue` 使用 `antdv-next` Modal/Form/FormItem/ColorPicker，支持自定义及清除颜色。打开时捕获宿主与版本；切换工作簿、活动表或颜色来源变化会取消旧操作；关闭后请求恢复标签焦点。
- FWorkbook 的插入、删除和复制，以及 FWorksheet 的 setName、隐藏／显示／状态、颜色、setIndex 使用同一命令。`getHiddenState` 复用领域 helper，不维护重复隐藏语义。

## 实现差异和未核实项

本地使用领域快照事务实现逆操作，没有复刻上游 interceptor/mutation 内部类。插入重名采用可复现数字后缀并守住 Excel 31 UTF-16 限制，未复制上游可变计数器。已有字符串隐藏状态兼容读取，新隐藏操作保存数值枚举。

3D 引用删除端点时向区间内部收缩，内点删除保持端点；该行为已写入草稿，仍需 Office 文件样本核实。所有未建模的 Pro 资源、媒体／绘图关系及未知插件资源的删除／复制保真，不能由本记录推断为覆盖完成。

跨表透视的源查找已有显式缺表错误；错误发生后的旧缓存保留策略、图表 `source-not-found` 场景及 UI 恢复仍待联合验证。工作簿结构权限的服务端约束、持久历史和协作广播见[后端需求](../BACKEND_REQUIREMENTS.md)，当前暂缓实现。

## 后置验证

已编写但未运行：`sheet-worksheet-lifecycle.contract.test.mjs`（引用词法、资源清理、名称作用域、隐藏、权限、重排、DTO）和 `sheet-worksheet-lifecycle-api.contract.test.mjs`（事件取消／重入、单次内容历史、撤销重做、JSON 保存重开、过期宿主）。菜单草稿加入自定义颜色动作。

浏览器主题切换、键盘、拖放、弹层焦点恢复、IME、standalone 历史、Office 保存重开与跨表资源错误场景仍未验收。大样本候选复制、重命名扫描和重排的 Worker/WASM 路径及内存／耗时基准仍未验收。

按用户要求，本批未运行测试、类型检查、构建、Rust/WASM、浏览器、dev 或性能基准。SHEET-01 整组保持未验收。
