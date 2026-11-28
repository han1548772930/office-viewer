# SHEET-18：表格选择器与创建入口续核

记录日期：2026-11-27；创建接线续作日期：2026-11-28。状态：固定表格创建根、真实范围弹窗、Studio 命令、权限、一次内容历史及宿主生命周期已落盘，合同与浏览器草稿未运行；SHEET-18 仍未验收。

## 固定来源与入口清点

固定上游为 `E:/rust_test/univer` v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。以下为公开 OSS 文件；本轮未安装或执行 Univer SDK。

| 来源 | 已核实契约 |
| --- | --- |
| [table UI schema](../../../univer/packages/sheets-table-ui/src/menu/schema.ts) | Data / ORGANIZATION 的 sheet.operation.open-table-selector，order 2、grid row 1 / column 2；MAIN_AREA / LAYOUT 下表专属插入、删除两个上下文根，各有行/列两个子项 |
| [table UI menu](../../../univer/packages/sheets-table-ui/src/menu/menu.ts) | 根是 Button，Sheets 工作簿决定显隐；getCurrentRangeDisable$(accessor, {}, true) 决定禁用。表内选区决定上下文显隐，表头行另有限制 |
| [menu-util.ts](../../../univer/packages/sheets-ui/src/menu/menu-util.ts) | 第三个参数为 supportCellEdit，允许编辑桥/公式栏处于编辑态，不能解释成单选条件。实际禁用包括无工作簿/活动表、图形焦点和 WorkbookEditablePermission；默认参数没有另指定 worksheet/range 权限类型 |
| [open-table-selector.operation.ts](../../../univer/packages/sheets-table-ui/src/commands/operations/open-table-selector.operation.ts) | 使用最后选区，单格向上/左/右/下扩展连续区域，打开 300px 范围弹窗；取消返回无范围，确认调用 sheet.command.add-table |
| [SheetTableSelector.tsx](../../../univer/packages/sheets-table-ui/src/views/components/SheetTableSelector.tsx) | 单范围、不跨 sheet；拒绝合并区、其他表重叠和仅一行。既有表范围更新另要求原表头位置及范围相交 |
| [add-sheet-table.command.ts](../../../univer/packages/sheets-table/src/commands/commands/add-sheet-table.command.ts) | 缺省或无效名称按本地化前缀自动生成，冲突集合包括表与定义名称；创建列，执行 Add mutation，并按 workbook 保存 Delete 逆 mutation 的内容历史 |
| [add-sheet-table.mutation.ts](../../../univer/packages/sheets-table/src/commands/mutations/add-sheet-table.mutation.ts)、[table.service.ts](../../../univer/packages/sheets-table/src/services/table.service.ts) | mutation 将 range/header/options/meta 交给 SheetTableService → TableManager，并返回 true；这两层未见创建权限 guard，不能把 UI 禁用等同命令层拒绝已证实 |
| [table-manager.ts](../../../univer/packages/sheets-table/src/models/table-manager.ts)、[util.ts](../../../univer/packages/sheets-table/src/util.ts) | 表模型、列头和 JSON；convertCellDataToString 支持富文本、字符串、数字和布尔 TRUE/FALSE，getExistingNamesSet 纳入定义名称 |
| [sheets-table.controller.ts](../../../univer/packages/sheets-table/src/controllers/sheets-table.controller.ts)、[const.ts](../../../univer/packages/sheets-table/src/const.ts) | 注册 SHEET_TABLE_PLUGIN 的资源读写与单位生命周期 |

分类计数：该固定 schema 有 Sheets 扩展 Ribbon 根 1、MAIN_AREA 上下文根 2、上下文动作 4；创建根没有 selector 子项或独立快捷键。More 复用同一根，没有另加业务根。本轮未将全表格 API、动态表头控件或其他上下文命令与这 7 个静态入口相加；完整 SHEET-18 清点和验收范围保持不变。

## 本地现状与创建边界

| 本地文件/链路 | 已有行为 | 实际缺口 |
| --- | --- | --- |
| [sheet-tables.ts](../../packages/office-studio/src/features/sheet-tables/sheet-tables.ts) | create/resize/totals/delete 纯函数；create 需要显式名称，拒绝单行、表重叠和合并交叠，默认 TableStyleMedium2。真实 add-table 现经命令候选调用 create | 后续 resize/totals/delete 的完整工作台与公开资源入口仍须逐条核对，不能由创建链路推定完成 |
| 同一模型创建与调整 | 创建读取完整本地 richText 文本、数字和 TRUE/FALSE，保留原空表头与唯一列名规则；命令传入本地化表/列前缀，resize 保留列数 | 完整公开 p/body 优先契约与变列数的调整未实现；本地化资源不足时沿用 Table/Column 缺省 |
| [domain/model.ts](../../packages/office-studio/src/domain/model.ts)、[workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts) | 原生表格的 ID/name/range/columns/header/totals/style/autoFilter 基础字段；定义名称保存在 workbook.metadata.definedNames/localDefinedNames | 没有完整上游 options/meta 或 SHEET_TABLE_PLUGIN 公共资源映射；这些基础字段不能当作完整表模型保真 |
| [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) | 保留本地 tables，并从表筛选派生隐藏行；创建后工作台应用该投影，不重复提交内容历史 | 后续专属操作、表头控件与筛选排序仍需完整闭环 |
| [xlsx_table.rs](../../crates/office-writer/src/xlsx_table.rs)、[xlsx.rs](../../crates/office-writer/src/xlsx.rs) | 写原生 table XML、列名、header/totals/style 和 worksheet tableParts | 当前表 XML 总写空 autoFilter ref；不能声明表筛选条件完整往返或创建后独立 XLSX 重开已验证 |
| [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | 新增固定 Data/organization order 2、grid row 1 / column 2 创建根；增强“分组显示”移到 row 1 / column 4，保留原大纲设置位置 | Classic/grid/More、主题、真实禁用态和窄屏入口仍未运行验收 |

权限结论：领域 createSheetTable 保持纯模型函数，不接受用户或宿主。真实命令现复用 Studio 的 assertCanViewWorksheet/Range、assertCanEditWorksheet/Range，在打开与确认时核对文档、工作表、revision、history generation、user 与活动宿主，创建权限覆盖最终完整范围；候选前、候选后及 commit mutator 再查 owner 与范围权限。对上游其他全局拦截器的直接 API 权限效果仍未核实，不以本项目额外 guard 冒称上游已具备相同实现。

## 2026-11-28 实际创建接线

前轮创建方案已按现有架构实施。保持已有严格领域调用行为；新增外部参数在命令入口校验，内部使用同一领域候选。单格范围复用现有 expandSheetContinuousRange，多格保留末选区原范围，没有另建连续区域算法。创建根使本地常规 Sheets 根增量为 **+1**；Docs、表格上下文、独立溢出和选择器子项增量为 0，此增量不计作功能验收。

| 文件 | 实际职责 |
| --- | --- |
| [sheet-table-create.ts](../../packages/office-studio/src/features/sheet-tables/sheet-table-create.ts) | 真实两个命令 ID、类型化参数/请求、身份校验、连续区域、地址与公开范围 DTO；有效名称保留，无效/冲突名称按表及 definedNames/localDefinedNames 生成默认名 |
| [sheet-table-commands.ts](../../packages/office-studio/src/api/sheet-table-commands.ts) | 注册 open-table-selector / add-table，typed 事件请求不写历史；确认生成领域候选后一次 commitWorkbook，显式保存 revision/history 与权限 guard |
| [SheetTableSelectorDialog.vue](../../packages/office-studio/src/components/spreadsheet/SheetTableSelectorDialog.vue) | 300px Ant Modal、Form/FormItem/Input/Button/Alert；单范围草稿、确认错误、取消和组合输入；afterClose 使用库的函数 prop，关闭库自带回焦 |
| [use-sheet-table-workbench.ts](../../packages/office-studio/src/components/spreadsheet/use-sheet-table-workbench.ts) | 活动宿主、模型对象身份、请求 revision/history/user 与会话 generation；切换/销毁取消，提交与 UI 刷新后再查归属，动画结束后恢复合法焦点 |
| [office-studio.ts](../../packages/office-studio/src/api/office-studio.ts)、[events.ts](../../packages/office-studio/src/api/events.ts)、[index.ts](../../packages/office-studio/src/index.ts) | 主任务注册/释放命令，公开类型化 SheetTableSelectorRequested 事件及命令/请求类型 |
| [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | 固定创建根、tableAvailable/tableEditable 状态和 grid 位置；移动本地增强大纲项 |
| [ExcelWorkbench.vue](../../packages/office-studio/src/components/ExcelWorkbench.vue) | hook/弹窗与命令装配；普通单格编辑先 finishEdit，组字/图形/冲突弹层拒绝开窗；已有弹窗的延迟回焦避开表格输入；发布新投影而不重复内容历史 |
| [sheet-table-command.contract.test.mjs](../../packages/office-studio/tests/sheet-table-command.contract.test.mjs) | 10 个未运行声明：纯候选与真实 Studio 创建、默认名/表头、完整范围权限、合并/重叠、非法参数、owner、取消、Undo/Redo 与内部保存重开 |
| [sheet-table-host.contract.test.mjs](../../packages/office-studio/tests/sheet-table-host.contract.test.mjs) | 10 个未运行声明：真实 Studio/hook 的请求、阻塞、取消、错误校正、失效、同 ID 模型更换、销毁、重入发布/刷新与延迟焦点 |
| [sheet-table-selector.browser.test.mjs](../../packages/office-studio/tests/sheet-table-selector.browser.test.mjs)、[真实夹具](../../packages/office-studio/tests/fixtures/sheet-table-selector.html) | 7 个参数化声明，计划展开为 classic/grid 的 14 场景；真实 ExcelWorkbench、Ant Modal 与 Canvas2D，覆盖输入/IME、保护/冲突、取消、历史和内部重开 |

模块归属为 table feature 的同步模型、API 命令和 Vue 工作台视图；当前创建候选复用 TypeScript 模型，连续区域路径仍需同数据集基准决定大规模 Worker/Rust/WASM 分派。本批不变更 WASM 内存所有权或跨边界协议，不能把逐单元格读取记为已优化。

创建闭环的验收必须覆盖：打开后不修改内容，取消不写历史，拒绝合并/重叠/受保护输入，确认只产生一次内容事务，撤销重做恢复模型，保存后独立重开保留表身份、范围、列与样式。当前没有运行任何这些验证。

完整 SHEET-18 仍包括表专属插删、范围/列调整、计算列、自动扩展、筛选排序联动、结构化引用、转换为范围、公开 Facade/资源和 Office 全面往返。创建接线是后续批次的依赖，不降低这些目标，也不以新增 root 标记功能完成。

<a id="table-header-merge-domain-continuation"></a>

## 既有表头与合并交叠续修

前轮小补仅修改既有同步 [sheet-tables.ts](../../packages/office-studio/src/features/sheet-tables/sheet-tables.ts)，当时尚未实施创建接线；2026-11-28 的实际接线以上节为准。固定 [util.ts](../../../univer/packages/sheets-table/src/util.ts) 的 `convertCellDataToString` 和 [table-manager.ts](../../../univer/packages/sheets-table/src/models/table-manager.ts) 的 `getColumnHeader` 明确读取正文、标量值和布尔 TRUE/FALSE；[SheetTableSelector.tsx](../../../univer/packages/sheets-table-ui/src/views/components/SheetTableSelector.tsx) 明确拒绝与合并区相交的表范围。

本地 `readUniqueHeaders` 现在优先连接完整 `richText` 运行，数字使用保存值而非格式化 display，布尔及公式布尔缓存复用 `normalizeSheetCellValueType` 输出 TRUE/FALSE。强制字符串保留文本；空运行正文不借用旧缓存值。原有 trim、ColumnN 空表头和不区分大小写的 `_2` 后缀保持，创建不修改源单元格、运行、公式或样式。

`createSheetTable` 在构造表候选前拒绝包含、被包含、边缘接触及部分交叠的合并区域；四侧相邻合并允许。这里把公开 selector 的合并语义落实在本项目既有领域创建入口，不声称固定上游 Add mutation 已有相同 guard。没有新增内部重复校验、UI、命令、公共 DTO 或 Worker/WASM 协议。

新增 [sheet-tables.contract.test.mjs](../../packages/office-studio/tests/sheet-tables.contract.test.mjs) 的 6 个测试声明，直接测真实领域函数与既有快照归一化：标量/重复/空表头、声明和缓存布尔、富文本与来源保留、合并交叠拒绝及原子性、四侧邻接、内部保存重开和旧严格名称调用。全部未运行；内部重开草稿不能替代真实 UI 内容历史或 Office 重导入证据。

完整公开 `p.body.dataStream` 与非文本正文、表列调整和结构资源/Office 目标仍有缺口；创建/权限/历史接线已按上节实施，但真实验收仍后置。此补项没有性能实测，也没有将现有逐列表头读取记为性能优化。SHEET-18 保持未验收。

## 前轮只读核对与验证

前轮仅新增本证据，表格实现与公开接口均未改变。前轮已静态读取上述固定源码和本地调用，原 19 个本地来源链接存在；该历史记录不覆盖本次续修。所有测试、类型检查、构建、Rust/WASM、dev、浏览器及基准继续后置。没有文件样本验证或性能实测，SHEET-18 保持 **未验收**。

2026-11-28 创建批的只读收尾覆盖上述 10 个生产/草稿/夹具文件：10 段脚本、2 个 Vue 模板、2 个样式块及 185 处本地模块引用语法/路径通过；三份 `.mjs` 均通过 `node --check`。本证据的 37 个本地文件目标（其中 15 个固定上游目标）存在，命令/host/browser 声明分别为 10/10/7，浏览器计划按两种布局展开为 14 场景。限定 `git diff --check` 通过，仅提示现有 LF/CRLF 约定；没有运行测试、类型检查、构建、Rust/WASM、dev、浏览器、Office 往返或性能基准。
