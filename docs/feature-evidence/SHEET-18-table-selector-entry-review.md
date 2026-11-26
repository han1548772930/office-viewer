# SHEET-18：表格选择器与创建入口续核

记录日期：2026-11-27。状态：源码与现有实现已静态核对，创建接线方案待确认，功能未验收。本轮未新增按钮、命令或弹窗。

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
| [sheet-tables.ts](../../packages/office-studio/src/features/sheet-tables/sheet-tables.ts) | create/resize/totals/delete 纯函数；create 需要显式名称，拒绝单行和表重叠，默认 TableStyleMedium2 | 生产 Studio/工作台未调用这些函数，当前调用只有测试夹具/切片器草稿。没有真实 add-table 命令、打开范围操作、菜单状态、owner、权限或内容历史接线；create 未拒绝合并区，名称冲突仅查表 |
| 同一模型创建与调整 | 字符串表头可生成唯一列名；resize 保留列数 | 数字、布尔、富文本表头落为 ColumnN，未对齐上游转换；变列数的调整未实现 |
| [domain/model.ts](../../packages/office-studio/src/domain/model.ts)、[workbookNormalizer.ts](../../packages/office-studio/src/parser/workbookNormalizer.ts) | 原生表格的 ID/name/range/columns/header/totals/style/autoFilter 基础字段；定义名称保存在 workbook.metadata.definedNames/localDefinedNames | 没有完整上游 options/meta 或 SHEET_TABLE_PLUGIN 公共资源映射；这些基础字段不能当作完整表模型保真 |
| [excelViewAdapter.ts](../../packages/office-studio/src/services/excelViewAdapter.ts) | 保留本地 tables，并从表筛选派生隐藏行 | 没有表格创建和后续专属操作的工作台命令入口 |
| [xlsx_table.rs](../../crates/office-writer/src/xlsx_table.rs)、[xlsx.rs](../../crates/office-writer/src/xlsx.rs) | 写原生 table XML、列名、header/totals/style 和 worksheet tableParts | 当前表 XML 总写空 autoFilter ref；不能声明表筛选条件完整往返或创建后独立 XLSX 重开已验证 |
| [sheet-ribbon.ts](../../packages/office-studio/src/components/spreadsheet/sheet-ribbon.ts) | Organization 有排序、筛选和三个大纲根 | 没有表格创建根；本地增强“分组显示”占用固定表格的 grid row 1 / column 2，接入表格时需重新安排增强项 |

权限结论：本地 createSheetTable 是未接线的领域函数，不接受用户或宿主，不提供创建权限/历史结果。未来 API 入口应复用 Studio 的 assertCanEditWorksheet 和既有保护范围能力，并在确认时核对文档、表、revision、history generation、user 与活动宿主；不得靠按钮 disabled 代替真实命令拒绝。对上游其他全局拦截器的直接 API 权限效果本轮未核实。

## 可审阅实施方案

下列为约 **785 行、10 个实现/草稿文件** 的创建闭环方案，尚未实施。保持已有领域函数调用方行为；新增外部参数在命令入口校验，内部使用同一领域候选。连续区域复用主线方案，待该方案确认后接入，不在这里维护第二份权威算法。

| 文件 | 预计行数 | 具体职责 |
| --- | --- | --- |
| 既有 features/sheet-tables/sheet-tables.ts | 60 | 补创建合并区边界和标量/富文本标题契约，保留现有严格名称调用方式；命令按现有 metadata 定义名称和表名集合生成合法默认名 |
| 新 api/sheet-table-commands.ts | 160 | 类型化 open-table-selector / add-table 入口、权限/owner 与确认后一次 commitWorkbook；重用领域创建结果及现有历史，不另建模型 |
| 新 features/sheet-tables/SheetTableSelectorDialog.vue | 140 | 300px Ant Modal、Form/FormItem/Input/Button，单范围编辑、错误与确认/取消；确认调用真实 add-table，取消不编辑 |
| 新 components/spreadsheet/use-sheet-table-workbench.ts | 110 | 操作请求订阅、活动宿主/草稿所有者、关闭/换表/销毁与焦点恢复；应用结果刷新 |
| 既有 api/office-studio.ts | 5 | 仅注册新命令，由主线修改 |
| 既有 api/events.ts | 10 | 类型化表格选择器请求与所有者字段 |
| 既有 components/spreadsheet/sheet-ribbon.ts | 15 | 固定创建 root、权限状态和 grid 位置；移动本地增强项，由主线修改 |
| 既有 components/ExcelWorkbench.vue | 35 | hook、弹窗装配和统一命令路由，由主线修改 |
| 新 tests/sheet-table-command.contract.test.mjs | 150 | 直接测真实 Studio：创建、默认名/表头、合并/重叠/权限拒绝、owner 失效、Undo/Redo 和原有调用兼容 |
| 新 tests/sheet-table-selector.browser.test.mjs | 100 | 真实工作台开窗、连续区域、确认/取消、焦点/换表、内容历史及保存/重开；本阶段仅草稿 |

模块归属为 table feature 的同步模型、API 命令和 Vue 工作台视图；当前创建候选复用 TypeScript 模型，连续区域计算按主线 Worker/Rust/WASM 协议和同数据集基准选路径。本方案不变更 WASM 内存所有权或跨边界协议。

创建闭环的验收必须覆盖：打开后不修改内容，取消不写历史，拒绝合并/重叠/受保护输入，确认只产生一次内容事务，撤销重做恢复模型，保存后独立重开保留表身份、范围、列与样式。当前没有运行任何这些验证。

完整 SHEET-18 仍包括表专属插删、范围/列调整、计算列、自动扩展、筛选排序联动、结构化引用、转换为范围、公开 Facade/资源和 Office 全面往返。创建方案是后续批次的依赖，不降低这些目标，也不以新增 root 标记功能完成。

## 实际修改与验证

本轮只新增本证据，表格实现与公开接口均未改变。已静态读取上述固定源码和本地调用，本文件 19 个本地来源链接存在；所有测试、类型检查、构建、Rust/WASM、dev、浏览器及基准后置。没有文件样本验证或性能实测，SHEET-18 保持 **未验收**。
