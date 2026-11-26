# SHEET-25：切片器与时间线样式资源

记录日期：2026-11-26。状态：**领域类型、样式目录、默认/状态投影、导入保留与草稿已落盘，未验收**。

按用户当前要求只做静态读取和编辑；未运行测试、类型检查、构建、Rust/WASM、dev、服务、浏览器、基准或 `git diff --check`。本文不把样式名、状态枚举或计划用例数量作为功能完成证明。

## 固定公开来源

OSS 基线为 `E:/rust_test/univer` v1.6.0、commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；下列公开发布包均为 1.6.0 静态研究资料，没有安装或执行商业 SDK。

| 精确来源 | 公开事实与本地处理 |
| --- | --- |
| `.research.local/univer-1.6.0/slicer/lib/types/style.d.ts` 第 2–65 行 | 14 个 Slicer 内置名称、12 个 Timeline 内置名称；8 种 selected/hasData/hovered 组合与 7 种 Timeline 元素；文档默认值与自定义定义使用稳定名称引用 |
| 同包 `lib/es/index.js` 的 `De/Oe/ke/Ae/je/Me/Ne/Pe/Fe/Ie/Le/Re/ze/A/Be/Ve` | 名称保留原大小写/空白，非空且最多 255 字符；自定义重复按小写核对，精确内置名称不能覆盖；重复或未知状态拒绝；默认引用必须存在；未知扩展不进入格式投影 |
| `E:/rust_test/univer/packages/core/src/types/interfaces/i-style-data.ts` 的 IColorStyle、ITextDecoration、IBorderData，及 `packages/engine-render/src/components/docs/extensions/line.ts` 第 61 行 | 保留 rgb/theme、上划线 ol 和 10 个边框方向；ol 在固定渲染源码中为 overline，不能只存一种颜色后声称完整模型 |
| `.research.local/univer-1.6.0/slicer/lib/types/rendering/theme.d.ts` 与 ES 的 SlicerThemeService / `_defaultSlicerItem` / `_defaultTimelineElement` / `_resolveFormatting` | 样式定义与颜色主题分离；容器、表头、每种成员/时间元素各自先用主题/内置 palette 默认，再覆盖作者 formatting。本批提供稀疏作者投影；实际颜色与绘制由原 Canvas/Ant 管线接入 |
| 同 ES `on/sn`（parseSlicerResource） | 发布资源中有 slicerStyles/timelineStyles，且视图可保存可选 style。固定解析器遇到无效家族或未知引用拒绝整项资源；本地保留原数据与诊断，使另一有效家族可独立使用，这是明确实现差异 |
| `.research.local/univer-1.6.0/packages/sheets-slicer/lib/types/facade/f-workbook.d.ts` 第 32–55 行 | Workbook 样式列表返回去重的内置与已保存自定义名称，并返回防御性列表 |
| 同目录 `f-slicer.d.ts` 第 198–225 行与 `f-timeline.d.ts` 的 setStyle/clearStyle | setStyle 只改当前 view，clearStyle 写 own undefined、恢复文档默认；source/cache/filter 不改。验证、权限、内容历史、Before 取消与事件仍由主任务统一命令承担 |

本批没有新增顶部入口，既有样式 gallery/显示设置继续按[切片器连接 UI](SHEET-25-slicer-connections-ui.md)、[时间线入口](SHEET-25-slicer-timeline-affine.md)及[顶部来源清点](SHEET-RIBBON-SOURCE-INVENTORY-2026-11-21.md)对齐。

## 模块与实际服务

按[完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)，新 [slicer-styles.ts](../../packages/office-studio/src/domain/slicer-styles.ts) 保存公开 author DTO。主任务为 WorkbookModel 协调新增 `slicerStyles` / `timelineStyles`；视图 `SlicerModel.style` 仍为可选字符串，默认解析不把派生值写入作者视图。

[slicer-style-catalog.ts](../../packages/office-studio/src/features/slicer/slicer-style-catalog.ts) 保留既有内置常量，新增文档范围的 list/has/get/default helpers。列表按内置顺序后接作者顺序并去重；getter 与全状态 projection 返回防御性副本。`getSlicerStyleType` 保持八种状态身份，hasData 为 false 不会清掉 selected。缺少/未知 view 引用显示使用文档默认，再回退各 Light1；未知作者名称仍保留，并返回 unresolvedName。

Slicer 完整 projection 分开 wholeStyle/headerRowStyle/items，Timeline 分开 wholeStyle/headerRowStyle/elements，缺少作者区域是空稀疏对象。颜色主题、可用性/hover 状态与数据筛选继续由既有服务驱动，不增加第二套 filter authority。

新 [slicer-style-resources.ts](../../packages/office-studio/src/features/slicer/slicer-style-resources.ts) 实现：

- normalization 与 project：保留 bg/cl null、bd null 和各边框 null、ht/vt null；theme 0–11、上划线 ol 的 t 0–19、border s 0–13、ht 0–6、vt 0–3、bl 0/1。非法默认、重复名称、未知/重复状态和非法格式拒绝。
- `importSlicerStyleResources`：有效家族归一化；完整原资源保存在 `metadata.slicerStyleResourcesSource`，扩展、无效资源和未知控件引用写入 `metadata.slicerStyleResourceDiagnostics`。原 metadata 保留。重新打开维持最初来源和其诊断，来源保留不复活后来删除的有效资源。
- read 与 author mutations：资源 read 返回副本；set/upsert/remove/default 先完整验证再原子修改。author 编辑中的未支持扩展拒绝，导入继续保留。不能删除内置样式、仍被控件引用的自定义样式或保留为默认但已移除的样式。
- mutation 只增加有效样式变化的 view revision；添加未使用样式不动 view revision，显式视图 override 不因文档默认改变而变化。cache、selection、连接、原生 Table filter、报告和图表不改；boolean 表示是否变化，无变化不伪造内容修改。权限、历史、取消与事件由主任务接统一命令。

主任务负责 `model-utils.ts` / `workbookNormalizer.ts` 导入接线、公开 Facade/设置、默认创建、Canvas 和 Ant 控件。无效来源不能继续作为有效 styles 字段强制断言。本文只记录领域核心和草稿落盘，未单独宣称上述产品管线已通过。

## 未运行草稿

新 [slicer-style-fixture.mjs](../../packages/office-studio/tests/fixtures/slicer-style-fixture.mjs) 复用[七条真实稀疏 native/report 记录](../../packages/office-studio/tests/fixtures/pivot-filter-fixture.mjs)，包含原生 Table、普通/owned 多报告、shared cache views 与 Timeline；样式是依据公开 DTO 的合成样本。primary 自定义定义覆盖全部成员/时间状态、颜色/theme、上划线、十种边框方向和 null 重置。

[slicer-style-resources.contract.test.mjs](../../packages/office-studio/tests/slicer-style-resources.contract.test.mjs) 静态清点 **17 条 test 声明、76 个计划用例，全部未运行**：

| 草稿范围 | 计划用例 |
| --- | ---: |
| 内置+自定义列表、防御性、跨家族引用 | 2 |
| 八种 selected/hasData/hovered 身份与投影 | 8 |
| 全格式与 null 保留、完整状态区域 | 2 |
| 默认/显式/未知引用，不物化作者字段 | 2 |
| 非法 formatting 原子拒绝 | 20 |
| 两家族非法默认/重复/状态/定义拒绝 | 22 |
| 扩展来源保留、稳定重开 provenance | 2 |
| 无效家族独立保留 | 2 |
| 未解析 view 引用诊断 | 2 |
| 发布资源的独立 native control 数组 | 1 |
| domain model/raw snapshot 导入与 JSON 重开 | 2 |
| Pivot/Table 默认编辑的 view 与过滤权威 | 2 |
| 两家族 upsert/no-op/防御性/有效视图 revision | 2 |
| 内置/默认/in-use 删除拒绝与 unused 删除 | 2 |
| 可选资源清除及 no-op | 2 |
| author 扩展拒绝、import 源保留 | 2 |
| 跨工作簿与 metadata 隔离 | 1 |

本批计数不与缓存选项、typed selection 或 Values position 草稿重复；没有运行结果和通过数。

## 剩余缺口

1. 实际 setStyle/clearStyle 的权限、Before 取消、单次内容历史、shared-view 隔离、通知和 Facade 重开仍须统一命令/浏览器验收；纯 author-resource 服务不代表新增样式编辑 UI 已交付。
2. Canvas/Ant 的完整默认值、全部状态颜色、hover/no-data、明暗 theme、边框/上划线/对齐、ConfigProvider 与焦点/IME 由主任务接线。本批未新增 UI，也未启动浏览器验证。
3. 上游 Canvas `_resolveFormatting` 主要消费 bg/cl/bd/bl；本地保留 ol/ht/vt 模型不等于已验证其所有 Office 可观察效果。类型/theme 的 12 项实际色值与语言切换待真实样本核实。
4. 当前 Office parser/writer 的原生 Slicer/Timeline style parts、关系、默认样式和独立文件重开未验收。保留在 JSON/metadata 不可冒称原生 XLSX 样式往返已完成。
5. 本批有界样式 normalization/projection 使用 TypeScript，源保留需要副本；没有 Worker/WASM、大样本、复制计量或性能/内存测量。不声称优化、零拷贝或达标。[WASM 平台契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)仍适用。

以上缺口继续属于完整 Sheets / Docs / Pro 范围。
