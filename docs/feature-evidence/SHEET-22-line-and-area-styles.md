# SHEET-22：折线与面积的 Line / Smooth / Step

状态：**P18 前端实现及 34 个新增测试草稿已落盘，未验收**。日期：2026-11-26。本批继续 [P17 图表面板 route](SHEET-22-chart-edit-panel-routing.md)，将真实线型控件接入二维折线、面积、组合系列和既有三维投影；不以控件出现、草稿数量或内部 JSON 保留证明全部功能完成。

## 固定来源与入口清点

对照 Univer **v1.6.0**，commit **`eaf4cfa04fabdbcfc922c692f3450459f8a29546`**。基础源码在 `E:\rust_test\univer`；Pro 依据为本项目固定版本的公开 ES / 类型快照，没有安装或运行商业 SDK。

| 已读取来源 | 可复查的公开契约 | 本地处理 / 未核实项 |
| --- | --- | --- |
| [chart-ui ES](../../.research.local/univer-1.6.0/packages/chart-ui/lib/es/index.js) 的 `Ui`、`Ji` 和 `ic / ChartLineAndAreaSection` | `Ui=[Line,Area,Scatter,Combination]`；`ic` 按 Scatter 或 `baseOn(Line/Area)` 判断；线型选项为 Line / Smooth / Step；写入 `{...area,lineStyle}`，保留其他 area 字段；空值选项 Break / Link / Zero | 本地按家族接入 11 个 Line / Area / Combination 图型，Scatter 只保留空值处理。能力声明与 `ic` 判断存在不同层次，Combination 的上游最终装配 / 真实效果没有完全核实 |
| [enum.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/enum.d.ts)，AreaLineStyle 约 437 行、InvalidValueType 约 228 行 | 字符串枚举 `line/smooth/step` 和 `zero/break/link` | 使用同一作者值和公开枚举，不用布尔 smooth 替代三种线型 |
| [area.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/chart-builder/configuration/area.d.ts) | `setAreaLineStyle` / `clearAreaLineStyle` 配置操作 | 本地专用 patch 复用已有 options / chart update |
| [f-area-chart-builder.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/builders/f-area-chart-builder.d.ts)；[f-area-three-d-chart-builder.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/builders/f-area-three-d-chart-builder.d.ts) | 二维、三维面积 builder 均有 set / clear；二维类型注释明确 clear 恢复点间直线 | 本地通用 ChartConfigBuilder 提供 set / clear，并扩展到所有适用 Line / Area / Combination；不声称 builder 类层次与上游相同 |
| [f-chart.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart.d.ts)；[f-chart-builder-base.d.ts](../../.research.local/univer-1.6.0/packages/engine-chart/lib/types/facade/f-chart-builder-base.d.ts) | live / builder 的 `setInvalidValueStrategy` / `resetInvalidValueStrategy`，枚举名 ChartInvalidValueType；builder 类型注释明确 reset 恢复 Break | 本地 reset 清除覆盖后采用 break；live 写方法沿用 `Promise<this>` 和真实异步命令，公开类型中的同步 `this` 返回语义与本地存在差异 |
| [ChartLineAndAreaSection.d.ts](../../.research.local/univer-1.6.0/packages/chart-ui/lib/types/editor/sections/ChartLineAndAreaSection.d.ts) | 线型 section 和可注入控件呈现方式 | 使用统一 `antdv-next` Form / Select / Button；不复制 React 或商业控件实现 |
| [ECharts 官方 line 配置](https://github.com/apache/echarts-doc/blob/master/en/option/series/line.md)，本批经 Context7 查阅 | `step` 支持 start / middle / end，`smoothMonotone` 支持 x / y | native adapter 用 `step:'middle'`、`smooth:true`、`smoothMonotone:'x'`。middle 与本地三次插值是本项目实现选择，不冒称已核实 Univer 内部算法 |

未发现可读取的固定 `engine-chart` ES 实现，当前依据包括其公开类型、`chart-ui` ES 和官方 ECharts 文档。上游内部平滑算法、Step 转折位置、属性 section 注入的默认配置、Combination 最终渲染及 3D 视觉结果仍需可访问演示 / Office 行为和真实样本补证；公开 builder 的 clear / reset 基准已由类型注释核实。缺少内部实现不缩减范围。

本批分类计数：**新增常驻 Ribbon 根 0、插件扩展根 0、溢出根 0、独立上下文菜单根 0**；复用 P17 的 `lineAndArea` route 1。属性编辑区替换旧二维面积开关，提供 **线型 Select 1（3 个值）和恢复默认动作 1**；复用空值 Select 1（3 个值）。按本地目录复查，4 个 Line、6 个 Area 和 1 个 Combination 支持线型，另 1 个 Scatter 只支持空值 section。这些值、动作和图型不能相加作为功能总数。

| 属性入口 | 本地位置与顺序 | 显隐 / 禁用 / 命令 |
| --- | --- | --- |
| 折线与面积线型 | 样式 → 图型专属设置；3D 相机 / 深度控件和柱条间距控件之后，其他图型专属项之前；Line / Area / Combination 的 route 同时展开 common、type、series 并聚焦 `area.lineStyle` | 按 Line / Area / Combination 家族显示，包括堆积、百分比、3D；Scatter 不显示。面板 disabled / busy / 权限沿用现有门禁；选择只改 draft，应用走 `sheet.command.update-chart` |
| 折线与面积线型恢复默认 | 同一 Form.Item 的既有 Ant Button | 新 lineStyle 或旧 smooth 有作者值时可执行；清除两者并保留其他字段；同一次应用进入内容历史 |
| 空值处理 | 样式 → 标题、图例与外观中的既有 Select | Scatter 的 lineAndArea route 聚焦 `invalidValueStrategy`；Break / Link / Zero 保持原数据准备与命令路径 |

本批没有更改顶部 classic / grid 根位置；属性抽屉布局复用已有清点与表单，不新增另一个 Ribbon 清点基线。

## 实际行为与分层

作者字段为 `options.area.lineStyle?: 'line'|'smooth'|'step'`。保留可序列化未知字段；值在 UI / Builder / Facade / options 边界校验。新线型优先于旧 `area.smooth`；专用写入和 reset 清除旧开关，避免保存重开后重新激活。仅 patch 旧 smooth 时退出新 lineStyle 模式，保留原调用的面积语义。与线型无关的 patch、切换到暂不使用该配置的图型都保留作者值。

旧 smooth 仅作用于实际面积系列。旧 Combination 可能是面积平滑、折线直线：Select 不伪选单一全局值，表单明确说明继承状态；显式选择后统一作用于其线 / 面积系列。`FChart.getAreaLineStyle()` 是本地全局读便捷方法；旧 Combination 的单个系列有效值应结合 `getInfo()` / `resolveChartAreaLineStyle(chart,seriesType)` 判断，不把它当作上游 mixed-state API。

二维和三维使用同一纯线段几何：Line 为原端点；Step 在 middle 同时切换顶边和堆积底边；Smooth 使用加权单调三次插值，每段 12 个区间，保留端点并限制在对应端点值域。不跨 Break / Zero 缺口借邻点，Link 可跨缺口；日期优先使用 sequenceIndex，兼容逆序、不规则间距和重复 X。重复 X 的 Smooth 回退直线。

二维画布绘制真实采样线段与面积填充，保留正负交叉拆分、系列 / 数据点样式、渐变、标记和数据标签；面积底边使用同一采样，不把阶梯顶边配上斜底边。Combination 柱形继续原路径。native ECharts 映射实际 smooth / step / connectNulls、堆积和百分比配置；native 与本地平滑实现的像素 / 曲率细节尚未对照验收。

三维先生成世界空间的完整系列坐标和共享线型，再投影，保留相机、分类 / 日期、深度、堆积、百分比和原始标记锚点。曲线细分面只绘制两条纵向边线，避免内部接缝；边线透明度只应用一次。隐藏或零线宽的三维 Line 不绘制带状面，标记仍由原路径处理。3D 继续是本项目投影 Canvas 场景，不宣称 ECharts 原生 3D 或图内 WebGPU 已验收。

Scatter 不增加连线：Break / Link 仍是散点；Zero 可将缺失 Y 转零，缺失 X 无法定位，继续排除。该行为在契约 / 浏览器草稿中使用真实准备与渲染入口，不以 UI 文案替代算法。

路径为 **Ant Select / Button → settings draft → chartSettingsPatch → chart.update / sheet.command.update-chart → 文档隔离的历史与模型 → 数据准备 / Canvas → 内部 JSON 保存重开**。线型仅改变作者绘制，`chartSettingsBindingKey` 继续复用已有准备数据；空值策略属于数据绑定依赖。没有新增解析器、第二套统计引擎、Worker 协议、WASM 内存接口或后端服务。

## 修改文件与公开接口

本批涉及 **22 个文件：13 源码、7 测试 / fixture、2 证据**。P18 新增 3 个源码、4 个测试文件和本证据；其余为现有图表归属文件的局部修改。共享 Facade、领域类型和 index 仅修改图表局部；中央清单由主任务同步。本批未改 ExcelWorkbench、package / lock、Rust 或 Office writer。

| 源码 | 实际改动 |
| --- | --- |
| [chart-line-and-area.ts](../../packages/office-studio/src/features/charts/chart-line-and-area.ts)（新增） | 公共枚举、默认值、类型能力、边界校验、有效线型和专用 patch |
| [chart-line-and-area-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-line-and-area-settings-controls.ts)（新增） | 真实 Select / reset 描述、适用性、旧 Combination 继承状态 |
| [chart-line-path.ts](../../packages/office-studio/src/features/charts/chart-line-path.ts)（新增） | 可独立测试的连接判断、直线 / 阶梯 / 单调曲线与上下边界 |
| [model.ts](../../packages/office-studio/src/domain/model.ts) | `ChartAreaLineStyle`、`ChartInvalidValueStrategy`、`ChartLineAndAreaOptions`，兼容旧 smooth |
| [chart-option-validation.ts](../../packages/office-studio/src/features/charts/chart-option-validation.ts) | 新线型枚举验证，保留已有作者校验 |
| [chart-options.ts](../../packages/office-studio/src/features/charts/chart-options.ts) | 新字段、旧开关和 explicit undefined 的 merge / reset 规则 |
| [chart-type-settings-controls.ts](../../packages/office-studio/src/features/charts/chart-type-settings-controls.ts) | 替换二维面积 Switch，接入统一线型 Select |
| [chart-edit-panel-route.ts](../../packages/office-studio/src/features/charts/chart-edit-panel-route.ts) | 11 个适用图型聚焦真实线型；Scatter 保留空值定位 |
| [chart-draw.ts](../../packages/office-studio/src/features/charts/chart-draw.ts) | 二维真实线型 / 面积底边、共享缺口连接和渐变 bounds |
| [echarts-option.ts](../../packages/office-studio/src/features/charts/echarts-option.ts) | native smooth / smoothMonotone / middle step，保留数据和系列语义 |
| [chart-3d.ts](../../packages/office-studio/src/features/charts/chart-3d.ts) | 世界空间共享几何、采样面积填充、纵向边线和透明度处理 |
| [facade.ts](../../packages/office-studio/src/api/facade.ts) | FOffice 枚举、Builder / live set / clear、空值 set / reset、本地读方法 |
| [index.ts](../../packages/office-studio/src/index.ts) | 导出枚举、默认值、能力 / 解析函数及 3 个领域类型 |

公开新增：`FOffice.Enum.ChartAreaLineStyle`、`ChartInvalidValueType`；Builder 与 FChart 的 `setAreaLineStyle(value)` / `clearAreaLineStyle()`、`setInvalidValueStrategy(value)` / `resetInvalidValueStrategy()`；FChart 的 `getAreaLineStyle()` / `getInvalidValueStrategy()`。Builder 同步返回 this，live 写方法返回 `Promise<this>`。专用线型方法拒绝不适用的 Scatter / Column 等类型；通用 options 可保留暂不参与当前图型的配置。get 方法及通用 Line / Combination 线型 setter 是本项目便捷扩展，不冒称上游所有 facade 类均有同名签名。

## 草稿与验收状态

新增 **34 个场景草稿**，按实际 `test` 入口手工计数；循环参数不另计，改写 P17 原入口也不另计：

| 新测试文件 | 场景 | 实际验证意图 | 运行 |
| --- | ---: | --- | --- |
| [chart-line-and-area.contract.test.mjs](../../packages/office-studio/tests/chart-line-and-area.contract.test.mjs) | 13 | 枚举 / 适用性 / 无效输入，未知字段，legacy 优先 / reset，真实控件 writer / draft / binding，旧 Combination 混合态，双边 Step、断点外邻点、逆序 / 重复日期和连接错误 | 未运行 |
| [chart-line-and-area-rendering.contract.test.mjs](../../packages/office-studio/tests/chart-line-and-area-rendering.contract.test.mjs) | 8 | native 与共享 Canvas 真实几何，堆积 / 百分比上下边界，组合柱形 / 次轴，空值三策略，逆序 / 重复日期，3D 面 / 边线 / 透明度和 Scatter 缺失 X / Y | 未运行 |
| [chart-line-and-area-facade.contract.test.mjs](../../packages/office-studio/tests/chart-line-and-area-facade.contract.test.mjs) | 6 | Builder / 枚举 / 输入隔离，单次真实更新、内容历史、数据读与 JSON 重开，权限 / veto / 无效值，类型切换、跨 workbook 历史和关闭文档 | 未运行 |
| [chart-line-and-area.browser.test.mjs](../../packages/office-studio/tests/chart-line-and-area.browser.test.mjs) | 7 | 真 ExcelWorkbench / Ant Select、Line 曲线、面积 / 组合 Step、空值和 Scatter、3D reset / undo / redo / JSON、ConfigProvider 明暗 / 回焦、IME 导航与一次提交 | 未运行 |

另改 [P17 route 契约](../../packages/office-studio/tests/chart-edit-panel-route.contract.test.mjs) 的原第 8 个场景及 [P17 route 浏览器](../../packages/office-studio/tests/chart-edit-panel-routing.browser.test.mjs) 的原第 4 个场景，使其定位 / 编辑真实 `area.lineStyle`；**P17 原新增 33 个草稿数量不变**。[chart-object-actions.html](../../packages/office-studio/tests/fixtures/chart-object-actions.html) 增加可选 rows / nativeOptions，保存真实 ECharts option，并通过真实 FChart 提供 scene / data / shape / invalidValues 观察；未替换计算器、Canvas、编辑器或 Ant 控件。另一修改证据为 [P17 后续记录](SHEET-22-chart-edit-panel-routing.md)。

涉及的可替换 UI 全部复用 `antdv-next` Form / FormItem / Select / Button、既有 Drawer / Collapse 和输入组件，走统一 ConfigProvider / token；没有新增原生表单或输入桥。浏览器草稿关闭 navigator.gpu，验证意图为 Canvas2D；IME 场景使用合成 composition / input 事件，不能代替真实系统输入法验收。

本批只执行文件 / 固定来源读取、`rg` 检索、手工静态审阅和源码 / 文档写入。**未执行测试、类型检查、构建、浏览器、dev / Vite、Cargo / WASM、基准、服务验收或 `git diff --check` 等自动检查；未提交、发布、更新依赖或修改锁文件。** 所有草稿、焦点 / 主题 / IME、真实文件和性能仍为未验收。

## 剩余缺口

1. 34 个新增草稿及相关旧草稿尚未运行。所有图型、普通 / 大样本、字体 / DPR、明暗主题、焦点恢复、真实 OS 输入法、读取权限改变、缓存 / 取消和设备失效仍待统一浏览器验收；不能把静态作者路径当作运行证明。
2. 上游属性 section 注入默认值、Step 转折位置、平滑曲率、Combination 最终装配和 3D 视觉结果未完全核实。本地 line / break 恢复基准与公开 builder clear / reset 注释一致；middle Step 和加权单调插值是明确的本地实现选择，仍须固定演示与 Office 可观察行为对照，不虚报等价内部实现。
3. 本批未扩展可编辑 XLSX 线型适配、Office 读取或写出。既有 [xlsx_chart_warnings.rs](../../crates/office-writer/src/xlsx_chart_warnings.rs) 约 35 行已记录传统原生面积无 smooth-boundary 设置；新的 `area.lineStyle`，尤其 Step 的原生映射、诊断及真实 Excel 重开仍待本项目解析 / writer 补齐。保留作者 JSON / manifest 或导出图片均不能替代 Office 验收。
4. P17 的全部 Canvas element 命中 / 直接编辑、任意 runtime seriesId 对照，以及 P16 的 chart-sheet、GRID / CHART paste placement、跨 workbook 来源依赖和一般 drawing 闭环仍继续在完整前端范围内，本批不将它们记为完成。
5. 首屏、滚动、编辑、缩放、线段采样、Worker / WASM 任务、跨边界字节、帧时间、长任务和内存尚无同数据集实测。当前纯几何仍在既有 TypeScript 绘制层，特别大数据的预算、路径阈值和 Worker / Rust 接入仍需基准决定；不声称性能已优化、已达标或全链路零拷贝。

以上仍属于完整 Sheets / Docs / Pro 对齐目标，状态保持“未验收”，没有缩减范围。
