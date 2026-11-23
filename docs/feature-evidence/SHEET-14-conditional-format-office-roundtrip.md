# SHEET-14 条件格式 Office 来源与往返

记录日期：2026-11-22。验收状态：**未验收**。本批只做静态源码与公开资料读取、实现和测试草稿；未运行测试、类型检查、前端构建、Rust/WASM、浏览器、语法检查、性能基准或 `git diff --check`。下文的实现描述仅指代码已落盘，不构成可编译、可运行或 Office 互操作通过的结论。

本页记录 SHEET-14 与 SHEET-29 的条件格式读写链路，以及格式刷接入该模型后的来源保留。完整功能范围仍遵守[项目标准](../../PROJECT_EXECUTION_STANDARD.md)、[源码规则](../UNIVER_REFERENCE_RULES.md)、[功能清单](../UNIVER_FEATURE_PARITY.md)和[渲染与计算标准](../RENDERING_AND_COMPUTE_STANDARD.md)，包括全部 Sheets / Docs 及 Pro 目标。前一阶段的入口、命令、公式基准、历史和浏览器草稿见[格式刷与条件格式编辑](SHEET-04-format-painter.md)；2026-11-21 的验证结果仅属于[当日数字比较切片](conditional-formatting-2026-11-21.md)。

## 1. 固定来源与公开依据

Univer 参考为 **v1.6.0**，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，本地只读根 `E:\rust_test\univer`。以下路径相对于该根，链接均固定到同一提交。

| 来源 | 读取结论与本地选择 |
| --- | --- |
| [sheets-conditional-formatting/src/models/type.ts][U1] | `cfId` 与范围、条件、样式、`stopIfTrue` 分开；文本 literal 与公式是不同字段；`IAverageHighlightCell` 允许 `> / >= / < / <= / = / <>`。本地保留独立 ID，并分别保存 literal 与标准 predicate。 |
| [sheets-conditional-formatting/src/models/conditional-formatting-rule-model.ts][U2] | 规则数组顺序表达叠加顺序，新规则插在首部。本地使用显式 priority 加稳定数组顺序；稀疏及相同 priority 不因普通列表编辑全量重编号。 |
| [sheets-conditional-formatting-ui/src/menu/schema.ts][U3]、[menu/manage-rule.ts][U4] | DATA 管理入口、INSERT 快插、selector 子项与权限状态分别存在；单个面板按钮不能代替全部入口。 |
| [sheets-conditional-formatting/src/commands/commands/set-cf.command.ts][U5]、[UI RuleEdit.tsx][U6]、[RuleList.tsx][U7] | 上游逐规则命令、mutation、逆 mutation 和历史；本地面板以完整规则列表原子提交，共用命令、范围权限及工作簿历史。 |
| [UI rule-edit/Rank.tsx][U8] | 固定版面板只列前/后几项及高于/低于平均值，且会把 inclusive 平均值归为 strict。上述六种比较来自模型契约；本地面板分别保留六种条件，不复制这个收窄行为。 |
| [Open XML SDK v3.3.0 的 base schema][O1] 与 [x14 schema][O2] | 标准 cfRule、dxf、x14 cfRule/dataBar/cfIcon/extLst 的类型、子节点和 GUID 约束。固定 schema 为本次映射依据；没有运行 SDK validator。 |
| Microsoft Learn 的 [cfRule][O3]、[dxf][O4]、[Office 2010 cfRule][O5]、[dataBar][O6] | 公开字段及条件格式规则、差异样式和增强数据条的可观察契约；文档为动态页面，不能代替文件样本验收。 |
| MS-XLSX [ST_CfvoType][O8]、[CT_Cfvo][O9] | 2026-11-22 静态读取确认 `autoMin = min(0, 区域最小值)`、`autoMax = max(0, 区域最大值)`；自动阈值只允许用于数据条，最小端不得用 autoMax，最大端不得用 autoMin，自动阈值没有 f 操作数。本地前端与 Rust 投影校验按此限定用途。 |
| [XlsxWriter RELEASE_3.2.9 worksheet.py][O7] | base 与 linked x14 数据条写出、GUID 关联和颜色/轴字段的独立公开实现参考；本项目仍使用自己的 Rust writer。 |
| [quick-xml 0.41.0 NsReader][O10] | 续作通过 Context7 与本机该版本源码核对 `BufRead`、resolved event、namespace 和属性解码接口；本项目新增流式 selector，未运行编译或性能验证。 |

续作还通过 Context7 `/dotnet/open-xml-sdk` 读取了 extension 的公开 schema 入口；检索返回的 `main` 文件仅用于定位，版本结论以以上 v3.3.0 固定文件为准。没有引入 Open XML SDK、XlsxWriter 或 Univer 商业 SDK 作为运行依赖。

本批新增 Ribbon 根 **0**。条件格式入口清点沿用[既有专项清点](SHEET-04-format-painter.md#1-固定来源与入口清点)及[完整入口表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)：上游 DATA 管理根 1、INSERT 快插根 2、管理 selector 子项 10 分开统计；本地管理按钮为数据 / 规则、classic 第 2 项、grid `row: 2, column: 1`，使用 `sheet.operation.open-conditional-formatting`。完整 selector、快插、选区过滤和权限展示仍有差异。本批添加平均值及自动数据条阈值子选项，不将它们计作新的 Ribbon 根或全部入口完成证明。

## 2. 模型、身份与权威数据

`SheetConditionalFormatRule` 新增可选 `officeSource`，来源版本独立为 1；模型继续保留公式基准和未知 JSON。来源契约在 TypeScript 领域模型、normalizer 与 Rust `OfficeSource` 中对应：

| 字段 | 含义 |
| --- | --- |
| `version: 1`、`format: 'xlsx'` | 来源记录版本与文件格式。 |
| `originalId`、`storage: 'base' \| 'x14'` | 导入时的本地身份与规则存储位置；与 priority、dxfId 或当前本地 ID 分开。 |
| `ruleXml`、可选 `groupXml` | 原 cfRule 与分组来源片段，包括 namespace、属性、公式、扩展和混合内容。 |
| 可选 `differentialFormatXml`、`x14Xml`、`themeXml` | 原差异样式、关联增强分组及其依赖的主题。独立 x14 的 inline dxf 也可作为差异样式来源。 |
| `baseline` | 仅含当前模型可检查的字段；排除 ID、officeSource、metadata 及未知 JSON，避免把来源记录或未知扩展变化误认为条件字段变化。 |

原始片段经过 XML 解析与重序列化；保留解析后的内容、namespace、注释和处理指令，不承诺与原压缩包或 XML 字节一致。CDATA 与 entity 可以按等价文本写出，属性顺序也可变化。

本地写出在标准 rule 的 extLst 中附加版本化模型投影，保存原本地 ID、type/subType 表示、公式拼写、稀疏/相同 priority、范围基准及未知 JSON；投影移除嵌套 `officeSource`，避免每次重开递归增长。再次导入时，rule、group、dxf、linked x14 和 theme 的来源指纹一致才恢复该投影的已知字段；不一致时以标准 Office XML 的当前字段为准，只分别保留合法本地 ID 和未映射 JSON。

指纹用于检测意外陈旧数据，**不是安全签名**。`valid_known_fields` / `assertConditionalFormatKnownFields` 校验可检查的基本字段形状；`valid_saved_model` 对可替代标准解析结果的私有模型进一步校验公式、排名、日期期间、阈值和图标。无效投影不能因指纹一致覆盖合法标准 XML。

自动阈值在 `valid_saved_model` 中仅限数据条的对应端点；放入色阶、图标集或互换两端的自动类型不能覆盖标准 XML，也不能附加为有效本地投影。导入时标准规则若含这些无效 visual 条件，沿既有校验路径保留为 office 来源规则，仍属于未知条件计算缺口。

没有本项目投影的 Office 规则以原生 revision UID、x14 ID、linked GUID 或稳定来源摘要生成本地身份。复制导致当前 ID 不同于 originalId、或同工作表来源 GUID 重复时，writer 分配新 GUID；原规则尽量保留原 GUID，输出统一大写。解析器区分重复本地身份并发 warning。跨工作表复制的全局原生 GUID 唯一性还未专项验收。

扩展 URI 分开使用：

| 用途 | URI |
| --- | --- |
| base rule 与 x14 关联 | `{B025F937-C7B1-47D3-B67F-A62EFF666E3E}` |
| worksheet 条件格式扩展 | `{78C0D931-6437-407D-A8EE-F0AAD7539E65}` |
| 本项目版本 1 模型投影 | `{A7174764-91A7-4C36-96D4-5C1391AF06EB}` |

x14:extLst 的 extension 子节点仍是 base SpreadsheetML `ext`；私有模型位于该扩展允许的独立 namespace 中，没有把本地字段直接添加为标准 cfRule 属性。

## 3. 已落盘的读写合并

导入沿 `ConditionalFormatSidecar` → `SheetPayload.conditionalFormats` → `normalizeWorkbook` → 可编辑规则模型。写出沿 `ConditionalFormatInput` → 公式基准投影 → 来源字段合并 → StyleRegistry / worksheet / theme package。Office 转换仍使用本项目 Rust/WASM 读写链路，不依赖 Univer Server、Exchange 或私有协议。

| 规则或资源 | 本批行为与边界 |
| --- | --- |
| 数值比较 / expression | 已知 formula1、formula2 及表达式 value 以当前模型为权威。安全基准时平移普通 A1 引用；无共同安全基准时以 INDEX / ROW / COLUMN 投影逐格引用，保留同一条规则的统计范围。该投影不是所有结构变更的引用处理器。 |
| 文本、空白及错误 predicate | matching literal 与标准辅助公式分别保存。未改 literal 时保留并移动标准 predicate 的引用，literal 中的 `A1` 不变；修改 literal 或条件后重建 predicate，跳过 vendor 内容。 |
| 平均值 | `>= / <=` 使用 aboveAverage / equalAverage 标志；`= / <>` 用标准 expression 与绝对、去重的 AVERAGE 范围表示。面板分别映射六种比较，保留 `highlightCell/average` subtype，平均值编辑不再写入排名项数。本地非有限输入和非有限平均值不参与匹配。 |
| dxf | 按编辑过的字段合并 fill、fontColor、bold、italic、underline、strike。只改 bold 时保留 name/size、double underline、theme/tint、numFmt、alignment、border、protection 及 vendor 扩展；清除字段也是显式修改。未映射的样式仍以来源 XML 保留，不能据此声称 Canvas 已绘制这些样式。 |
| 数据条 | base 与 linked x14 配置先合并；可选 x14 fillColor 缺失时不丢掉 base 颜色。x14 使用 fillColor、xm:f，逐字段更新负色、边框、轴、方向、长度及显示标志。`autoMin/autoMax` 可由面板编辑和命令提交，按含零边界计算；base 使用 min/max 兼容映射，linked x14 保留自动类型。全零区域没有正长度数据条。未映射的 bar/scale gte 不因阈值编辑被删除。高级字段及自动数据条的最终 Canvas / Office 效果未验收。 |
| 色阶 | 阈值与颜色分别比较；仅修改中间阈值保留其他原颜色和阈值 XML。两/三色切换保留相应低、高端来源节点及扩展。 |
| 图标 | local glyph 身份与 Office 图标 ID 分开映射，rating 与箭头的顺序分别处理。只改阈值时保留未改 cfIcon、NoIcons 及 vendor 属性；独立 x14 支持 inline dxf 与严格 gte。`_5Felling` 没有合法 Office stock 等价，写出合法 rating fallback 并 warning，私有扩展保留本地 glyph；不是 Office 原生等价实现。 |
| linked x14 | 关联 GUID 保持同步；关联增强 rule 不写 priority，独立 rule 使用独立 priority。更新 xm:sqref 不再删除 cfRule。复制时关联 ID 与来源 revision UID 按所支持路径重新分配。 |
| theme | 可检查的 sysClr/srgbClr 与 theme/tint 投影为本地颜色，支持读取 transitional / strict DrawingML。保留并写出 `xl/theme/theme1.xml`、content type 和 workbook relationship。不同来源主题冲突真实拒绝写出；主题合并/重绑定没有实现。 |
| worksheet extLst | CF、既有 sparkline 与 axis 扩展由 `xlsx.rs` 装配到唯一 worksheet extLst。使用末尾 closing extLst 位置追加，避免把内容插入嵌套列表；cfRule 自身的 extLst 保持独立。axis 来源记录在追加扩展前生成，不把最终 worksheet 扩展递归纳入自身来源。本批只有静态装配与草稿，三类扩展共存的 Office 程序验收未执行。 |

来源中声明了 dxfId 却没有对应 dxf 时，writer 在任何新样式生成前拒绝写出；用户补一个 bold/fill 不再绕过该拒绝。此类损坏来源无法恢复原差异样式，不能把新生成样式当作保真替代。

### 3.1 流式选择条件格式与相关资源

本轮补齐 `conditional_format_xml_selection.rs`。`extract<R: BufRead>` 顺序扫描 XML part 至 EOF，selector 接收已解析 namespace 的 `XmlName` 路径与瞬时名称/属性 `Element`，只为选中的最外层子树建立完整树；返回根 metadata、独立 fragments 及选中内容的字节/元素计量。未选中的单元格、xf 和 dxf 内容不建立树、不消耗 CF 片段或选中内容预算。末尾 XML 损坏、重复根、非法属性或 namespace 等读取失败仍拒绝整个结果，不能发布此前已捕获的片段。

workbook 只捕获 `sheets/sheet`；relationships 只保存工作表与 styles/theme 所需的内部关系 metadata。worksheet 捕获直接 base `conditionalFormatting`，以及 worksheet `extLst` 内完整 x14 分组。先扫描各表收集实际 `dxfId`，再从 styles 按原始顺序捕获引用的 dxf，以 `HashMap<usize, Element>` 保存原索引；无引用 dxf 不读取 styles。只有已支持的 SpreadsheetML 颜色或 major/minor font scheme 声明主题依赖时才读取 theme；普通 vendor namespace 的 `theme` 属性、无颜色语义的 payload 与 `scheme val="none"` 不构成该依赖。颜色/tint 和 major/minor 字体资源校验已移到共同 `validate_theme_references`，parser 与 writer 共用，writer 在任何替代 dxf 注册前拒绝缺失/空主题中的不可用声明资源。无 CF 或无相关引用时，CF sidecar 不读取损坏/缺失的无关 styles/theme；核心 `rxls::Workbook::open` 仍在 sidecar 前运行，此项不代表整份 Office 文件可以包含损坏的公共样式。未映射的主题表达式和 vendor 资源协议仍需单独核实。

脱离祖先的片段保存全部有效 namespace bindings、混合文本、注释、PI，以及继承的 MC 和 `xml:lang/space/base` 属性。继承 MC 的 prefix/QName 列表在下层前缀重绑定后按原 namespace 重新绑定，供 vendor 内容只在属性值中引用的 namespace 使用。`xml:base` 当前复制最近继承值；多层相对 base 的组合语义未核实，保持后续验收项。

`parse_rules_tree`、来源生成和资源读取都返回真实 `Result`；损坏范围、引用资源缺失、来源超限等失败保存在 `ConditionalFormatSidecar.failure`，失败的 `by_sheet` 为空。生产路径不再使用返回空 Vec 的吞错包装；旧包装仅供 `cfg(test)` 夹具使用。`parser.rs` 已接线在导入发布前拒绝 CF 失败，原样保留详情中的错误前缀与已有 part 路径：

| CF 导入详情前缀 | 解析 envelope 错误码 | 含义 |
| --- | --- | --- |
| `CF_SOURCE_FRAGMENT_LIMIT` | `ParserFailure` | 单片段字节、深度、元素或面积数超限。 |
| `CF_SOURCE_TOTAL_LIMIT` | `ParserFailure` | 选中内容、整本资源、序列化模型或独立 scanner guard 超限。 |
| `CF_SOURCE_XML_INVALID` | `InvalidOfficeFile` | 损坏 XML、namespace、sqref、优先级或 dxfId 等来源数据无效。 |
| `CF_SOURCE_REQUIRED_RESOURCE` | `InvalidOfficeFile` | 已声明的 worksheet、dxf 或 theme 等依赖不可用。 |
| `CF_SOURCE_PACKAGE_INVALID` | `InvalidOfficeFile` | ZIP/相关内部关系等包结构错误。 |

### 3.2 来源与模型预算

预算以 UTF-8 字节计，不使用 JavaScript UTF-16 长度作为字节替代。canonical XML 的计量和 `render()` 共用遍历，包含实体转义、继承 namespace、注释与 PI；parser 保存来源与 writer 读取来源时先检查片段深度、元素和 canonical 字节，最终装配后的 CF 另由 worksheet guard 检查。Rust 使用 serde counting writer 计量实际 JSON 表示，TypeScript 使用扫描 helper 计量 `JSON.stringify` 的字节/转义表示，计量不额外生成整份 part XML 或完整模型 JSON。

| 范围 | 默认限制与口径 |
| --- | --- |
| 每个选中 XML 片段 | 输入子树及脱离祖先后的 canonical UTF-8 各 **16 MiB**；深度 **128**、元素 **262,144**。 |
| 每个 part 的选中内容 | canonical **64 MiB / 1,048,576 元素**；不是整份 worksheet/styles 的文件大小限制。 |
| 未选中内容的 scanner | 独立深度 guard **4,096**；不占用片段深度预算。 |
| 整本 CF 公共资源 | 所有 worksheet CF 分组、实际引用的 dxf 和选中 theme 累计 **128 MiB / 2,097,152 元素**；workbook/relationships metadata 独立计量，不混入这项资源总量。 |
| 单个版本 1 OfficeSource | 五个 XML 字段各受片段字节限制，总和 **80 MiB**；`originalId` **4,096 UTF-8 字节**。 |
| 整本序列化 CF 模型 | **256 MiB**，计入数组括号、逗号、每条规则/来源/未知 JSON，以及各规则中重复保留的 dxf/theme 等来源副本。 |

`finish_model` 先检查生成片段与来源，再计入累计模型；重复 ID 追加 ASCII 后缀增加的序列化字节也计入。writer 的 `theme_for_workbook` 在任何 package XML 添加前检查整本输入模型预算和来源主题。`xlsx.rs` 在 pivot/table/axis/sparkline 装配后的最终 worksheet XML 上调用 `output_source_usage`，跨表 checked-add 累计；公共 part 添加前调用 `validate_output_resource_budget`，加入实际注册输出的 dxf 与复用的 theme，同时单独检查 styles part 的选中 dxf **64 MiB / 1,048,576 元素**。单 theme 受 **16 MiB / 128 / 262,144** 片段限制并仅计入整本资源一次。失败不返回部分 ZIP 作为成功产物；`ConditionalOutput` 仍只有 `base` / `extension`，没有添加统计字段。writer 仍会装配最终 CF/worksheet 字符串后再执行对应最终 guard，峰值分配与复制尚未验收。

TypeScript 的来源 normalizer 在最终深克隆前检查 UTF-8 XML、已过滤 baseline 和全部保留来源扩展；命令在候选列表克隆前检查单规则及列表预算。整本 helper `assertWorkbookConditionalFormatBudget(sheets, limit?)` 同时支持 parser sheet 数组与 Facade sheet map，逐表用 `conditionalFormatRuleArraySerializedSize` 在 `conditionalFormats` 属性上下文计量完整规则数组，保留各表实际索引。累计预算使用一个逻辑数组：扣除每表两侧括号，在非空表之间补一个逗号，空表不增加分隔符。该 helper 已接入 `normalizeWorkbook` 的克隆前/归一化后、`ensureWorkbookId` / `workbookFromInput` 的克隆前、`commitWorkbook` 的发布前及 `exportOfficeModel` 的 JSON 前。部分数组/map 快照使用 `clone({ conditionalFormats: input.conditionalFormats }).conditionalFormats`，保留规则数组的属性上下文和来源。失败后不发布模型或历史的合同草稿由整本工作簿专项测试覆盖，尚未运行。

TypeScript JSON helper 面向领域 JSON：循环、BigInt、过深嵌套及无 JSON 转换的非领域 prototype 会真实拒绝；支持确定性的 `toJSON`（包括通过该转换保留的函数属性）、重复共享引用、对象 undefined 字段省略、数组空位/undefined 转为 null 和非有限数字转为 null。规则数组计量在单次调用中只执行一次数组根转换，返回非数组会显式拒绝；每条规则的转换 key 来自该表的实际数组索引。通用 iterable helper 仍按单一数组的索引计量，与整本 helper 的每表上下文分开。`normalizeWorkbook` 在整本预检后以 `cloneModel({ conditionalFormats: raw.conditionalFormats }).conditionalFormats` 先物化相同属性表示，再逐条投影领域字段，避免规则或扩展换到另一个 JSON 根时改变确定性转换结果；部分 Facade 快照也使用相同属性上下文。上述预算与转换已有契约草稿，尚未执行；任意 getter、有副作用 `toJSON`、其他入口的转换和克隆峰值仍未验收。流式提取和计量保护阈值仅为代码事实；quick-xml 单个巨大 Text token 仍可分配大 buffer，扫描还可能生成瞬时解码文本，不能声称内存或性能优化已达标。

## 4. 未知 Office 条件的实际边界

未识别规则、未支持图标、非零 stdDev、无效排名/日期条件、无效 visual 基本字段、超出本地两个 formula 字段的额外标准公式，以及 base 与 linked x14 的 type/body/scope 不匹配，按已覆盖路径保留为 `type: 'office'` 并 warning。原 XML 与关联来源保留；本地计算明确退出，不调用 expression/numeric 回调，也不让未知规则的 stopIfTrue 阻断后续已知规则。writer 同时拒绝将包含额外标准公式的来源伪装为已知编辑模型，避免公式合并静默删除第三个节点。

共享 CF 命令允许修改该规则的范围、差异样式、priority、叠加标志和删除；未知 type/subType、operator、formula1/formula2、value、排名标志及 config 的语义变化真实拒绝，失败不提交部分内容。面板不提供未知条件的语义编辑，来源规则仍可执行上述保留型编辑。这是当前开发边界，未知条件的计算及公开公式编辑协议仍是待开发项，不是最终范围取消。

内部 `project_opaque` 仅处理标准 formula/cfvo 路径的原点投影，跳过 vendor ext。关联来源 scope 不同时，范围变换拒绝。vendor opaque XML 中的 A1、结构化引用或其他坐标协议尚未转换；原点平移不能当作插删行列、轴资源、多段移动或剪切粘贴的完整引用变换验收。

## 5. UI、命令与接口

条件格式面板沿用 `sheet.command.set-conditional-formats`，输入为完整 rules 与可选 unitId、sheetId/subUnitId、expectedRevision；受影响旧/新范围、查看/编辑权限、版本与原子历史沿用[上一阶段合同](SHEET-04-format-painter.md#32-条件格式面板的真实提交闭环)。本批没有改变该命令或面板 submit/closed 的返回契约。

面板的 Drawer、Form / FormItem、Input、InputNumber、Select、Checkbox、ColorPicker、Button、Tooltip 与 Alert 均静态可见为 `antdv-next` 控件；阈值子组件也使用库控件。Drawer 保持 **`:after-open-change` prop**，关闭回调由宿主处理焦点恢复，沿用统一 ConfigProvider 与主题 token。本批没有新增原生表单桥，也没有运行明暗主题、弹层焦点或真实输入法验收。

`ConditionalThresholdInput` 仅在数据条相应端点提供“自动最小值 / 自动最大值”，这些类型和普通 min/max 不显示数值 Input；切换到无操作数边界时发出 `{ type }`。色阶、图标集没有自动选项，命令也拒绝自动阈值及端点互换，包含从数据条复制未变阈值后切换类型的情况。计算器仅在数据条分支启用自动边界，避免在其他 visual 类型中产生新的含零语义。

接口变化包括：

- 领域规则的可选 officeSource、`type: 'office'`、来源 baseline 及数据条高级 config 字段。
- Rust parser 的 conditionalFormats 使用可保留扩展的 JSON 值；writer 的规则及 style 保留未映射 JSON，来源以共同 `OfficeSource` 类型接入。
- 共同 `conditional_format_xml`、`conditional_format_source` 与 `conditional_format_color` 模块承载解析、投影、校验、身份与颜色，不在组件或 store 中实现 Office 读写。
- writer 的 `ConditionalOutput` 分开输出 base 与 worksheet extension；xlsx 装配同时接入 preserved theme。
- 规则编辑模块增加 `conditionalFormatAverageEditorOperator` / `conditionalFormatAverageRuleOperator`，供面板的双向平均值条件映射复用。
- `validateConditionalThreshold` 新增可选 `allowAutomatic = false`；只有数据条调用放行自动类型。阈值子组件新增可选 `automatic` prop，仍使用原 `update:value` 契约；命令 input/submit 与 `conditionalFormatStyleAt` 的公开接口未改变。
- 流式 selector 的签名为 `extract<R: BufRead>(input, limits, FnMut(&[XmlName], &Element) -> Result<bool, String>) -> Result<XmlSelection, String>`；共同 `Element` 新增 `rendered_size` / `validate_fragment` / `detached_child`。
- 共同来源模块新增 `serialized_model_size`、`ConditionalFormatModelBudget::{new, consume, used, reserve_bytes}` 与资源/model 常量；writer 新增 `output_source_usage(xml) -> Result<(usize, usize), ExportError>`、`validate_output_resource_budget(bytes, elements, registry, theme) -> Result<(), ExportError>`。
- 前端新增 UTF-8、JSON 和整本规则预算 helper；`conditionalFormatSerializedSize(value, limit?, rootKey = '')` 的第三参数保留 `toJSON` 的属性上下文。`conditionalFormatRuleArraySerializedSize(values, limit?, rootKey = 'conditionalFormats')` 另外要求根转换结果仍是数组，供整本 helper 保留各表索引并累计单个逻辑数组预算。来源版本仍为 1，未知扩展继续保留。来源基线的公开返回仍深克隆，内部预算使用不先克隆的 projection。
- 共同颜色模块新增 `validate_theme_references(node, theme) -> Result<(), String>`；parser 和 writer 在同一声明依赖规则下校验可检查的主题资源，vendor 普通属性不冒充 Office 颜色/字体引用。

共享 `excelViewAdapter`、单格内容命令、公式引用结构处理器、axis 事务和全局对齐清单由根任务维护；本文只记录它们涉及 CF 的读写预算与 snapshot 接线，不据此交付其他功能。

## 6. 测试草稿与静态计数

2026-11-22 使用 `rg -c '^\s*#\[test\]'` / `rg -c '^test\('` 静态重数；只统计声明，不执行 runner，循环内样本不另计。Office 来源/字段往返前阶段新增 **36** 个声明，当时七个相关文件合计 **58**。本轮流式与来源预算新增 **46**（Rust **35**、TypeScript source-budget **11**），整本工作簿接线另新增 **8**；本轮共新增 **54**。下表是相关十一份文件当前总声明数 **112**，包含历史已有声明，不能与任何一轮新增数量混用。

| 文件 | 当前声明数 | 结果 |
| --- | ---: | --- |
| [conditional_format_xml.rs](../../crates/office-io-common/src/conditional_format_xml.rs) | 6 | 未运行 |
| [conditional_format_xml_selection.rs](../../crates/office-io-common/src/conditional_format_xml_selection.rs) | 10 | 未运行 |
| [conditional_format_source.rs](../../crates/office-io-common/src/conditional_format_source.rs) | 10 | 未运行 |
| [conditional_format_color.rs](../../crates/office-io-common/src/conditional_format_color.rs) | 3 | 未运行 |
| [xlsx_conditional_source.rs](../../crates/office-writer/src/xlsx_conditional_source.rs) | 12 | 未运行 |
| [xlsx_conditional_formula.rs](../../crates/office-writer/src/xlsx_conditional_formula.rs) | 7 | 未运行 |
| [conditional_format_sidecar.rs](../../crates/office-parser/src/conditional_format_sidecar.rs) | 17 | 未运行 |
| [conditional_format_sidecar_stream_tests.rs](../../crates/office-parser/src/conditional_format_sidecar_stream_tests.rs) | 15 | 未运行 |
| [conditional-format-office-source.contract.test.mjs](../../packages/office-studio/tests/conditional-format-office-source.contract.test.mjs) | 13 | 未运行 |
| [conditional-format-source-budget.contract.test.mjs](../../packages/office-studio/tests/conditional-format-source-budget.contract.test.mjs) | 11 | 未运行 |
| [conditional-format-workbook-budget.contract.test.mjs](../../packages/office-studio/tests/conditional-format-workbook-budget.contract.test.mjs) | 8 | 未运行 |
| 合计 | 112 | 仅声明数量 |

草稿覆盖：合法/无效私有投影、baseline 过滤与深克隆；重复 native writer→parser→writer 的 ID、稀疏/相同 priority、未知 JSON 与来源大小；外部标准公式修改优先于陈旧私有条件；rich dxf/theme；linked bar GUID/revision UID 复制；standalone x14 custom icons、NoIcons、严格 gte 与 inline dxf；未知条件计算隔离和原子编辑拒绝；文本 predicate/literal；高级 bar 字段；六种平均值编辑比较；未解析 dxf 在样式编辑后仍拒绝；额外标准公式以未知条件保留并参与受支持的原点投影；自动数据条在导入归一化后的样式命令中保留端点、正数/负数/混合/全零及区域外值隔离、普通 min/max 与自动边界的区别、错误用途与互换的原子拒绝，以及 Rust 本地投影的自动用途限制。上述均为预期与断言草稿，没有运行结果。

本轮草稿另外覆盖：超过 16 MiB 的 worksheet 中较晚 base/linked x14 的读取与来源编辑重开；巨大无关 styles/dxf 与引用原索引；无 CF 或无依赖时不读损坏公共资源；缺失/损坏的相关 dxf/theme、极端 dxfId、非法 sqref 与跨表原子失败；尾部损坏、Unicode/CDATA canonical 膨胀、继承 namespace/MC QName 重绑定；片段、part、整本资源与重复来源模型预算；最终 worksheet、styles part、单主题/整本资源的计量及原生错误 envelope。共同主题草稿覆盖 strict/transitional 颜色与 major/minor 字体、缺失颜色/字体/非法 tint、vendor 属性和 none scheme 不读主题；writer 原边界草稿同时断言不可用主题在替代样式注册前拒绝。TypeScript 使用 TextEncoder / JSON.stringify 小样本 oracle，覆盖 UTF-8、转义、孤立 surrogate、undefined/非有限数字、确定性 toJSON（函数属性、数组项、单数组根条目 key 和 `conditionalFormats` 属性上下文）、精确数组边界、两个小来源累计超限不修改输入、实际 UTF-8 超过 16 MiB 的 XML、originalId 字节限制、版本 1 与未知扩展、循环/BigInt/深度拒绝；整本草稿覆盖数组/map、每表转换 key、空表分隔符、单次计量的数组根转换与非数组拒绝、部分快照属性上下文克隆、跨表克隆前拒绝和模型/版本/选区/历史不发布。

theme fixture 是手写 XML，包含 12 个颜色、major/minor font 以及各 3 项格式样式；大型 styles、巨大 Text 与 Unicode XML 也是合成边界 fixture，均不是由 Office 程序生成并验收的真实样本。大型 CF 往返草稿只断言 CF 来源和编辑字段，不证明注入的大量单元格全部经过重写保存。此前格式刷、CF 命令、公式与浏览器声明数见原专项记录，本轮未重新统计或运行那些文件。

## 7. 未完成项与后续验收

| 缺口 | 当前事实及后续要求 |
| --- | --- |
| 大 part 来源提取 | 流式选择 CF 子树及相关资源的实现、独立 scanner guard 和合同草稿已落盘；整份 worksheet/styles 不再消耗 CF 片段预算。大 token 缓冲、核心解析器、包资源和真实大文件往返仍未测量/验收，不能把此项标为大文件完整支持。 |
| 继承上下文 | namespace、混合内容和 MC QName 重绑定已有草稿；`xml:base` 当前只保存最近继承值，多层相对 base 的组合及更多 Office/vendor 上下文语义未核实。 |
| 未知条件及引用协议 | 未知条件计算、公开公式编辑协议、vendor A1/结构化引用、多段/轴结构变换未完成。需真实资源事务与保存/重开合同，当前原点投影不足以替代。 |
| 高级视觉闭环 | bar 的边框、负边框、轴、方向、长度及 rich dxf 未映射样式的最终 Canvas 效果没有验收；来源保留和字段映射不等于 UI 可操作或渲染等价。 |
| 主题与包互操作 | 冲突主题当前拒绝；合并/重绑定、strict theme 的 transitional package 互操作和各种 Office 保留/改写扩展行为未验收。 |
| ID 与跨表 | 同表复制草稿已补；全局 native GUID 唯一性及跨表复制/结构变化仍需独立验收。 |
| 已知条件语义与公式范围 | 现有条件 Canvas 公式工作簿与完整跨表、外部、结构化引用及函数目录仍有差距。日期期间求值目前以 UTC 与 1899-12-30 为基点；1904、locale/week 与时区行为还需接入并验收。文本及错误条件的完整 Office 语义也需样本核实。 |
| 实际验证 | 没有编译、类型、语法、测试或浏览器结果。Excel/LibreOffice、真实 Office 文件、撤销/保存/重开、主题、焦点、IME、WebGPU/Canvas2D 与 Worker/WASM 均未验收。 |
| 性能与内存 | 没有设备、数据规模、耗时、帧时间、长任务、传输或内存测量。quick-xml 的巨大单个 token buffer、瞬时文本分配、各规则重复来源/主题、克隆及模型序列化成本仍需真实基准；大型条件计算、取消和跨 Worker/WASM 传输同样未验收，不声称优化完成。 |

本批静态读取过上述改动路径、正确的面板与阈值组件、固定 Univer 契约和公开 schema 依据，读取工作区状态并重数测试声明。功能清单仍保持 **未验收**；源码与草稿数量不代替实际行为及文件样本证据。

[U1]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/models/type.ts
[U2]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/models/conditional-formatting-rule-model.ts
[U3]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/schema.ts
[U4]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/menu/manage-rule.ts
[U5]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting/src/commands/commands/set-cf.command.ts
[U6]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/views/panel/RuleEdit.tsx
[U7]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/views/panel/RuleList.tsx
[U8]: https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/sheets-conditional-formatting-ui/src/views/panel/rule-edit/Rank.tsx
[O1]: https://raw.githubusercontent.com/dotnet/Open-XML-SDK/v3.3.0/data/schemas/schemas_openxmlformats_org_spreadsheetml_2006_main.json
[O2]: https://raw.githubusercontent.com/dotnet/Open-XML-SDK/v3.3.0/data/schemas/schemas_microsoft_com_office_spreadsheetml_2009_9_main.json
[O3]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.conditionalformattingrule
[O4]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformat
[O5]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.conditionalformattingrule
[O6]: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.databar
[O7]: https://github.com/jmcnamara/XlsxWriter/blob/RELEASE_3.2.9/xlsxwriter/worksheet.py
[O8]: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/df3812c6-077a-4ecd-853f-95e9c593ed38
[O9]: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/15094f93-f587-4f4e-be74-1fbdc93da768
[O10]: https://docs.rs/quick-xml/0.41.0/quick_xml/reader/struct.NsReader.html
