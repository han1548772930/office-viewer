# Univer 源码、页面布局与分层规则

审阅日期：2026-11-20  
参考仓库：dream-num/univer  
固定提交：`6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68`  
本地只读检出：`D:\rust_test\univer-reference`  
适用项目：本项目的 Vue / TypeScript / Pinia / Rust / WASM 实现。

本文件是项目执行标准的补充。下面把“上游代码事实”和“本项目执行要求”分开记录。功能范围以[完整功能对齐清单](UNIVER_FEATURE_PARITY.md)为准，覆盖 Sheets / Docs 的所有 Pro 功能。没有公开源码时，反推可观察行为并在本项目独立实现。

用户指定的渲染和计算路径为 Canvas / WebGPU 优先，以及大计算优先 Rust/WASM；具体约束以[渲染与计算标准](RENDERING_AND_COMPUTE_STANDARD.md)执行。此为本项目技术决策，不冒称上游 Univer 已使用 WebGPU。

## 1. 参考方法与版本控制

1. 阅读对应 feature 的菜单 schema、menu factory、命令实现、model/service、controller、view 和现有测试，形成从入口到模型的调用链；不能看一个组件就宣称完成源码研究。
2. 参考链接固定到上面的 commit；路径与符号是复查依据。升级参考版本时记录新增功能、菜单顺序、默认配置和数据结构差异，不能静默改变验收目标。
3. 文档站是动态内容；Pro 规格以审阅日期记录，新增规格补入清单，不能覆盖旧证据。以公开文档、类型/API、可访问演示和标准 Office 文件逐项反推输入输出、状态变化、错误、界面交互和重开保真。把已证实行为和待核实推断分开，不能据此宣称掌握闭源内部实现。
4. 借鉴本地参考仓库中有用的架构、数据结构、算法、界面、API 与测试。按 Vue / TypeScript / Rust/WASM 的边界改写；优化要维持公开行为，用同一数据和验收场景做前后对照。本地 reference 不安装依赖、不进生产依赖树。
5. 源码限制是当前上游行为的证据，不自动成为本项目的最终限制。用户要求的增强功能仍保留为待实现目标。

## 1.1 可访问工作台视觉基线

用户指定以下 Univer 公共工作台作为 Sheets / Docs 布局与交互视觉对照入口（核对日期：2026-11-20；页面内容可能随线上版本变化）：

| 产品模式 | 参考页面 |
| --- | --- |
| Sheets | [office.univer.ai/?filter=sheets](https://office.univer.ai/?filter=sheets) |
| Docs Modern | [office.univer.ai/?filter=docs-modern](https://office.univer.ai/?filter=docs-modern) |
| Docs Traditional | [office.univer.ai/?filter=docs-traditional](https://office.univer.ai/?filter=docs-traditional) |

需要视觉对照时，使用相同浏览器、视口尺寸、设备像素比、语言和产品模式分别记录参考页与本项目截图；记录截图时间、可见页面状态、视口和无法复现的动态内容。优先比较 Header、Ribbon 分类/组、产品 Header、左右侧栏、编辑画布、底栏与弹层边界，并逐个执行可观察的控件交互。线上演示只证明被观察到的页面行为，不代替源码调用链、模型/命令契约、Office 导出重开或性能基准；未知/登录后/动态内容单列待核实，不推断内部实现。

## 1.2 Exchange Office 互操作对照入口

用户指定 [Univer Exchange Import / Export Compatibility](https://office.univer.ai/exchange) 作为上传 Office 文件检查导入、内容、格式和布局，并在线编辑、下载后与原文件比较的公开演示入口（核对日期：2026-11-20）。页面说明的流程是上传 → 检查导入 → 编辑并下载 → 与原文件对照；这不是自动字段差异或像素差异报告。验收记录需注明格式、文件哈希、样本内容/特性、浏览器、可见诊断、编辑动作、导出文件与独立重开结果；只用合成样本或获准公开的样本，不上传用户/业务文档。演示接受的扩展名或格式不证明本项目已支持对应格式。

Exchange 仅用于观察公开转换行为和样本互操作对照，不是本项目运行依赖或服务端协议。项目仍用自身 `office-parser` / `office-writer` 链路；不依赖其后端、私有协议或内部实现。不能由页面宣称文件被服务端处理的措辞推断具体服务器实现、保留时长或隐私属性；上传前只使用无敏感合成数据。

## 2. 总工作台区域

**代码事实：** `DesktopWorkbenchContent` 通过 UI Part 装配顶部、内容、侧栏、底栏和浮层；区域注册及销毁由服务管理。[U04] [U05]

**项目要求：**

~~~text
应用壳：文件名、文件操作、Sheets / Docs 工作台切换（应用负责）
└─ EditorWorkbench
   ├─ Ribbon 分类标签与工具区
   ├─ 内容横向布局
   │  ├─ 左侧区域（按文档类型和上下文显示）
   │  ├─ 编辑器容器
   │  │  ├─ 产品 Header（Sheets 是公式栏）
   │  │  └─ 视口 / 排版画布 / 输入桥 / 选区与对象覆盖层
   │  └─ 右侧上下文面板
   └─ 产品 Footer
全局宿主：菜单、对话框、浮动工具、通知和弹出层
~~~

- App.vue 只装配应用服务和工作台；不能在 App、工作台、网格分别再放一套文件操作与撤销按钮。
- Header、Toolbar、Footer、ContextMenu 和左右面板可配置；显隐不能销毁正在编辑的模型、改变文件内容或清空历史。
- 编辑区独立滚动，父布局要有 `min-height: 0` / `min-width: 0` 约束；固定工具区不随画布滚走。
- 弹层必须使用统一宿主、主题和层级，考虑视口边缘、滚动、焦点归还及 Escape；不能被网格 overflow 裁断。
- 普通/加载/失败/只读状态来自真实任务与权限；就绪骨架不等同于内容成功加载。[U16]

## 3. Ribbon 和菜单注册

### 3.1 语义结构与布局变体

上游共享位置为 START、INSERT、FORMULAS、DATA、VIEW、OTHERS，最终可见分类由各插件的菜单注册与显隐状态共同决定。[U06] [U07] [U09]

- 本项目保持 Tab → Group → Item 的 schema。顺序以 `order` 等元数据确定，不能依赖模块偶然加载顺序。
- START 下保留历史/剪贴板、字体格式、段落或单元格布局、数字格式、其他工具等语义分组；INSERT、DATA 等由对应 feature 注入。
- 默认采用上游 desktop 的 classic 变体：分类行 + 紧凑工具行。按上游默认 16px 根字号换算，分类行 36px、工具行 40px；将尺寸提取为 token。
- grid 变体是独立布局：分类行 32px，工具区 73px，按 `gridLayout` 排两行、支持跨行/跨列与可见标签；不得把 grid 的双行配置强行套进 classic。[U10] [U11]
- simple/collapsed 是可选布局能力，不能用一个 simple 工具条替代完整功能范围。
- classic 根据可用宽度把溢出项放入更多菜单；grid 先收紧组密度，再按组收纳。溢出后必须仍可操作，不能只裁掉按钮或减小字体。[U10] [U11]
- 没有当前可用功能的空分组/空分类不显示；表格、图形、页眉页脚等上下文分类结束后恢复到有效分类。[U09]

### 3.2 菜单项与状态

每一项配置应至少包括稳定 ID、命令 ID、类型、标题/tooltip、图标、分组顺序、参数与状态来源。支持按钮、选择器、主按钮加下拉、纯子菜单四类语义。[U08]

- `hidden`、`disabled`、`activated`、`value` 分别表达可见性、可执行性、激活态和当前值，不能合并成一个布尔状态。
- 格式控件由当前单元格/文本范围/光标处的模型属性反推；混合选区有明确混合态，不能展示上一次点击的本地值。
- 一项多入口允许不同的菜单 ID，但必须指向同一命令。下拉项与主按钮可指定不同命令，参数由注册项传入。[U12]
- 权限、选区、活动文档、编辑中状态、剪贴板能力变化后，状态要同步更新；命令入口同样验证，不能只禁用按钮。
- 菜单禁用时不执行命令；提示说明具体原因。尚未实现的目标可在开发阶段标明状态，但不能以永久禁用验收。
- Tab 有键盘焦点和选中语义；方向键、Home/End 切换，RTL 方向相应调整。[U13]
- 工具栏操作保留当前文档范围，执行后按交互类型恢复编辑器焦点；颜色选择、字体选择不能让选区丢失。

## 4. Sheets 页面与交互

### 4.1 区域和尺寸

`SheetUIController` 分别注册产品 Header / Content / Footer，Header 提供公式栏。[U17] [U18]

~~~text
Ribbon
名称框 | 取消 | 确认 | fx | 公式编辑区 | 展开
行列头 + 网格 + 冻结分区 + 编辑器 + 选区/对象层 + 滚动条
新增 / 全部工作表 / 可滚动标签 | 选区统计 | 底部扩展菜单 | 缩放
~~~

- Footer 是一条 36px 的组合栏，工作表区域占剩余空间，统计、菜单和缩放按内容宽度排列。子区域独立可配置，全部隐藏才不显示该 Footer。[U18]
- 不在 sheet 标签下面重复增加全宽状态栏；解析器名称和技术调试信息不常驻挤占办公区域。
- 公式栏折叠 28px、展开 80px；名称框外宽 100px。这些是上游默认样式的参考 token，不使用随视口变化的字体来凑布局。[U20]

### 4.2 公式栏和工作表

- 公式栏和网格共用编辑会话与原始值；公式文本、计算值、显示值分开，取消恢复原值，确认成功才提交。[U20]
- 名称框支持范围跳转、命名范围查找/创建、Enter 提交、Escape 恢复；跳转同步选区和视口，不能仅显示 A1。[U21]
- 跨 sheet 公式引用不因公式栏失焦而清空引用选区；编辑器与网格坐标保持同步。
- SheetBar 顺序为新增 → 全部工作表 → 标签 → 溢出导航。导航按真实溢出和边界启用。[U22]
- 标签支持激活、键盘导航、拖拽排序、重命名、颜色、保护标记；空名、非法名和重名有明确校验。[U23]
- 区分可见、隐藏与 veryHidden；全部工作表菜单不能无条件列出不可发现项，恢复普通隐藏表走命令。[U24]
- 删除、复制、重排等影响公式、名称和对象引用的操作，要连同关联更新形成一次可撤销事务。

### 4.3 选区、编辑和右键

- 行头选择整行、列头选择整列、左上角全选；单元格、合并区域、多选、拖选和填充分别维护命中规则。[U30]
- 在已有选区内右键保留选区；在外部先确定目标。主区域、行头、列头、工作表标签使用各自菜单。[U19] [U29] [U30]
- 主区域包含特殊粘贴、清除类别、单元格移位与结构操作；行头包含行高/适应内容、插删/隐藏行；列头包含列宽/自动宽、插删/隐藏列；标签包含表级操作。不能所有位置复用同一份静态菜单。
- 双击/F2 延续原内容，普通键入启动覆盖编辑；Escape 取消，提交被拒绝时不得移动。Enter/Tab 与 Shift 方向分别处理。[U31]
- 输入法 composition、剪贴板、编辑器内快捷键与网格快捷键要有明确作用域；不能仅监听 window 键盘事件。
- 冻结按选区、当前视口和冻结分区计算；真实影响渲染、命中、滚动与打印，不能只画一条分隔线。
- 新数据或视图回传不重置有效活动表和选区；切换表保留各表视图状态。

### 4.4 状态与缩放

- 统计基于真实选区，重叠范围去重，按源码约定处理隐藏/过滤行与空值。[U25] [U26]
- Footer 展示当前选定的一项统计，向上菜单切换 SUM/COUNT/COUNTA/AVERAGE/MIN/MAX；无有效统计时不画固定 0。[U27]
- 上游数值统计需存在数值且非空数量超过一项，纯计数也有显示条件；最终实现写明确测试，而非随手累加字符串长度。[U25]
- 缩放是工作表视图状态；网格、行列头、浮动对象、编辑器与选区坐标使用同一变换，不能只 CSS 缩放文字。[U28]

### 4.5 功能注入

Sheets UI 包自身不等于完整产品。数字格式、公式、筛选、排序、条件格式、验证、表格、绘图、评论等来自各自插件；以 examples 的 preset 装配和各 feature 的 schema 交叉核查。[U42]

Pro 图表、透视表、切片器、打印等在本项目沿用同一模型、命令、菜单、面板、渲染扩展机制；不另做一套绕开历史和权限的临时页面。

## 5. Docs 页面与交互

### 5.1 文档壳与排版

`DocUIController` 注册 Footer、文档大纲、菜单、浮动工具和快捷键。[U33]

- 左侧导航采用按标题生成的文档大纲。上游默认在现代文档启用、传统文档关闭，并允许配置；本项目以可配置大纲对齐，不能误称上游所有文档默认都有大纲。标题使用稳定段落 ID，点击跳转，滚动同步高亮；无标题、标题删除/重命名/重排时同步更新。[U36]
- 大纲支持停靠/浮动，在空间不足或宽表发生遮挡时调整；不能挤压纸张直至不可编辑。上游浮动边界及最大宽度由 DocSideMenu / SideMenu 处理。[U37] [U38]
- 中央是文档布局视口：页尺寸、边距、节、页眉页脚、段落与文本运行影响排版；换行与分页来自布局，不由空行或固定文本长度伪造。
- 用稳定坐标连接文本索引、布局结果、渲染和选择；字体度量变化、表格增高或图片加载后重新布局受影响范围。[U46] [U47] [U48] [U49]
- 页面外画布留白、页面物理几何、正文页边距分开建模。上游渲染控制器的页面外间距不能当成 Office 页边距；传统纸张设置与现代正文宽度设置分别处理。[U50] [U51]
- 上游渲染和排版属于 engine-render，菜单/输入/交互属于 docs-ui。项目可独立实现渲染，但仍要分开文档模型、布局引擎、绘制与输入桥；整页 contenteditable 不能充当全部内核。
- 传统分页和现代连续文档是不同 flavor，交互、页眉页脚、分节等能力按模式判断；默认以用户 Office 文档的分页工作台验收，公开的现代模式能力另列场景，不能混用语义。

### 5.2 工具和上下文

- 开始：字体/字号/字符格式/命名样式、段落对齐、列表与页眉页脚；插入：表格、分隔线、分隔符、符号等；视图：页面设置、缩放、全屏。新增功能按对应语义组注入。[U34]
- schema 中 row/column/rowSpan 是 grid 变体数据；classic 仍使用紧凑单行工具区。
- 格式应用于实际文本范围或段落范围，并从选区反映当前值；点击工具栏不丢失光标。
- Docs 粘贴支持保留源格式、匹配目标格式、仅文本，以及最近一次粘贴后切换模式；纯文本处理表格、图片和链接的行为单独测试。
- 文本范围、对象范围、表格、页眉页脚和正文使用不同上下文。插表/分节等检查合法位置和模式，不能无选区直接改全篇。[U35]
- 上游开源插表在部分上下文受限，属于参考事实；增强表格仍是本项目交付目标，不能用此限制排除需求。
- 页眉页脚支持正文与 header/footer segment 切换、节关联、首页/奇偶页、距页边距离、页码/总页数/日期时间字段；仅在合法上下文显示专用分类，结束后回到有效正文分类。页码不是手写静态数字。[U53]
- 复制、剪切、粘贴和删除分别判断有效内容及 Copy/Edit 等权限；不能把“有选区”当作所有菜单的唯一开关。[U35] [U52]
- 段落、表格、图片、绘图、脚注等统一进入模型、命令、撤销和保存；表格不以 metadata.html 作为唯一数据源。
- Footer 展示字数/统计、排版任务进度及 CountBar 的产品控件；未真正保存时不能固定显示“已保存”。[U39]
- 文档布局 Worker 与应用装配分离，任务带模型版本，过期排版结果不能覆盖新输入。[U43]

## 6. 分层、Pinia 和 feature 边界

依据上游贡献规范与跨端说明，模型和业务逻辑独立于 UI。[U01] [U02]

依赖方向：

~~~text
common ← models ← services ← commands/mutations
                            ↑
                      controllers ← views
                            ↑
                         plugin 装配
UI stores 组合状态，不能倒置领域层依赖方向
~~~

- common 只放真正通用基础；models 不依赖 Vue、Pinia、DOM、Canvas 或文件下载。
- services 实现公式、布局、读写、权限与业务规则；通过接口访问模型和平台能力，不调用组件实例。
- mutations 仅依赖显式可序列化参数和模型，不直接读当前 DOM 选区；controllers 把 UI 上下文转换成命令参数。
- views 展示状态和派发意图；watcher 不承担隐式格式转换、文件写出或全量文档克隆。
- feature 内可按 common / models / services / commands / controllers / views / stores 组织，以 plugin.ts 注册，根 index.ts 输出稳定公共 API；不创建一层层转导出的空 barrel。
- Pinia 按 model/session/selection/view/task 分职责。大规模稀疏模型采用受控更新和切片订阅；DOM 节点、Canvas、WASM 实例、Worker、服务对象不放进可持久化文档状态。
- 序列化快照是显式过程，不能对 Vue Proxy 直接 structuredClone。数据回传有来源/版本标记，避免父子 watcher 循环和选区重置。
- 超过 400 行的 SFC 或同时持有解析/历史/渲染/工具栏职责的文件必须评估拆分；按职责提取 feature，不按随意行数切成互相隐式依赖的碎片。
- 新增功能只修改所属 feature 及必要注册入口，不能每次都改 App 和两个工作台的大段分支。

## 7. 命令、历史、生命周期

上游明确区分业务编排、非持久交互、持久模型变更。[U14]

| 类型 | 本项目职责 | 历史与同步 |
| --- | --- | --- |
| command | 校验权限/上下文，组织一次业务动作 | 组织成功后记录一组正向/逆向 mutation |
| mutation | 用显式参数修改持久化模型 | 可序列化、可重放；作为协作和版本基础 |
| operation | 选区、滚动、面板、临时编辑会话等 | 默认不进入内容历史；是否同步单独声明 |

- ID 采用 sheet/doc 等单数业务域及类型，例如 `sheet.command.set-cell-value`；API、菜单、快捷键共用命令。[U03]
- 撤销以文档/工作簿 ID 隔离，批量粘贴、合并、行列结构变更合为一次业务事务；每个组件另存一份内容历史属于架构缺陷。[U15]
- 持久化 savepoint 与 dirty 状态来自真实内容版本，撤销回到保存点可以清除 dirty；选区/缩放变化不标记内容修改。
- 命令失败不能部分写入模型后仍报成功，失败结果包含稳定错误码；异步动作校验文档 ID 与版本。
- 注册入口返回可释放句柄，卸载回收监听、菜单、快捷键、计时器、Worker 和渲染资源。
- 区分 Starting / Ready / Rendered / Steady 的意义或实现同等状态机：创建模型、首次可绘制、首屏完成、后台任务就绪不能混称“初始化完成”。[U16]
- Sheets / Docs 惰性加载，切换保留模型快照和各自状态，防止旧异步任务覆盖新工作台；卸载交互实例时释放资源。[U41]

## 8. Rust/WASM、保真与服务边界

**转换参考（2026-11-20）：** [Univer 的应用架构](https://docs.univer.ai/server/architecture#integrate-file-import-and-export)把 Office 文件与文档数据之间的转换交给 Exchange，可由服务端或本地 Node.js 进程调用。独立快照转换不需要协作；协作文件导出取已确认版本。公开入口和适用格式见 [Sheets 导入导出](https://docs.univer.ai/guides/sheets/features/import-export)及 [Docs 导入导出](https://docs.univer.ai/guides/docs/features/import-export)。

- 本项目复用已有 Office 转换能力：`crates/office-parser` 负责读取，`crates/office-writer` 负责写出，浏览器通过 Worker 调用；需要在服务端运行时扩展本项目后端。对齐公开 Facade 的输入、输出、格式和失败语义，从样本验证模型及资源保真，不接入 Univer Server 或 Exchange，也不依赖其私有协议。
- `crates/office-parser` 包含 façade/API、版本化 DTO、格式路由、duke-pages 文档适配、rxls 工作簿适配、诊断与测试；复杂适配继续按功能拆模块，不全部放 parser.rs。
- WASM 负责真实文件解析、规范化的解析 DTO 以及适合其环境的计算；不依赖 Vue，不操作 DOM，不管理工具栏与菜单。
- `schemaVersion/parserVersion/format/fileName/size/model/payload/diagnostics` 契约由 Rust 和 TypeScript 共同验证；payload.modelType/data 不能被前端旧字段假设替代。
- Worker 协议包含 requestId、documentId、modelVersion、任务类型、进度、取消/过期与错误；取消解析不等于仅隐藏 loading。
- 解析 DTO → 领域模型 → 布局模型 → 渲染分层。字段映射记录来源、默认值、保留字段和诊断，不能将真实段落替换成一句占位文案。
- 数据、公式、样式、节、关系和媒体保留与 UI 是否已实现分开；无法解释的 OOXML part 要有保留/诊断策略，不静默丢弃。
- duke-pages/rxls 的读取能力不意味着具备写出能力。XLSX/DOCX writer 是单独的服务/模块；纯文本、HTML、JSON 或改扩展名不满足 Office 导出。
- 多人协作、跨设备版本和持久权限需要独立服务端。必要服务与协议留在本项目；本地解析/编辑不依赖服务端才能运行。
- 浏览器能力、上游 parser 能力或模块尚未实现，只能形成开发缺口，不改变完整功能目标。
- 公式与透视等计算各自只有一个权威引擎。Rust 和 TypeScript 的缓存可以服务渲染，但不能独立决定互相冲突的结果；异步计算结果带输入版本和应用版本。
- 公式解析、依赖图、增量重算和结构变更后的引用调整属于计算/业务引擎，读取公式文本及缓存值不能算实现公式功能。[U54] [U55]
- 筛选、排序、验证和条件格式拥有独立模型与服务。插删/移动/粘贴应通知各资源的范围变换处理器；不可只改变 DOM 顺序或永久写入条件底色。[U56] [U57] [U58] [U59]
- 数字格式由原始值、locale 和日期系统决定，编辑后重新格式化；不能用 String(value) 覆盖全部格式。[U60]
- “仓库有插件”“示例注册插件”“类型/Facade 可导入”“已连通实际服务”是不同状态。Docs 查找替换的范围、本地评论 datasource 的成功返回和 preset 的协作配置均需单独核实。[U61] [U62] [U63] [U64]

## 9. 实施和验收

1. 先在功能清单中选定 ID，注明参考文件/公开规格及具体交互。
2. 从模型、参数、命令、服务、状态到视图连通，确保各入口一致。
3. 对变更执行相应单元、契约或浏览器行为测试；涉及 Office 互操作时保留真实文件样本和独立重开结果。
4. 用布局、事件和交互测试验证相同视口、设备缩放、编辑器缩放、字体和 Ribbon 变体下的行为。只有用户明确要求视觉回归时才加截图或像素检查。
5. 性能测试记录设备、数据规模、渲染节点、耗时与内存；禁止仅靠小型样例宣称大型文件可用。
6. 当前静态草稿、解析支持、构建成功与完整功能验收分开记录；未通过的项保持未验收。
7. 不以“Pro 没源码”“解析库不支持”结束任务；先登记缺口，明确独立实现模块和下一步验证，继续推进有依据的实现。

## 10. 固定源码索引

下列文件已在固定提交中检查；名称与路径用于后续开发者复查，链接不是浮动 dev 分支。

| 编号 | 源码文件 | 审阅重点 |
| --- | --- | --- |
| U01 | [CONTRIBUTING.md][U01] | 目录组织与依赖方向 |
| U02 | [docs/ISOMORPHIC.md][U02] | 模型/逻辑与浏览器 UI 分离 |
| U03 | [docs/NAMING_CONVENTION.md][U03] | 命令与文件命名 |
| U04 | [packages/ui/src/views/workbench/Workbench.tsx][U04] | DesktopWorkbenchContent |
| U05 | [packages/ui/src/services/parts/parts.service.ts][U05] | BuiltInUIPart / UIPartsService |
| U06 | [packages/ui/src/services/menu/types.ts][U06] | RibbonPosition / ContextMenuPosition |
| U07 | [packages/ui/src/services/menu/menu-manager.service.ts][U07] | MenuManagerService / IRibbonGridLayout |
| U08 | [packages/ui/src/services/menu/menu.ts][U08] | 菜单类型及独立的 hidden/disabled/value/activated |
| U09 | [packages/ui/src/services/ribbon/ribbon.service.ts][U09] | DesktopRibbonService |
| U10 | [packages/ui/src/views/components/ribbon/Ribbon.tsx][U10] | Ribbon 布局变体及溢出 |
| U11 | [packages/ui/src/views/components/ribbon/RibbonGrid.tsx][U11] | 两行 grid 与整组溢出 |
| U12 | [packages/ui/src/views/components/ribbon/ToolbarItem.tsx][U12] | 菜单控件到统一命令的调用 |
| U13 | [packages/ui/src/views/components/ribbon/ribbon-menu/use-ribbon-tabs.ts][U13] | 方向键、Home/End 与焦点 |
| U14 | [packages/core/src/services/command/command.service.ts][U14] | CommandType / ICommand / IMutation / IOperation |
| U15 | [packages/core/src/services/undoredo/undoredo.service.ts][U15] | LocalUndoRedoService / IUndoRedoItem |
| U16 | [packages/core/src/services/lifecycle/lifecycle.ts][U16] | LifecycleStages |
| U17 | [packages/sheets-ui/src/controllers/ui.controller.ts][U17] | 工作台区域与菜单注册 |
| U18 | [packages/sheets-ui/src/views/sheet-container/SheetContainer.tsx][U18] | 表格 header/content/footer |
| U19 | [packages/sheets-ui/src/menu/schema.ts][U19] | 表格菜单结构、排序及 gridLayout |
| U20 | [packages/sheets-ui/src/views/formula-bar/FormulaBar.tsx][U20] | 公式栏与编辑桥 |
| U21 | [packages/sheets-ui/src/views/defined-name/DefinedName.tsx][U21] | 名称/范围导航 |
| U22 | [packages/sheets-ui/src/views/sheet-bar/SheetBar.tsx][U22] | 新增/目录/标签/滚动 |
| U23 | [packages/sheets-ui/src/views/sheet-bar/sheet-bar-tabs/SheetBarTabs.tsx][U23] | 标签、重命名和重排 |
| U24 | [packages/sheets-ui/src/views/sheet-bar/sheet-bar-menu/SheetBarMenu.tsx][U24] | 隐藏工作表语义 |
| U25 | [packages/sheets-ui/src/services/status-bar.service.ts][U25] | 统计项及显示条件 |
| U26 | [packages/sheets-ui/src/controllers/status-bar.controller.ts][U26] | 统计去重、过滤与更新 |
| U27 | [packages/sheets-ui/src/views/status-bar/StatusBar.tsx][U27] | 统计值与向上菜单 |
| U28 | [packages/sheets-ui/src/views/sheet-slider/ZoomSlider.tsx][U28] | 工作表缩放状态 |
| U29 | [packages/sheets-ui/src/controllers/render-controllers/contextmenu.render-controller.ts][U29] | 右键目标路由 |
| U30 | [packages/sheets-ui/src/services/selection/selection-render.service.ts][U30] | 区域选择与命中 |
| U31 | [packages/sheets-ui/src/controllers/editor/editing.render-controller.ts][U31] | 编辑开始、提交、取消及移动 |
| U32 | [packages/sheets-ui/src/menu/menu.ts][U32] | 选区与权限驱动菜单状态 |
| U33 | [packages/docs-ui/src/controllers/ui.controller.ts][U33] | 文档菜单、快捷键和 UI Part 注册 |
| U34 | [packages/docs-ui/src/menu/schema.ts][U34] | 文档工具栏与上下文菜单 |
| U35 | [packages/docs-ui/src/menu/menu.ts][U35] | 文档范围、flavor 与权限状态 |
| U36 | [packages/docs-ui/src/views/DocSideMenu.tsx][U36] | 标题大纲与滚动同步 |
| U37 | [packages/docs-ui/src/views/SideMenu.tsx][U37] | 大纲浮动与宽度 |
| U38 | [packages/docs-ui/src/views/doc-side-menu-layout.ts][U38] | 大纲和宽表的布局冲突判断 |
| U39 | [packages/docs-ui/src/views/doc-footer/DocFooter.tsx][U39] | 字数、布局进度和缩放/导航 |
| U40 | [packages/engine-render/src/components/docs/document.ts][U40] | 文档渲染边界 |
| U41 | [examples/src/main.ts][U41] | 产品惰性切换与设置生命周期 |
| U42 | [examples/src/sheets/mount.ts][U42] | Sheets 实际 preset 装配 |
| U43 | [examples/src/docs/mount.ts][U43] | Docs 实际 preset 与布局 Worker 装配 |
| U44 | [presets/src/preset.ts][U44] | preset 展开及依赖装配 |
| U45 | [README.md][U45] | 开源仓库与单独分发扩展的边界 |

[U01]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/CONTRIBUTING.md
[U02]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/docs/ISOMORPHIC.md
[U03]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/docs/NAMING_CONVENTION.md
[U04]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/views/workbench/Workbench.tsx
[U05]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/services/parts/parts.service.ts
[U06]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/services/menu/types.ts
[U07]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/services/menu/menu-manager.service.ts
[U08]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/services/menu/menu.ts
[U09]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/services/ribbon/ribbon.service.ts
[U10]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/views/components/ribbon/Ribbon.tsx
[U11]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/views/components/ribbon/RibbonGrid.tsx
[U12]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/views/components/ribbon/ToolbarItem.tsx
[U13]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/ui/src/views/components/ribbon/ribbon-menu/use-ribbon-tabs.ts
[U14]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/core/src/services/command/command.service.ts
[U15]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/core/src/services/undoredo/undoredo.service.ts
[U16]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/core/src/services/lifecycle/lifecycle.ts
[U17]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/controllers/ui.controller.ts
[U18]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/sheet-container/SheetContainer.tsx
[U19]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/menu/schema.ts
[U20]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/formula-bar/FormulaBar.tsx
[U21]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/defined-name/DefinedName.tsx
[U22]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/sheet-bar/SheetBar.tsx
[U23]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/sheet-bar/sheet-bar-tabs/SheetBarTabs.tsx
[U24]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/sheet-bar/sheet-bar-menu/SheetBarMenu.tsx
[U25]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/services/status-bar.service.ts
[U26]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/controllers/status-bar.controller.ts
[U27]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/status-bar/StatusBar.tsx
[U28]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/views/sheet-slider/ZoomSlider.tsx
[U29]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/controllers/render-controllers/contextmenu.render-controller.ts
[U30]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/services/selection/selection-render.service.ts
[U31]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/controllers/editor/editing.render-controller.ts
[U32]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-ui/src/menu/menu.ts
[U33]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/controllers/ui.controller.ts
[U34]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/menu/schema.ts
[U35]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/menu/menu.ts
[U36]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/views/DocSideMenu.tsx
[U37]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/views/SideMenu.tsx
[U38]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/views/doc-side-menu-layout.ts
[U39]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/views/doc-footer/DocFooter.tsx
[U40]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/engine-render/src/components/docs/document.ts
[U41]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/examples/src/main.ts
[U42]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/examples/src/sheets/mount.ts
[U43]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/examples/src/docs/mount.ts
[U44]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/presets/src/preset.ts
[U45]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/README.md

## 11. 引擎与扩展的补充源码索引

以下路径由并行审阅任务核对，用于补齐 UI 调用链背后的模型/引擎与服务事实。

| 编号 | 源码文件 | 审阅重点 |
| --- | --- | --- |
| U46 | [packages/core/src/docs/data-model/document-data-model.ts][U46] | DocumentDataModel 与正文/页眉页脚结构 |
| U47 | [packages/docs/src/services/doc-skeleton-manager.service.ts][U47] | 模型到布局发布 |
| U48 | [packages/engine-render/src/components/docs/layout/doc-skeleton.ts][U48] | DocumentSkeleton 与文本位置/命中 |
| U49 | [packages/docs-ui/src/services/selection/doc-selection-render.service.ts][U49] | 逻辑选择和 DOM 输入桥 |
| U50 | [packages/docs-ui/src/services/doc-page-layout.service.ts][U50] | 页位置与视口几何 |
| U51 | [packages/docs-ui/src/views/PageSettings.tsx][U51] | 现代/传统页面设置 |
| U52 | [packages/docs-ui/src/menu/context-menu.ts][U52] | 复制/剪切/粘贴/表格上下文权限 |
| U53 | [packages/docs-ui/src/menu/header-footer-ribbon.ts][U53] | 页眉页脚上下文分类 |
| U54 | [packages/engine-formula/src/plugin.ts][U54] | 公式服务、语法/求值与依赖注册 |
| U55 | [packages/sheets-formula/src/plugin.ts][U55] | 表格公式引用与结构更新 |
| U56 | [packages/sheets-filter/src/plugin.ts][U56] | 筛选逻辑插件 |
| U57 | [packages/sheets-sort/src/plugin.ts][U57] | 排序逻辑插件 |
| U58 | [packages/sheets-data-validation/src/plugin.ts][U58] | 数据验证规则与服务 |
| U59 | [packages/sheets-conditional-formatting/src/plugin.ts][U59] | 条件格式规则和计算 |
| U60 | [packages/sheets-numfmt/src/plugin.ts][U60] | 数字格式命令与业务 |
| U61 | [packages/docs-find-replace/README.md][U61] | 仓库存在与示例启用、查找范围的区别 |
| U62 | [packages/thread-comment/src/services/tc-datasource.service.ts][U62] | 本地 datasource 与持久评论的边界 |
| U63 | [presets/packages/preset-sheets-core/src/preset.ts][U63] | 核心模型、引擎和 Worker 配置 |
| U64 | [docs/API_STABILITY.md][U64] | 类型、Facade 导入和插件注册的区别 |

[U46]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/core/src/docs/data-model/document-data-model.ts
[U47]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs/src/services/doc-skeleton-manager.service.ts
[U48]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/engine-render/src/components/docs/layout/doc-skeleton.ts
[U49]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/services/selection/doc-selection-render.service.ts
[U50]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/services/doc-page-layout.service.ts
[U51]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/views/PageSettings.tsx
[U52]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/menu/context-menu.ts
[U53]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-ui/src/menu/header-footer-ribbon.ts
[U54]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/engine-formula/src/plugin.ts
[U55]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-formula/src/plugin.ts
[U56]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-filter/src/plugin.ts
[U57]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-sort/src/plugin.ts
[U58]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-data-validation/src/plugin.ts
[U59]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-conditional-formatting/src/plugin.ts
[U60]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/sheets-numfmt/src/plugin.ts
[U61]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/docs-find-replace/README.md
[U62]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/packages/thread-comment/src/services/tc-datasource.service.ts
[U63]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/presets/packages/preset-sheets-core/src/preset.ts
[U64]: https://github.com/dream-num/univer/blob/6a0969cc3f1c4ad078a0f80eda0c5e61ec7b2c68/docs/API_STABILITY.md
