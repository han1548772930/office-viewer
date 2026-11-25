# DOC-15 LaTeX 模板与上下文 Ribbon

记录日期：2026-11-26。状态：**前端实现与测试草稿已落盘，未验收**。本批续接 [公式编辑、公开范围与资源](DOC-15-latex-editing.md)，完成固定目录及其源码插入路径，不代表完整 TeX、Office 数学互操作或 DOC-15 全量完成。

按用户先实现、后统一测试的要求，未执行测试、类型检查、构建、dev、浏览器、Cargo/Rust/WASM 或基准。以下核对是读取源文件、声明计数和文本比较；没有加载或执行商业包及本项目 parser/layout。

## 1. 固定依据与逐入口清点

OSS 为 `E:\rust_test\univer` 的 v1.6.0，固定提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。Pro 依据为本项目缓存的 v1.6.0 发布 ES 和类型；不安装商业 SDK，不把其数学内核视为本项目已有实现。

| 来源 | 静态核实的契约 |
| --- | --- |
| [docs-latex-ui ES](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/es/index.js) 的 `H / mt / V / U / vt / yt / bt / St / Ct / wt` | 71 个模板、11 类及顺序、源字符串、预览字符串和首占位；category menu ID 为 `docs-latex-ui.ribbon.menu.${category}`，gallery component ID 为 `docs-latex-ui.component.formula-gallery.${category}` |
| [模板类型](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/common/formula-templates.d.ts)、[模板插入类型](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/common/insert-formula-template.d.ts) | 模板字段及 source-range planner；选中源码替换首占位后光标停在整段插入内容后，空选区选中首占位 |
| [Ribbon 常量](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/const/ribbon.d.ts)、[上下文控制器](../../.research.local/univer-1.6.0/packages/docs-latex-ui/lib/types/controllers/docs-latex-contextual-ribbon.controller.d.ts) | formula popup 打开时显示并激活上下文 tab，关闭时隐藏；不是公式 atom 被选中就永久显示 |
| ES `Nt / Ft / z.requestTemplateInsert / K` | 操作 `docs-latex.operation.insert-formula-template` 仅改当前 popup draft；无打开 popup 或坏 ID 返回 false；上下文生命周期与 draft owner 一致 |
| ES `W / xt / Tt / Et / Dt` | 常驻 Insert/Media order 8、grid row 2 / column 7；浮动文字样式 order 13；段落 Layout order 8，插入 quickBottom order 4、insert order 3，编辑 quickBottom order 5 |
| [公开公式规格](https://docs.univer.ai/zh-CN/guides/docs/features/latex) | 完整产品目标；线上演示、真实 Office 数学文件与最终效果尚待统一验证 |

本批静态逐项比较本地目录与 ES 中 71 次 `H(category,id,String.raw...)`：**类别、ID、源文本、预览文本、首占位 UTF-16 起止均一致，差异 0**。比较只解读字符串字面量与占位符，没有运行 ES、TeX 或本项目算法。下表展开具体 source，供后续复查。

本地入口口径：

- 常规 Ribbon 新增 **0** 根。已有 `doc-latex` 仍在 Docs 52（Home 33 / Insert 16 / View 3）的 Insert/Media 中，不能重复增加常驻计数。
- 新增独立上下文 tab **1**、group **1**、selector 根 **11**，目录选择项 **71**。classic / grid 是同一集合的两种呈现，不相加。
- 弹窗内增设类别 Select **1** 和同目录 gallery；它复用 71 模板，不再计作 71 个 Ribbon 根或新的模板集合。
- 既有公式右键 3 项、双击路径保持前批口径；本批没有新增段落 Layout、quickBottom、insert、移动和浮动文字样式挂载。常规溢出走共享 Ribbon 宿主，尚未实际验收。

上下文 tab ID 为 `docs-latex-ui.ribbon.formula`，order 100，contextual=true；group ID 为 `docs-latex-ui.ribbon.group.formula`，order 0。classic 依下表 order 排序；grid 的 row / column / rowSpan 如下，每项 showLabel=true。

| classic order | 类别 | selector ID | 模板子项数 | grid row / column / rowSpan |
| ---: | --- | --- | ---: | --- |
| 0 | `fractions` | `docs-latex-ui.ribbon.menu.fractions` | 4 | 1 / 1 / 2 |
| 1 | `scripts` | `docs-latex-ui.ribbon.menu.scripts` | 4 | 1 / 2 / 2 |
| 2 | `radicals` | `docs-latex-ui.ribbon.menu.radicals` | 3 | 1 / 3 / 2 |
| 3 | `integrals` | `docs-latex-ui.ribbon.menu.integrals` | 6 | 1 / 4 / 2 |
| 4 | `large-operators` | `docs-latex-ui.ribbon.menu.large-operators` | 6 | 1 / 5 / 2 |
| 5 | `brackets` | `docs-latex-ui.ribbon.menu.brackets` | 8 | 1 / 6 / 2 |
| 6 | `functions` | `docs-latex-ui.ribbon.menu.functions` | 8 | 1 / 7 / 2 |
| 7 | `accents` | `docs-latex-ui.ribbon.menu.accents` | 8 | 1 / 8 / 2 |
| 8 | `limits-logarithms` | `docs-latex-ui.ribbon.menu.limits-logarithms` | 6 | 1 / 9 / 2 |
| 9 | `operators` | `docs-latex-ui.ribbon.menu.operators` | 12 | 1 / 10 / 2 |
| 10 | `matrices` | `docs-latex-ui.ribbon.menu.matrices` | 6 | 1 / 11 / 2 |

上下文根显隐由编辑器 open 状态决定；disabled 同步捕获目标是否有效、当前 edit 权限/锁与 IME composing。实际操作入口再次检查，不依赖按钮禁用作为唯一防线。关闭时恢复打开前的有效 tab；共享 Ribbon 处理已失效分类回退。浏览器生命周期和窄屏溢出未验收。

## 2. 全部 71 个源模板

首占位使用 **0 起始 UTF-16、end exclusive**。source 包含的空格、反斜线、`&` 与矩阵换行命令均保留；本地语义标签可本地化，source 不翻译。所有预览 source 与该表 source 相同。

| 模板 ID | 类别 | LaTeX source | 首占位源码 | [start, end) |
| --- | --- | --- | --- | --- |
| `fraction-basic` | `fractions` | <code>\frac{a}{b}</code> | <code>a</code> | [6, 7) |
| `fraction-expression` | `fractions` | <code>\frac{a+b}{c+d}</code> | <code>a+b</code> | [6, 9) |
| `fraction-derivative` | `fractions` | <code>\frac{\mathrm{d}y}{\mathrm{d}x}</code> | <code>y</code> | [16, 17) |
| `fraction-partial` | `fractions` | <code>\frac{\partial f}{\partial x}</code> | <code>f</code> | [15, 16) |
| `script-superscript` | `scripts` | <code>a^{b}</code> | <code>a</code> | [0, 1) |
| `script-subscript` | `scripts` | <code>a_{b}</code> | <code>a</code> | [0, 1) |
| `script-combined` | `scripts` | <code>a_{b}^{c}</code> | <code>a</code> | [0, 1) |
| `script-left-right` | `scripts` | <code>{}_{a}^{b}c_{d}^{e}</code> | <code>c</code> | [10, 11) |
| `radical-square` | `radicals` | <code>\sqrt{a}</code> | <code>a</code> | [6, 7) |
| `radical-nth` | `radicals` | <code>\sqrt[n]{a}</code> | <code>a</code> | [9, 10) |
| `radical-nested` | `radicals` | <code>\sqrt{\sqrt{a}}</code> | <code>a</code> | [12, 13) |
| `integral-basic` | `integrals` | <code>\int f(x)\,\mathrm{d}x</code> | <code>f(x)</code> | [5, 9) |
| `integral-bounded` | `integrals` | <code>\int_{a}^{b} f(x)\,\mathrm{d}x</code> | <code>f(x)</code> | [13, 17) |
| `integral-double` | `integrals` | <code>\iint f(x,y)\,\mathrm{d}x\,\mathrm{d}y</code> | <code>f(x,y)</code> | [6, 12) |
| `integral-triple` | `integrals` | <code>\iiint f(x,y,z)\,\mathrm{d}V</code> | <code>f(x,y,z)</code> | [7, 15) |
| `integral-contour` | `integrals` | <code>\oint f(z)\,\mathrm{d}z</code> | <code>f(z)</code> | [6, 10) |
| `integral-contour-bounded` | `integrals` | <code>\oint_{C} f(z)\,\mathrm{d}z</code> | <code>f(z)</code> | [10, 14) |
| `large-sum` | `large-operators` | <code>\sum_{i=1}^{n} a_i</code> | <code>a_i</code> | [15, 18) |
| `large-product` | `large-operators` | <code>\prod_{i=1}^{n} a_i</code> | <code>a_i</code> | [16, 19) |
| `large-coproduct` | `large-operators` | <code>\coprod_{i=1}^{n} A_i</code> | <code>A_i</code> | [18, 21) |
| `large-union` | `large-operators` | <code>\bigcup_{i=1}^{n} A_i</code> | <code>A_i</code> | [18, 21) |
| `large-intersection` | `large-operators` | <code>\bigcap_{i=1}^{n} A_i</code> | <code>A_i</code> | [18, 21) |
| `large-logical-product` | `large-operators` | <code>\bigwedge_{i=1}^{n} P_i</code> | <code>P_i</code> | [20, 23) |
| `bracket-parentheses` | `brackets` | <code>\left(a\right)</code> | <code>a</code> | [6, 7) |
| `bracket-square` | `brackets` | <code>\left[a\right]</code> | <code>a</code> | [6, 7) |
| `bracket-braces` | `brackets` | <code>\left\{a\right\}</code> | <code>a</code> | [7, 8) |
| `bracket-angle` | `brackets` | <code>\left\langle a\right\rangle</code> | <code>a</code> | [13, 14) |
| `bracket-absolute` | `brackets` | <code>\left&#124;a\right&#124;</code> | <code>a</code> | [6, 7) |
| `bracket-norm` | `brackets` | <code>\left\&#124;a\right\&#124;</code> | <code>a</code> | [7, 8) |
| `bracket-floor` | `brackets` | <code>\left\lfloor a\right\rfloor</code> | <code>a</code> | [13, 14) |
| `bracket-ceiling` | `brackets` | <code>\left\lceil a\right\rceil</code> | <code>a</code> | [12, 13) |
| `function-sine` | `functions` | <code>\sin\left(a\right)</code> | <code>a</code> | [10, 11) |
| `function-cosine` | `functions` | <code>\cos\left(a\right)</code> | <code>a</code> | [10, 11) |
| `function-tangent` | `functions` | <code>\tan\left(a\right)</code> | <code>a</code> | [10, 11) |
| `function-logarithm` | `functions` | <code>\log\left(a\right)</code> | <code>a</code> | [10, 11) |
| `function-natural-logarithm` | `functions` | <code>\ln\left(a\right)</code> | <code>a</code> | [9, 10) |
| `function-exponential` | `functions` | <code>\exp\left(a\right)</code> | <code>a</code> | [10, 11) |
| `function-minimum` | `functions` | <code>\min\left(a,b\right)</code> | <code>a</code> | [10, 11) |
| `function-maximum` | `functions` | <code>\max\left(a,b\right)</code> | <code>a</code> | [10, 11) |
| `accent-hat` | `accents` | <code>\hat{a}</code> | <code>a</code> | [5, 6) |
| `accent-bar` | `accents` | <code>\bar{a}</code> | <code>a</code> | [5, 6) |
| `accent-vector` | `accents` | <code>\vec{a}</code> | <code>a</code> | [5, 6) |
| `accent-dot` | `accents` | <code>\dot{a}</code> | <code>a</code> | [5, 6) |
| `accent-double-dot` | `accents` | <code>\ddot{a}</code> | <code>a</code> | [6, 7) |
| `accent-overline` | `accents` | <code>\overline{a+b}</code> | <code>a+b</code> | [10, 13) |
| `accent-underline` | `accents` | <code>\underline{a+b}</code> | <code>a+b</code> | [11, 14) |
| `accent-overbrace` | `accents` | <code>\overbrace{a+b+\cdots+z}^{n}</code> | <code>a+b+\cdots+z</code> | [11, 23) |
| `limit-basic` | `limits-logarithms` | <code>\lim_{x \to a} f(x)</code> | <code>f(x)</code> | [15, 19) |
| `limit-one-sided` | `limits-logarithms` | <code>\lim_{x \to a^{+}} f(x)</code> | <code>f(x)</code> | [19, 23) |
| `limit-infinity` | `limits-logarithms` | <code>\lim_{x \to \infty} f(x)</code> | <code>f(x)</code> | [20, 24) |
| `logarithm-base` | `limits-logarithms` | <code>\log_{a}\left(b\right)</code> | <code>b</code> | [14, 15) |
| `minimum-conditioned` | `limits-logarithms` | <code>\min_{x \in S} f(x)</code> | <code>f(x)</code> | [15, 19) |
| `maximum-conditioned` | `limits-logarithms` | <code>\max_{x \in S} f(x)</code> | <code>f(x)</code> | [15, 19) |
| `operator-plus-minus` | `operators` | <code>a \pm b</code> | <code>a</code> | [0, 1) |
| `operator-multiply` | `operators` | <code>a \times b</code> | <code>a</code> | [0, 1) |
| `operator-divide` | `operators` | <code>a \div b</code> | <code>a</code> | [0, 1) |
| `operator-not-equal` | `operators` | <code>a \ne b</code> | <code>a</code> | [0, 1) |
| `operator-approximately` | `operators` | <code>a \approx b</code> | <code>a</code> | [0, 1) |
| `operator-less-equal` | `operators` | <code>a \le b</code> | <code>a</code> | [0, 1) |
| `operator-greater-equal` | `operators` | <code>a \ge b</code> | <code>a</code> | [0, 1) |
| `operator-element` | `operators` | <code>a \in A</code> | <code>a</code> | [0, 1) |
| `operator-not-element` | `operators` | <code>a \notin A</code> | <code>a</code> | [0, 1) |
| `operator-subset` | `operators` | <code>A \subseteq B</code> | <code>A</code> | [0, 1) |
| `operator-union` | `operators` | <code>A \cup B</code> | <code>A</code> | [0, 1) |
| `operator-intersection` | `operators` | <code>A \cap B</code> | <code>A</code> | [0, 1) |
| `matrix-2x2` | `matrices` | <code>\begin{matrix} a &amp; b \\ c &amp; d \end{matrix}</code> | <code>a</code> | [15, 16) |
| `matrix-3x3` | `matrices` | <code>\begin{matrix} a &amp; b &amp; c \\ d &amp; e &amp; f \\ g &amp; h &amp; i \end{matrix}</code> | <code>a</code> | [15, 16) |
| `matrix-parenthesized` | `matrices` | <code>\begin{pmatrix} a &amp; b \\ c &amp; d \end{pmatrix}</code> | <code>a</code> | [16, 17) |
| `matrix-bracketed` | `matrices` | <code>\begin{bmatrix} a &amp; b \\ c &amp; d \end{bmatrix}</code> | <code>a</code> | [16, 17) |
| `matrix-braced` | `matrices` | <code>\begin{Bmatrix} a &amp; b \\ c &amp; d \end{Bmatrix}</code> | <code>a</code> | [16, 17) |
| `matrix-determinant` | `matrices` | <code>\begin{vmatrix} a &amp; b \\ c &amp; d \end{vmatrix}</code> | <code>a</code> | [16, 17) |

## 3. 数据、命令、历史和保存边界

领域权威数据保持前批公式 atom 和原 source；模板目录与 draft owner 不进入文档快照。顶部类别 gallery、弹窗内 gallery 和公开调用都使用同一 source planner。

| 入口 / 接口 | 行为与边界 |
| --- | --- |
| `docs-latex.operation.insert-formula-template`、本地 alias `doc.operation.insert-latex-template` | 进入已注册 API operation，使用 unit 与可选 revision/history guards，再请求当前 draft sink；不产生内容历史 |
| 本地扩展 `FDocument.bindLatexTemplateEditor(insert)` | 按 OfficeStudio/unit 绑定运行时 source editor，返回 `dispose()/isCurrent()`；新 owner 替换旧 owner，旧 handle 的 dispose 不删除新 owner |
| 本地扩展 `FDocument.insertLatexTemplate(templateId)` | 调用同一 public operation；无 editor / 无效模板返回 false；UI sink 对 stale target、host、权限和 composing 返回 false |
| source planner `insertDocumentLatexTemplate` | 遵循上游空选区/已有选区行为；有限整数范围按上游 clamp；非整数和切开 surrogate pair 的范围明确拒绝 |
| Ant source draft | 保留选区并插入 source；nextTick 恢复占位或光标前确认 open、readonly、composing、draft generation 与 source 内容 |
| 应用公式 | 继续使用 `doc.command.insert-latex-formula` 或 `doc.command.update-latex-formula`；其余删除/转源文本沿用前批 canonical 命令 |
| 关闭 unit / studio | 先清理对应 owner；重建相同 unit ID 不复用已关闭 editor。关闭弹窗、API 更换和组件卸载通过 watcher cleanup dispose |
| nested/story/segment 保存 | 沿用已有捕获 target、host 和 canonical apply，保持稳定公式 ID、CUSTOM5、named `DOC_LATEX_PLUGIN` 及公开 source offsets；新增草稿覆盖表格、header 和 note 重开，尚未运行 |

template operation 的作用域是当前 unit 的临时 editor；header/note/表格的内容写入位置由该 editor 原先捕获的真实 host/selection 决定，再经过 canonical 内容事务。未新增第二份内容历史或另一套 segment 保存协议。

运行时句柄、原生 textarea ref、Canvas 和组件实例均不进入模型或 Pinia 的持久状态。公开 inline 与本地 block 扩展的区别沿用前批记录，不能把本地 block 视为已确认的 Pro 公开 displayMode。

## 4. 算法、绘制与平台边界

数学 parser / typed AST / layout 已从 `latex-formula.ts` 拆到纯模块，旧导出路径保持。目录所需分支增加：

| 类型 / 命令 | 本地实现与仍待验收的限制 |
| --- | --- |
| 分式、左右上下标、根式 | 保留分子/分母、空基项左右 scripts 和 indexed sqrt；nested 共享 baseline |
| 积分及大型运算符 | `int/iint/iiint/oint` 及 `sum/prod/coprod/bigcup/bigcap/bigwedge`；积分 bounds 放侧面，large limits 置上下 |
| 函数、极限与对数 | roman 函数名、`lim/min/max` 条件范围及原 source；不实现数值求值 |
| 可伸展定界符 | `left/right`、绝对值/范数/尖括号/取整等；按高度放大字形，不是完整 TeX delimiter assembly |
| 重音 | hat/bar/vec/dot/ddot/overline/underline/overbrace；双点的横向测量纳入宽度，避免短 operand 的 dots 越出公式框；overbrace 用分段线近似 |
| 矩阵 | `matrix/pmatrix/bmatrix/Bmatrix/vmatrix/Vmatrix`；真实行列单元布局、最大列宽、行间距及 nested cells；Vmatrix 是目录之外已有独立语法扩展 |
| 文字与样式 | roman/bold/italic/sans/monospace、text-group 空白、显式数学 spacing、code point 与原始 UTF-16 错误位置 |
| 预算及错误 | 单次 source 65,536 UTF-16、sequence 解析项预算 16,384、readSequence 嵌套预算 64；坏环境、未配对结束符、预算/测量失败显式报错；这不是完整 AST 内存字节预算 |

这些是可审阅实现分支；没有运行全模板 parser/layout 草稿，不能将 71 条静态源匹配宣称为全模板数学效果已验收。完整数学 shaping、font ascent/descent、数学字距、复杂 script positioning、TeX style 和宏环境仍有缺口。`displaystyle/textstyle` 当前只消费声明，不完整实现其排版语义；默认普通数学字母/嵌套字体的精确效果尚待补齐。

正文继续输出共用 text/line draw commands，供现有 Canvas2D / WebGPU scene 使用。新共享预览是 **Canvas2D**，取 ConfigProvider token、DPR 和实际 Canvas 文本宽度；字体 ready/loadingdone 与 resize 触发重画，卸载移除监听，迟到回调检查 disposed。未宣称预览使用 WebGPU。

模块归属遵循 [完整工作台架构](../../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)：模型/源保持在 document feature，纯 parser/layout 与组件分离，UI 只编排 draft 和命令。当前解析与数学布局仍是主线程有界 TypeScript，未接入数学 Worker/Rust/WASM 域；[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)中的版本化控制 DTO、字体版本、取消、lease/释放、memory.grow 视图和复制计量属于后续工作。此批没有引入 WASM buffer 或声称端到端零拷贝/性能达标。

## 5. Ant UI 与输入技术依据

新增及修改交互控件均为 `antdv-next`：Modal、Form/FormItem、Select、Input.TextArea、Button、Tooltip、Alert。Canvas 只绘制公式预览和编辑内容，不替代 Form 或文本输入。gallery 的 Form 与 source Form 是 sibling，避免 nested forms。

Input.TextArea 使用实际 v1.8.2 的 `TextAreaRef.nativeElement` 找到它自己的 textarea，读取 selectionStart/End 并调用 setSelectionRange。该公开 ref 可 focus/blur，但没有公开 source-range setter，因此保留这段最小原生 selection 桥；没有新增原生 textarea 或自行管理控件树。source/ref 只属于当前临时 editor。compositionEnd 延后到 nextTick，并检查 composition generation，避免旧结束事件清掉新 composition；打开动画焦点恢复同样避开 composing。

为允许编辑期间使用顶部上下文 Ribbon，Modal 设置 mask=false、scrollLock=false、focus trap=false，wrap pointer-events none / dialog auto；source draft 与内容目标 guards 仍保留。此选择与上游 range-attached popup 有布局差异，**尚未验证实际弹层点击、键盘、关闭和焦点恢复**。

本地 gallery 为三列、描述标签/tooltip 与 Ant Button，使用 440px 宿主、90px 项高和局部滚动；上游 desktop gallery 420px、min-height 80px，aria-label 为类别+source。本地 label 使用中文/英文语义名称，中文简繁体已覆盖 71 项与 11 类；源码 tooltip 保留。类别 icon 统一 Lucide Sigma，与上游 11 个专用数学 icon 不同。布局、可发现性、主题、可访问性与 IME 效果仍待统一浏览器验收。

## 6. 修改文件与接口

路径均相对 `packages/office-studio/src`，保留其他 Docs / Sheets / Charts 在途改动；没有改 package/lock、后端或中央清单。

| 责任 | 主要文件 |
| --- | --- |
| 纯目录、parser、layout、类型 | `features/document/latex/latex-templates.ts`、`latex-parser.ts`、`latex-layout.ts`、`latex-types.ts`；`latex-formula.ts` 保持旧路径 re-export 与模型逻辑 |
| runtime owner 与 operation | `api/document-latex-template-editors.ts`、`document-latex-commands.ts`、`facade.ts`、`office-studio.ts` |
| 上下文及 UI | `components/document/document-latex-ribbon.ts`、`document-ribbon.ts`、`DocumentLatexTemplateGallery.vue`、`DocumentLatexPreview.vue`、`DocumentLatexPanel.vue`；`components/DocumentWorkbench.vue` |
| locale / 公共入口 | `i18n/docs/locales/document-latex.ts`、`index.ts` |

新增公共导出：目录、类别/模板/插入结果类型、planner/getters/category guard、public/local template operation 常量、operation/request 类型和 runtime handle 类型。FDocument 的 bind/insert 方法是本地 UI 扩展，未冒称上游同名 Facade。Math AST 新增节点是解析/布局内部接口，模型仍保存原始 latex。前批 `documentPublicParagraphs().text` / segment alias 读取变化继续见 [前批证据](DOC-15-latex-editing.md#5-文件归属及公开接口变化)。

## 7. 新增未运行草稿与统计口径

| 草稿 | 静态 test 声明 | 计划展开用例 | 内容 |
| --- | ---: | ---: | --- |
| [document-latex-templates.contract.test.mjs](../../packages/office-studio/tests/document-latex-templates.contract.test.mjs) | 11 | 81 | 一个目录循环声明展开 71 模板，另 10 条目录/source-range/UTF-16/locale 合约 |
| [document-latex-math-layout.contract.test.mjs](../../packages/office-studio/tests/document-latex-math-layout.contract.test.mjs) | 14 | 14 | 矩阵、定界符、根指数、limits、integral、重音、font、预算、错误与共用绘制协议 |
| [document-latex-template-facade.contract.test.mjs](../../packages/office-studio/tests/document-latex-template-facade.contract.test.mjs) | 12 | 12 | unit/studio owner、替换与销毁、revision/history guards、取消/no-op、canonical apply、稳定 ID 及 nested/header/note 保存 |
| [document-latex-templates.browser.test.mjs](../../packages/office-studio/tests/document-latex-templates.browser.test.mjs) | 10 | 10 | 真实 Ant/Canvas、上下文生命周期、Modal 打开时点顶部、source 占位/替换、grid 11 类、unicode、composition/stale、辅助 hosts、standalone 与预览主题 |
| 合计 | **47** | **117** | 全部未运行，计划展开数不是通过数 |

复用 [document-latex.html](../../packages/office-studio/tests/fixtures/document-latex.html)，新增 fixture-only API template 请求与 grid 参数。修正前批 text-group AST 草稿为 styled.value，并把旧未闭合 matrix 用例描述改为当前语义；没有执行这些草稿。前批新增 43 条和原有 7 条不计入本批 117。

静态核对确认目录 71、分类计数、上游 string/offset 匹配、菜单 ID 和文件路径；没有 Vue/TypeScript 解析器、test runner、浏览器、文件重开或性能执行结果。浏览器草稿的 synthetic composition 只检查 guard，不代表真实 Windows IME 验收。

## 8. 剩余工作与中央同步

- 本模板批时点的段落 Layout、insert/edit quickBottom、insert、浮动文字样式与移动 popup 缺口，已在 [后续五挂载 / 移动双动作批](DOC-15-latex-popmenus-mobile.md)落盘；该批新增 42 条声明 / 80 个计划用例，仍未运行，窄屏溢出和最终交互继续待验收。
- 全模板预览、复杂数学字体/样式、宏/环境、脚本位置、完整 TeX shaping 与更广语法仍未验收，部分尚待实现。
- DOCX/OMML 原生解析写出、真实 Office 数学样本、资源与内部 source 多样式的互操作保真仍有缺口；复用本项目 parser/writer，由主任务保留后置需求。
- 数学 Worker/Rust/WASM、取消/版本/字体缓存/内存/复制预算和大样本性能仍待实现与测量。
- 真实明暗主题、Modal 与 Ribbon 弹层顺序、焦点恢复、IME、辅助 host、内容锁和 observer/stale 回归均未验收。

中央清单由主任务同步两份 DOC-15 证据，记录常规根无增量、1 上下文分类/11 根/71 目录项、本模板批 47 条测试声明/117 个计划用例及前编辑批 43 条声明的独立口径；继续保留“未验收”。模板/入口数量不能作为数学功能完成、Office 保真或性能达标证明。此批到此收尾，不扩大到下一批。

以上保留模板批口径。最新入口同步还需关联 [popup 批记录](DOC-15-latex-popmenus-mobile.md#6-向主任务交付)，五个有效挂载、移动 2 动作与该批 42 声明 / 80 计划用例分别计数，不重复增加 selector 或模板。

