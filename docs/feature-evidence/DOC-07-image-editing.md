# DOC-07 图片编辑与对象入口

记录日期：2026-11-26。状态：**开发中，未验收**。包含第一批图片编辑、第二批 flip / 段内 atom / 双侧文字，以及第三批图片多选 / 对齐 / 分布 / 整体变换。三批均仅读取、编辑文本，不运行测试、类型检查、构建、dev、浏览器、Rust/WASM、基准或 `git diff --check`。

## 1. 实现前的入口清点

固定 OSS 为 Univer v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`，只读目录 `E:\rust_test\univer`。下表先按宿主自上而下、组内 order/index 排列，再实现命令与模型；数量不作为功能验收证明。

| 类别 / 宿主 | 入口、顺序与位置 | 命令 / 状态 / 固定源码 |
| --- | --- | --- |
| 常规 INSERT / MEDIA | `doc.menu.image`，order 0，grid row 1 / column 4；classic 媒体组图片按钮 | `doc.command.insert-float-image`；非 Docs hidden，无文字范围 / 表格内 disabled；`docs-drawing-ui/src/menu/schema.ts`、`image.menu.ts` |
| 段落 Layout / insert-below | 图片 order 5 | 同一 insert-float-image；`schema.ts` |
| 段落 empty / Layout | 图片 order 5 | 同上 |
| 段落 content-insert / Layout | 图片 order 5 | 同上 |
| 段落 T insert / insert | 图片 order 1，形状父项 order 2 单列 | 同上；图片与形状不混计 |
| 段落 T insert-below / insert | 图片 order 1，`.below` 别名，`paragraphMenuPlacement: below` | 同一命令，位置参数；`image.menu.ts` |
| 图片浮动 fallback | 编辑 index 0、裁剪 1、删除 2、重置 3 | `doc.operation.edit-doc-image`、`sheet.operation.open-image-crop`、`doc.command.remove-drawing`、`sheet.operation.image-reset-size`；重置上游固定 disabled/TODO；`drawing-popup-menu.controller.ts` |
| 绘图右键 OTHERS | group order 1 / arrange 2 / align 3 | `drawing-ui/src/menu/schema.ts`；三父项与子项分开统计 |
| group 子项 | 组合 0、取消组合 1 | 两项；本批图片目标不据此声称形状组合完成 |
| arrange 子项 | 置顶 0、上移一层 1、置底 2、下移一层 3 | 四项；`drawing-ui/src/menu/arrange.menu.ts`、`drawing-arrange.operation.ts` |
| align 子项 | 左 0、水平居中 1、右 2、上 3、垂直居中 4、下 5、水平分布 6、垂直分布 7 | 八项；`align.menu.ts`、`drawing-align.operation.ts` |

分类口径：常规根 **1**（已有本地图片入口，不新增根），段落有效挂载 **5**，图片浮动 fallback **4**，绘图右键父项 **3** / 子项 **14**。移动图片菜单与 sidebar 是呈现变体，单列核对。`DocDrawingFloatingToolbarAdapterService` 按 adapter 优先级替换 fallback，`FloatingObjectToolbarPosition.DOC` 再追加动态扩展；尚未运行插件装配，不能推导完整可见浮动总数。group / align 与形状及多对象有关，不算已完成图片功能。

上表源码简称相对 `E:\rust_test\univer\packages`。本地段落 mount 为 `layout-below`、`content-insert`、`empty-paragraph`、`quick-insert`、`quick-insert-below`；前三者图片 order 5，后两者 order 1。上游只有 `quick-insert-below` 的图片工厂明确传入 `paragraphMenuPlacement: below`；本地 `layout-below` / `content-insert` 也传入该参数以承载下方插入语义，这是宿主适配差异，未经过浏览器核验。

进一步追踪的固定源码：

| 契约 | 固定版来源（相对上游 `packages`） | 本地处理 / 未核实项 |
| --- | --- | --- |
| 图片公开对象、插入与查询 | `docs-drawing/src/facade/f-document.ts`、`f-document-image.ts` | 同名 Facade 和内容命令已落盘；UUID 来源需本项目媒体解析服务 |
| 尺寸、角度、定位、裁剪、翻转 | `docs-drawing/src/commands/commands/update-doc-drawing-transform.command.ts`、`core/src/types/interfaces/i-drawing.ts` | 原始 `srcRect` 是可见框外 CSS-pixel 偏移；本地模型使用比例，写出转回像素；第二批已映射 boolean flipX / flipY，全部未验收 |
| 五种环绕 | `docs-drawing/src/commands/commands/update-doc-drawing-wrapping-style.command.ts` | 第二批保留段内图片 atom，square 支持按阅读顺序消费全部可用区间；复杂锚点和分页效果未验收 |
| 多对象排列 | `docs-drawing/src/commands/commands/set-drawing-arrange.command.ts`、`drawing-ui/src/commands/operations/drawing-arrange.operation.ts`、`core/src/common/drawing-order.ts` | `ArrangeTypeEnum` 为 forward=0、backward=1、front=2、back=3；本地同容器排列采用相同顺序和索引 floor/clamp |
| 图片菜单、浮动和移动呈现 | `docs-drawing-ui/src/menu/image.menu.ts`、`menu/drawing-popup-menu.controller.ts`、`controllers/mobile/drawing-popup-menu.controller.ts` | 本地 Ant 浮动工具条、右键菜单、格式 Drawer、来源 Modal 已落盘；动态 adapter、侧栏、移动端和全部插件装配未核验 |

## 2. 模型、API 与计算边界

第二批新插入和 native drawing 读取使用段内图片 atom；旧 `officeStudio.type: image` 输出及旧图片内容控件的独立 image block 继续可读写。`DocumentImage` 包含来源与资源标识、替代文字、可见宽高、原始宽高、角度、flipX / flipY、比例锁、水平 / 垂直定位、五种环绕、环绕侧、四侧文字距离、比例裁剪及层级。atom 为一个 U+FFFC run，其 `metadata.imageAtom` 保存 version / imageId / runId / image；浮动图片也保留同一逻辑 atom，只改变派生 advance。原始 `customBlock` / drawing 记录留在实际图片 run 或 legacy block metadata 中，公开 body 仍用 `customBlocks`、`drawings` 与 `drawingsOrder` 保存。解析 DTO、领域图片、派生布局和共享 ImageCommand 分开。

图片 IO 由主任务明确授权复用 `features/sheet-images/sheet-image-io.ts` 的 `prepareSheetImageFiles` / `prepareSheetImageUrl` / `prepareSheetImageResources` / `decodeSheetImageSource`，适配到 Docs 模型。读取和解码已从普通图片 Workbench 表单移到服务；旧图片内容控件的读入也复用同一链路。单张 5 MiB、批次 32 MiB / 1000 张、16,777,216 像素沿用共享 IO。File、URL、BASE64 或资源 DTO 经格式与实际解码尺寸检查；URL 读取后内嵌字节并保留原始来源。`DocumentImagePatch` 不接受 `src`，来源替换走异步 `replace`。UUID 解析、EXIF / DPI、Office 原生媒体关系及跨设备资源仍有缺口。

图片排列、布局、命中与手势当前在主线程，布局是有界派生数据；持久变换通过同一内容命令 / 文档历史。共享场景 ImageCommand 的 `sourceRect` 使用 normalized fractions，`rotation` 使用度；源像素裁剪只在公开 drawing 边界转换。Docs 独立栅格缓存已移除，使用共享 Canvas2D / WebGPU 图片变换协议；没有运行两后端验证。AbortSignal 与内容 / 输入版本取消已有代码；Worker / Rust / WASM、lease、复制字节、内存和大样本性能契约仍未验收，不宣称优化或零拷贝。

## 3. 实现记录

下列路径相对 `packages/office-studio/src`。本批已落盘代码，全部仍为 **未验收**；共享 API 注册、导出及共享后端由主任务接线。

| 层 / 文件 | 已落盘行为 |
| --- | --- |
| `features/document/images/document-image-model.ts`、`features/document/model/document-model.ts` | 图片参数校验、比例 / 像素裁剪转换、资源预算、递归发现和排列纯函数；递归含装饰块、列、表格 cell 与合并延续内容 |
| `features/document/commands/document-image-commands.ts`、`document-commands.ts`、`api/document-editor-commands.ts` | update / replace / reset / remove / arrange / batch 内容命令及历史标签；候选模型完成锁、权限与范围校验后才提交 |
| `api/document-images.ts`、`document-image-commands.ts`、`document-image-editors.ts` | Facade 图片对象、异步 IO、公开命令、编辑器绑定和来源任务取消；按 Studio / 文档 / segment / 宿主隔离 |
| `api/document-image-body.ts`、`document-body-read.ts`、`document-body-write.ts`、`document-body-codec.ts` | 图片 drawing 解析 / 写出适配、裁剪像素换算、原始字段保留、图片层级槽位更新与删除输出 |
| `features/document/layout/document-image-layout.ts`、`document-block-measurement.ts`、`document-block-layout.ts`、`document-paragraph-layout.ts`、`document-text-layout.ts`、`document-layout-types.ts`、`document-layout.ts` | 旋转几何、可见框、定位、环绕障碍、图片布局、浮动边界；表格 / 辅助宿主逐行读取 wrap band，宽度变化只重测未放置 UTF-16 后缀 |
| `features/document/controllers/document-image-canvas.ts`、`features/document/rendering/document-scene.ts` | 旋转命中、角点 resize、拖拽预览、共享图片绘制与选区；pointer-up 一次内容命令，取消不发布预览 |
| `components/document/DocumentImageOverlay.vue`、`DocumentImagePanel.vue`、`use-document-image-editor.ts`、`DocumentCanvasViewport.vue`、`components/DocumentWorkbench.vue` | 图片选择、双击格式、浮动与右键操作、裁剪 / 来源面板、Canvas 编辑手势、焦点恢复和 composition 保护 |
| `components/document/document-ribbon.ts`、`document-structure-ribbon.ts` | 常规媒体根及五个段落 mount 共用真实插入命令和上下文能力判断 |

接口：`FDocument.insertImage(options): Promise<FDocumentImage | null>`、`getImage(id, segmentId?)`、`getImages(segmentId?)`、`bindImageEditor(editor)`；`FDocumentImage` 提供来源 / 尺寸 / 角度 / 定位 / 完整图片副本、update、setSize、setRotate、getFlipX / getFlipY / setFlipX / setFlipY、setPositionH / V、setWrappingStyle、setCrop、四种 arrange、resetSize、remove 和异步 replace。`getImageData()` 以原始 drawing 为底保留 transform / size / position 的未知字段并返回副本。

公开插入 `doc.command.insert-float-image` 返回 `Promise<boolean>`；无 source 时打开已绑定编辑器，有 source 时完成资源读取与内容插入。公开 transform / wrapping / arrange / remove 命令映射到同一 Docs 内容命令，edit / crop / replace operation 打开 Ant 面板。unit、segment / subUnit、image / drawing 身份冲突、过期 revision / history generation、缺失目标、无效参数均显式拒绝；不以消息提示充当成功。

多对象命令 `doc.command.arrange-images` 接收 `imageIds` 与 `arrange` 或单对象 `zOrder`。选中图片内部顺序保留；有限绝对索引 floor 后 clamp，负数到 0、过大到末尾，移动后统一写回图片层级。第一批纯排序函数沿用上游 NaN 到 0 的行为；第三批内容命令边界明确拒绝 NaN / Infinity，防止无效 API 输入产生层级编辑。仅允许同一容器，任何受影响 peer 的权限或内容控件锁失败均不部分提交；一次候选提交产生一次内容历史。第三批已增加多选 UI 和四种批量 Facade，见第 6 节，均未验收。

图片控件锁也检查旧 `metadata.pictureBlockId`。变化图片同步到对应 `metadata.pictureImage`；删除后清除图片指针 / 数据，保留 inline 控件锚点和未知 metadata，设置空值与 placeholder；整块控件删除沿用现有范围协调规则。不改写无关控件。替换保留可见框、角度、位置、环绕和层级，更新资源身份并清除裁剪；reset 使用已读原始尺寸，缺少原始尺寸时拒绝。

保存 / 重开代码保留图片身份、原始来源、资源尺寸、裁剪、定位及未知 drawing 字段；未编辑的 native crop 保留原值，修改尺寸或裁剪后按当前可见框换算像素，清除裁剪写出 `null`。未知-only `srcRect` 保留原始数据但不派生假裁剪。`drawingsOrder` 只重排图片槽位，非图片 drawing 槽位保留；这不等于已实现跨类型图形层级。实际 Office 原生文件写出及重开没有验收。

全部新增图片表单、按钮、菜单、浮动工具条、Tooltip、Drawer、Modal、Upload 与输入使用 `antdv-next`；Canvas 文本 / IME 桥仍为 Ant Input.TextArea。图片面板捕获文本和数字字段 composition，期间禁止提交、Escape、遮罩关闭、关闭按钮、取消和清除裁剪；打开后聚焦来源或替代文字，通过当前内容 / 输入版本判断是否恢复 Canvas 焦点。文字输入或 composition 清除图片选中，移动 beforeinput 删除走图片内容命令。控件沿用统一 ConfigProvider / token；真实明暗主题、焦点及 Windows 原生 IME 尚未验证。

`features/document/rendering/document-image-render-cache.ts` 已删除 / 不再存在；共享 IO、`api/facade.ts`、`api/office-studio.ts`、`index.ts`、`rendering/contracts.ts`、`canvas2d.ts`、`webgpu.ts` 属主任务，不在本子任务文件归属中。

## 4. 验证与剩余目标

仅静态阅读与编辑；没有运行测试、类型检查、构建、dev、浏览器、Rust / Cargo / WASM、基准或 `git diff --check`。下列第一批 **44 个 test 为未运行草稿**，数量保留；第二批另计，见第 5 节。它们不是通过结果；浏览器夹具也未启动。

| 草稿（相对 `packages/office-studio/tests`） | test 数 | 草稿覆盖目标 |
| --- | ---: | --- |
| `document-image-model.contract.test.mjs` | 10 | 裁剪 / 预算 / 比例、递归发现、replace / reset / remove、batch 原子性、多对象排列、权限 / 锁、inline picture 镜像与锚点保真 |
| `document-image-body.contract.test.mjs` | 6 | native crop / 未知字段 / 定位、图片层级槽位、保存重开与删除 |
| `document-image-canvas.contract.test.mjs` | 8 | 旋转命中 / 角点、对角固定 resize、drag / stale、文字优先级、wrap band、表格 / 辅助 host 与 UTF-16 offset、共享场景预览 |
| `document-image-facade.contract.test.mjs` | 11 | 真 Facade / 命令 / 历史、File / URL、abort / stale、相同 ID 的 Studio 隔离、五个段落 mount、常规 Ribbon 根、header / note 保存重开 |
| `document-image-workbench.browser.test.mjs` | 9 | Ant 上传 / URL、Canvas 编辑、crop / reset / delete、history、theme / 合成 composition / 焦点、host 隔离与 KeepAlive |
| `fixtures/document-images.html` | — | Canvas 生成真实 PNG 的浏览器夹具；原 Studio / standalone / table / header / note / locked 场景保留；第二批增加 inline / wrap 场景 |

Facade 草稿使用 Vite SSR 与假 Image / FileReader；已静态修正不存在的方法名、取消后假解码队列及双 Studio 等待数量。浏览器草稿以真实 Canvas / Ant UI 为目标；合成 composition 不等于 Windows 原生 IME 验收。没有实际文件样本、两渲染后端或运行结果。

剩余交付目标保留完整范围：

- 段内 image atom、两侧区间文字与 flipX / flipY 已有第二批实现及草稿，全部未验收。复杂旋转 / 大对象、精确行高的溢出、文本对齐、RTL、行边界 affinity、连续上下导航、段落内多个浮动锚点及真实分页仍需补场景。
- 第三批已落盘多选 UI、六种 align、两种 distribute、整体平移 / 均匀缩放 / 旋转；全部未验收。group、非均匀仿射、跨容器 / 跨类型层级仍须实现；非图片 drawing 的槽位保留不提供这些语义。
- 跨页图片拖动与逻辑锚点迁移、复杂 SDT / 嵌套和分页；表格 / 辅助宿主 page / margin 定位仍使用相对测量参考。
- UUID / 跨设备媒体解析、EXIF / DPI、Office 原生媒体关系、复杂环绕 / 锚点 / crop 与源资源冲突样本；转换继续扩展本项目解析 / 写出链路。
- Worker / Rust / WASM 图片布局、内存 / lease / 复制次数与字节 / 取消契约、首屏 / 滚动 / 缩放 / 编辑和大样本前后基准。
- Canvas2D / WebGPU、真实 Office 文件保存重开、主题、焦点、移动尺寸、Windows 原生 IME 和插件动态入口装配。

后端需求：本项目媒体资源解析与跨设备资源持久化、Office 原生媒体 / 锚点 / 裁剪关系适配、图片布局 Worker / Rust / WASM 协议及性能 / 内存计量。此子任务只更新独立证据文件，中央清单与 `BACKEND_REQUIREMENTS.md` 由主任务同步。不能将本批标为 DOC-07 图片全量完成。

## 5. 第二批：翻转、段内图片与双侧物理行

本节记录续作已落盘行为，均为 **开发中、未验收**。第一批 44 个草稿保留；第二批新增 **36 个 test 草稿**，两批合计 80 个，未运行。

### 5.1 固定来源与入口增量

固定版本和 commit 同第 1 节。下面路径相对 `E:\rust_test\univer\packages`，只静态读取源码或测试文本，没有执行上游测试。

| 依据 | 固定来源与符号 | 本地差异 / 未核实项 |
| --- | --- | --- |
| boolean flip transform | `docs-drawing/src/commands/commands/update-doc-drawing-transform.command.ts:39`，`IDrawingDocTransform` / `UpdateDrawingDocTransformCommand`；`core/src/types/interfaces/i-drawing.ts` | flipX / flipY 写入 docTransform；旧文档没有字段时上游 insertOp。本地候选命令直接更新 typed image，保存边界再写 drawing |
| custom block shaping | `engine-render/src/components/docs/layout/block/paragraph/shaping.ts:227` | 上游 inline custom block 一个 glyph，非 inline 为零 advance；本地图片 atom 采用相同可观察占位契约，不复制其 skeleton 内部对象 |
| 同一物理行多个区间 | `engine-render/src/components/docs/layout/model/line.ts:119`，`lineSke.divides` / `_calculateDividesByDrawings` | 本地 LayoutLine 以 rowKey / fragmentIndex / fragmentCount 表达区间，共享源后缀、行高和 baseline |
| 图片底边与文本 baseline | `engine-render/src/components/docs/layout/block/paragraph/__tests__/layout-ruler.spec.ts:2898` | 本地 inline 图片底边对齐共享 baseline；exact 行高小于图片时仍保留对象几何，裁剪 / 重叠最终效果未核验 |
| 相邻 custom block 与 topBottom | `engine-render/src/components/docs/layout/block/paragraph/linebreaking.ts:162` | 本地加入区间重测和有限分页保护；复杂相邻对象、分节 / 注释预留仍需真实样本 |

本批不新增 Ribbon 根、段落 mount、右键根或浮动根；第 1 节各分类数量不变。图片格式 Drawer 的旋转字段后增加 **2 个 Ant Switch**，顺序为“水平翻转”→“垂直翻转”，两列表单各占一项。值来自实际图片；只读、陈旧输入位置或任务进行中时沿既有 Form / Drawer 规则禁用；输入法组合期间禁止提交。它们提交同一 `doc.command.update-image` patch，并映射既有公开 `doc.command.update-drawing-doc-transform` 的 flipX / flipY key。表单控件数不计入 Ribbon 根或完整功能数。

### 5.2 模型、命令与保存

| 文件（相对 `packages/office-studio/src`） | 续作行为与接口 |
| --- | --- |
| `features/document/images/document-image-atom.ts`、`document-image-model.ts`、`features/document/model/document-model.ts` | U+FFFC 单 run 图片 atom；递归 entries 返回明确 imageId 和可选 inline source target / offsets。`entry.block` 是宿主段落，不能再当作图片 ID；legacy image block 保留 |
| `api/document-images.ts`、`api/document-image-commands.ts`、`api/document-image-body.ts`、`features/document/commands/document-image-commands.ts` | insert options、对象 getters / setters、公开 transform、normalize / patch 和 body 均传 flip；replace 保留翻转，reset 清两轴；batch 无效后项不应部分发布 |
| `features/document/commands/document-commands.ts`、`features/document/model/document-run-range.ts` | insert-image 用普通文字 mutation 插入 atom；below 用含 atom 的新段落。普通文字继承清除 image / drawing metadata，compactRuns 不合并图片 atom；Backspace / Delete / 范围替换复用真实文字范围协调 |
| `api/document-body-read.ts`、`document-body-write.ts`、`document-body-identities.ts` | native inline / floating 图片保留段落 atom；写 officeStudio.imageAtom marker、独立 drawing ID / flip / layer；重复绘图身份显式拒绝，段落身份与图片身份分开 |
| `features/document/services/document-block-paste.ts`、`document-clipboard.ts` | 复制 atom 分配新 run / drawing ID；真实 IMG 与 data-office-image payload 保留资源、crop、flip；选区 copy / cut / paste 共用真实 rich payload；destination mode 保留图片，text mode 投影 alt |
| `features/document/model/document-table-text.ts` | `documentTableTextMatrix(table, range?, projection?)` 增加可选纯文本投影，默认仍返回真实 run.text；剪贴板投影递归表格 / 续格中的图片 alt，不改源 U+FFFC 地址 |
| `features/document/permissions/document-permissions.ts`、`components/document/use-document-image-editor.ts` | 图片 permission targets 包含宿主和 image entity；邻接普通文字不自动锁整段，范围实际覆盖 atom 才加图片 entity；对象面板检查 view / edit |
| `features/document/services/document-image-controls.ts`、`services/content-controls.ts` | 对象命令和普通文字候选共用 pictureImage 镜像同步，删除清旧图片指针并保留未知 metadata；existing picture 替换已可找到实际 atom run。picture 控件的最终输入语义仍由内容控件服务校验，不能据此声称允许任意文字替换图片控件；新建旧图片控件尚未全部迁移到 atom |

`executeDocumentImageCommand` 的第五参数为现有文字 mutation：`(model, selection, text) => DocumentSelection`。删除 atom 因此协调字符地址、字段、评论和控件范围，最终仍由共享文档事务提交。新增 `synchronizeDocumentPictureControlImages(before, candidate)` 服务只处理图片实体改变 / 删除的镜像；有 picture 指针的文档通过同一 detached command candidate 更新，不新增历史源。

保存使用同一 public body / snapshot 管线：atom marker、源地址、imageId、drawing、flip、crop、未知来源和 layer 一并写回。HTML 复制为真实 IMG；纯文本使用 alt。尚无真实 Office 二进制 / 独立重开证据；不能用 JSON 草稿替代 DOCX 保真验收。

### 5.3 布局与 Canvas 输入

`document-text-layout.ts` 将 inline 图片作为一个有宽度 glyph，底边对齐文本 baseline；floating atom 为零 advance。`document-image-layout.ts` 的 `appendDocumentInlineImages` 输出实际图片和 source anchor，场景跳过 U+FFFC 文字并发送共享 flipX / flipY。裁剪与翻转仍交由主任务维护的共享 Canvas2D / WebGPU 后端，没有恢复 Docs 独立栅格缓存。

`documentImageWrapBands` 按阅读顺序返回可用区间；明确要求单区间的兼容方法仍选最宽。`measureDocumentImageTextRow` 逐区间消费同一原始 UTF-16 后缀，显式换行结束当前物理行，尾换行保留空行。body 和表格 / 辅助宿主相对布局都只消费一次行高；列表、checklist、code、note label、原始缩进保留。`documentPhysicalLineFragments` 按 rowKey、page、section、column、target、host、cell、repeatedHeader 归并，`documentPhysicalLineNavigation` 将 Home / End 定位到整行首尾，上下跳过当前区间组。

`DocumentSectionFlow.imageWrapBandsAt` 是不推进 cursor 的预测接口；实际 wrap overflow 重新核对 page / column / section。首行尚未提交而 frame 改变时移除旧浮动图片与障碍后重建；整栏 topBottom 障碍将文字推向后续 frame 时保留原图片锚页，避免图片随文字无限迁移。这些是代码策略，复杂分页、浮动锚点准确落在原始段内行、widow / keep / note 预留仍未验收。

Canvas 单击、右键、双击图片统一先发布实际 atom source range，等 Vue 选区更新后重新捕获当前输入会话，核对 source / host / revision 后发对象意图。inline 图片方向键转文字导航；浮动图片保留对象移动。Escape 清对象选中并保留文字范围，文字 / composition 输入沿 Ant Input.TextArea。copy / cut 提供真实图片 HTML，并分别核对 copy 和 scoped view；图片删除使用同一内容命令。被 view 拒绝的图片查询与编辑入口已有守卫，完整 Canvas 内容脱敏 / 权限渲染还没有验收。

本批所有新增可替换 UI 均使用 `antdv-next`（两轴 Switch、已有 Form / Drawer / Input.TextArea），沿统一 ConfigProvider 和主题 token。没有新增原生表单控件；夹具里的 ClipboardEvent / DataTransfer 和 composition 仅用于后续测试。真实明暗主题、焦点、输入法与两后端没有运行。

### 5.4 第二批草稿与未完成项

| 新增草稿（相对 `packages/office-studio/tests`） | test 数 | 草稿目标 |
| --- | ---: | --- |
| `document-image-inline.contract.test.mjs` | 12 | boolean flip、replace / reset、插入 / 邻接打字、Backspace / Delete、控件范围、同段 / 表格 ID、rich / destination / text paste、entity lock、clear format、表格 alt 投影与 picture 镜像候选 |
| `document-image-inline-body.contract.test.mjs` | 6 | native inline / floating atom、UTF-16 / flip、重复 ID、保存 / 删除、wrapping 切换、同段多图及 layer |
| `document-image-inline-layout.contract.test.mjs` | 8 | 混合字号 / 图片 baseline / scene flip、双区间消费、emoji / 显式换行 / 尾空行、物理行导航、table / auxiliary source、缩进、同页换栏、超高 topBottom 有限进展 |
| `document-image-inline-facade.contract.test.mjs` | 5 | textRange、flip 批次失败原子性、save / undo / redo / reopen、image entity view、header / note segment |
| `document-image-inline.browser.test.mjs` | 5 | 两个 Ant flip 开关 / theme / focus / history、点击 atom 范围及键盘 / 邻接文字、rich copy / cut / paste、合成 composition、双侧物理行 Home / End / Down |
| `fixtures/document-images.html`（扩展） | — | `?inline` / `?wrap` 真实 PNG 场景，selectRange / layoutRows / synthetic clipboard helper；未启动 |

上述草稿均未运行；Facade 草稿的 Vite SSR / 假 Image 和浏览器草稿的 Vite / Edge 仅写在测试代码中，没有实际启动。未运行类型检查、测试、构建、dev、浏览器、Rust / Cargo / WASM、基准或 `git diff --check`。

剩余目标：实际 browser / Canvas2D / WebGPU / Windows 原生 IME / 系统剪贴板 / Office 文件保存重开；多图浮动锚点、分页 / 分节 / note / exact 行高与对象溢出、复杂 RTL / 对齐 / 行边界 affinity；完整 picture 控件 atom 迁移和复杂 SDT；第三批多选 / align / distribute 的运行验收及 group / 跨容器与跨类型层级；媒体 UUID / EXIF / DPI / 原生关系；Worker / Rust / WASM 布局与取消、内存 lease / 复制计量和大样本同数据集性能基准。全部完整功能及 Pro 范围继续保留。

<a id="image-batch"></a>

## 6. 第三批：图片多选、对齐 / 分布与整体变换

本节记录真实模型和交互代码的续作，不是运行验证。第三批新增 **57 个 `test(...)` 声明草稿**：模型 / Canvas / scene 契约 19、Facade / 历史 / 布局借用契约 20、浏览器交互 18。第一批 44、第二批 36 保留，三批合计 **137 个未运行声明草稿**。一个模型 test 内有六种对齐黄金坐标，一个浏览器 test 内循环八个菜单场景；这些内部场景不重复计入 test 声明数，也不计入已完成功能数。

### 6.1 固定源码与实现差异

固定源码仍为 v1.6.0 / `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。下列路径相对只读 `E:\rust_test\univer\packages`，本轮只读取，不执行上游代码。

| 契约 | 固定来源 | 本地行为 / 未核实项 |
| --- | --- | --- |
| 对齐枚举与八个叶入口 | `drawing-ui/src/commands/operations/drawing-align.operation.ts`，`AlignType` / `SetDrawingAlign*Operation` | 保留 default=`'0'`、left / center / right / top / middle / bottom=`'1'…'6'`、horizon / vertical=`'7'/'8'`。共享 `sheet.operation.set-image-align` handler 仅 `return true`，未找到 Docs 覆盖；该 stub 不是本地功能证据。本地独立实现真实图片几何命令，不把上游空 handler 接成成功提示 |
| 对齐父项显隐 | `drawing-ui/src/menu/align.menu.ts`，`DrawingAlignContextMenuItemFactory` | 上游 focus 少于两项隐藏父项；本地仅多选显示对齐根，分布少于三项禁用，命令同样检查。三项门槛属于本地真实分布语义，不假称上游 stub 已执行 |
| 右键组与子项顺序 | `drawing-ui/src/menu/schema.ts` | Drawing / Others 下 group order 1、arrange 2、align 3；align 八叶顺序为左 / 水平居中 / 右 / 上 / 垂直居中 / 下 / 水平分布 / 垂直分布。本地多图右键排列在对齐前，group 保留为未完成目标 |
| 多对象层级 | `docs-drawing/src/commands/commands/set-drawing-arrange.command.ts`、`core/src/common/drawing-order.ts` | 保留源层级中选中对象的相对顺序，front / back 分区，forward / backward 按相邻块移动；同容器和所有受影响 peer 权限 / 锁统一验证。绝对索引的有限数值校验是本地命令边界差异 |
| 多对象几何效果 | 同上入口 / 枚举及图片公开 transform 契约 | 本地以可见旋转包围框对齐，等边间距分布，整体均匀缩放并绕选区中心旋转；图片局部宽高、crop、flip 和资源身份独立保留。Office 可观察行为的复杂重叠 / 分布、跨页 / 锚点和混合对象仍需样本核实，不能据当前算法宣称完全等价 |

上游链接可按同一 commit 复查：[对齐 operation](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/drawing-ui/src/commands/operations/drawing-align.operation.ts)、[对齐菜单](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/drawing-ui/src/menu/align.menu.ts)、[绘图 schema](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/drawing-ui/src/menu/schema.ts)、[Docs 排列 command](https://github.com/dream-num/univer/blob/eaf4cfa04fabdbcfc922c692f3450459f8a29546/packages/docs-drawing/src/commands/commands/set-drawing-arrange.command.ts)。公开枚举和菜单顺序有源码证据；线上完整装配、Office 文件和最终像素效果没有运行证据。

### 6.2 本地入口分类与状态

本批不新增常规 Ribbon 根或段落 mount，第 1 节常规根 1 / 段落挂载 5 不变。以下是 **本地图片覆盖层** 的分类计数，与上游 fallback 4 根、共享绘图 right-click 3 根 / 14 叶的统计分开。

| 本地类别 | 根 / 叶数量与顺序 | 状态 / 命令 |
| --- | --- | --- |
| 单图浮条 | 6 根：编辑、裁剪、删除、重置、替换、排列 | 既有图片面板 / 内容命令；重置还需要已读 intrinsic 尺寸 |
| 多图浮条 | 3 根：删除选中、排列、对齐与分布；“已选 N 张”是状态文字，不计根 | 单图 edit / crop / replace / reset 隐藏；逐选中目标 edit / view、控件锁、输入会话和 active owner 判断 |
| 单图右键 | 6 根：编辑、裁剪、替换、重置、排列、删除；分隔线不计根 | 同一单图内容命令 / operation |
| 多图右键 | 3 根：排列、对齐与分布、删除选中；分隔线不计根 | 右键已选对象保留所有 IDs，外部对象先改选择；无单图面板叶项 |
| 排列子项 | 4 叶：置顶、上移一层、置底、下移一层 | `doc.command.arrange-images`；不同容器或任一选中对象不可编辑时禁用 |
| 对齐与分布子项 | 8 叶：左、水平居中、右、顶部、垂直居中、底部、水平分布、垂直分布 | `doc.command.align-images`；同页 / 同容器且 geometry 当前；前六项至少两图，后两项至少三图 |
| Canvas 核心手柄 / 键盘 | 四角均匀缩放柄、旋转柄、整体拖动；不是 Ribbon 根 | `doc.command.transform-images`；Shift 旋转 15° snap；方向键 1px / Shift 10px；Delete / Backspace 批删，Esc 先取消手势、再次清对象选区 |

浮条 / 右键都是已有图片上下文呈现扩展；经典 / grid 主 Ribbon 根数不因多选而增加。未运行动态 adapter、插件扩展和溢出装配，不能把上述本地静态计数推广为 Univer 全部浮动入口总数。所有完整 Sheets / Docs / Pro 目标继续保留。

### 6.3 模型、命令、API 与持久化

以下文件相对 `packages/office-studio/src`。它们已落盘，实际运行仍未验收。

| 责任 / 文件 | 续作行为 |
| --- | --- |
| `features/document/images/document-image-batch.ts`（新增） | 派生 `DocumentImageBatchMember` 保存 imageId、当前 image 副本、LayoutImage 和 pageOrigin；1–1000 个唯一 ID、合法尺寸 / 坐标 / 页、同容器 / 同页、陈旧 image 输入验证。对齐采用可见旋转框；分布使用 union 的最小 / 最大端点；整体 transform 保留局部框和资源。inline 转 float 保持当前 paint 尺寸和原 atom 地址，360° / 相同位置不伪造编辑 |
| `features/document/commands/document-image-commands.ts` | 新 align / transform 先逐选中目标检查 view / edit / 控件锁，再编排同一 batch 候选。即使边界图片原地不动，也不能跳过其权限；arrange 检查所有受影响 peer；批删失败不发布先前删除，atom 删除继续协调 UTF-16 范围、资源和控件 |
| `api/document-image-layout.ts`（新增） | `bindDocumentImageLayout(source, read)` 用 source 实例绑定 live layout owner；旧 handle dispose 不删除新 owner。`readDocumentImageLayout` 核对 source / host，优先借用当前布局，缺失时复用同一布局引擎和辅助 inspection 布局；detached 调用允许传 measurer，geometry 不进入快照 |
| `api/document-images.ts`、`api/document-image-commands.ts` | 四个批量 helpers 和两个公开命令；本地 handle 必须属于同一文档实例和 segment。读取 geometry 后重新核对模型实例 / revision；callback 重入后拒绝旧结果；无效 IDs / 参数 / segment 显式拒绝 |
| `components/document/use-document-image-editor.ts` | selectedIds 是对象选择的唯一权威，selectedId 只是最后一项兼容 getter / 单选 setter；内容改动保留仍存在的成员，unit / host / API owner 切换和 deactivation 清选择。多选拒绝单图面板，显式打开单图面板转回单选；绑定 / 释放 live layout provider |
| `components/DocumentWorkbench.vue` | 给 Canvas 传所有 IDs；selection 请求可设置新 IDs，command / operation 必须匹配当前 IDs、source / host / session 并逐目标 editable。geometry 只借出与当前模型匹配的排版，内容修改仍调用同一 execute / 文档历史 |
| `components/document/DocumentCanvasViewport.vue`、`features/document/controllers/document-image-canvas.ts` | Ctrl / Cmd / Shift toggle、普通点击 / 右键保留已选对象；整体拖动 / 四角均匀缩放 / 旋转只改 preview，pointer-up 一次内容命令。模型 / 会话 / 文字输入 / composition / host / zoom / deactivation 取消旧手势；回到起点清 preview；捕获的全部 IDs 和命令目标匹配 |
| `features/document/layout/document-image-layout.ts`、`features/document/rendering/document-scene.ts` | 旋转 stem 保持 24 屏幕像素；每张图轮廓和一个整体框，geometry 不可用时不提供手柄。批量 preview 投影全部图片到同一共享 image 绘制命令，限定物理 page 避免移动重复页眉的其他出现位置；两后端共用场景，无独立 Docs 绘图引擎 |
| `components/document/DocumentImageOverlay.vue` | 三个多图根、八个对齐叶、四个排列叶；全部 Ant Button / Dropdown / Menu / Tooltip。受控 open、autoFocus、Esc 和焦点返回，主题由统一 ConfigProvider / token 提供 |
| `features/document/services/document-clipboard.ts` | `documentImagesClipboardData(model, ids)` 投影全部选中真实 IMG / alt / resource / crop / flip，单图 helper 委托它；多图 copy / cut 不再只取最后一个对象，源模型不改写 |

接口已与主任务协调，`api/facade.ts`、`api/office-studio.ts`、`index.ts`、Enum 和共享 rendering 由主任务接线。已静态回读四个 wrappers；没有类型 / 构建 / 浏览器证据。

```ts
type Images = readonly (FDocumentImage | string)[]
interface DocumentImageGeometryReadOptions { preferredPage?: number; measurer?: LayoutMeasurer }
interface DocumentImageBatchTransform { x?: number; y?: number; scale?: number; angle?: number; origin?: { x: number; y: number } }
FDocument.alignImages(images: Images, align: DocumentImageAlign, segmentId?: string, options?: DocumentImageGeometryReadOptions): boolean
FDocument.transformImages(images: Images, transform: DocumentImageBatchTransform, segmentId?: string, options?: DocumentImageGeometryReadOptions): boolean
FDocument.arrangeImages(images: Images, arrange: DocumentImageArrange | number, segmentId?: string): boolean
FDocument.removeImages(images: Images, segmentId?: string): boolean
```

代码块为公开签名说明，不是可执行整文件。公开 `doc.command.align-doc-images` 使用 `drawingIds / alignType / preferredPage`，`doc.command.transform-doc-images` 使用 `drawingIds / transform / preferredPage`；二者复用既有 `unitId / segmentId / subUnitId / expectedRevision / expectedHistoryGeneration` 守卫。`Enum.DocumentImageAlignType` 对应九个固定字符串值（含 default）；本地 Docs 几何命令没有借用 Sheets 的 `sheet.operation.set-image-align` 注册项。

transform / align / arrange / remove 只改变同一领域模型，通过现有正向 / 逆向候选进入一个文档内容历史；拖动 preview、对象 IDs 和布局不进内容历史。保存使用同一 `body / customBlocks / drawings / drawingsOrder` 编解码链路，保留图片 ID、位置、尺寸、角度、flip、crop、资源和 atom 地址。这里只说明落盘接线，真实保存重开仍无运行证据，Office 二进制不能由 JSON 草稿代替。

### 6.4 草稿、静态修正与验证边界

| 新草稿（相对 `packages/office-studio/tests`） | test 声明数 | 草稿目标 |
| --- | ---: | --- |
| `document-image-batch.contract.test.mjs` | 19 | 六种对齐独立坐标、水平 / 垂直 / 负间距分布、union 末端来自中间宽图、no-op、旋转 / 局部框 / crop / flip、对角固定均匀缩放、24px 旋转柄 / snap、合法大框中心、单图旋转、缩小 inline 转 float、错误 / 陈旧 / 跨页 / 跨容器、权限 / 锁原子性、同段 atom 删除、目标匹配、多图 clipboard、共享 scene 和重复页眉 page preview |
| `document-image-batch-facade.contract.test.mjs` | 20 | 四个真实 FDocument 批 API、单次事件 / 内容历史、undo / redo / save / reopen、两 atom / UTF-16、无效 / no-op、同 ID 不同 Studio、锁 / view / edit / peer、header / note segment、live layout 借用 / owner 替换 / cleanup / mismatch、回调重入版本保护、事件取消、公开命令 / Enum / version / segment 守卫 |
| `document-image-batch.browser.test.mjs` | 18 | Ctrl / Shift toggle、普通 / 右键保留选择、三根 / 八叶 / 两图分布 disabled、八叶实际几何闭环、拖动 / resize / rotate / Esc / 起点 no-op、方向键 / 批删、copy / cut / paste、取消 / 陈旧输入或模型、锁、主题 / focus、合成 composition、KeepAlive、standalone source 替换 |
| `fixtures/document-images.html`（扩展） | — | `?multi` / `?multi-locked` 四张不同位置 / 尺寸的真实 PNG，含 30° 图；groupPoint / imageRectangles / 可取消命令 / API update / transform helper。旧场景保留，夹具未启动 |

本轮只静态回读源文件和草稿，使用 `rg` 读取 test 声明计数；没有执行这些 test。已按现有契约修正：完整块控件锁草稿必须包含合法文字 anchor；重复页眉的 pageContent 使用实际 `cells / columnSeparators` 字段；选区旋转中心允许合法大框的派生坐标，大框草稿使用合法的布局 top-left 并覆盖成员校验和真实内容命令；绝对层级命令拒绝非有限值；浏览器重新选组先用 Esc 清旧 IDs。它们是静态修正，不是通过结果。

本批可替换 UI 全部仍用 `antdv-next`，包括 Canvas 的既有 Input.TextArea 和新增多图浮条 / menu；没有新增产品原生表单控件。已沿用 Context7 查询 Ant Dropdown / Menu 与 Playwright 修饰键 / 鼠标 API，版本以本地依赖为准；没有运行 UI、原生 Windows IME 或系统剪贴板。Facade 草稿的 Vite SSR、浏览器草稿的 Vite / Edge 仅写在测试代码里，不是已启动服务。

计算域仍是有界主线程 TypeScript 几何 / 命令编排和共享 Canvas 场景；这批没有新增 Worker / Rust / WASM 接口，不声称完成重计算迁移、性能优化或零拷贝。两个后端、内存与大样本指标仍须统一验证。

### 6.5 剩余目标

- 非均匀缩放 / 错切、group / ungroup、跨类型和跨容器的完整层级、混合图片 / shape / chart 对齐；本批同容器几何限制是当前开发缺口，不改写最终范围。
- 真正跨页拖动及逻辑 anchor 迁移、复杂分节 / RTL / exact 行高 / 表格和 note 分页、全部图片内容控件 atom 迁移与复杂 SDT。
- 极端重叠且宽度不同的分布：当前等负间距使用原 union 跨度，中间对象可能超出原 union，不能笼统声称所有外缘始终保持；需 Office 可观察样本和具体语义补核实。
- 媒体 UUID / EXIF / DPI、Office 原生图片 / 锚点 / crop / 关系与资源冲突，真实 DOCX 二进制 / 独立重开。
- Canvas2D / WebGPU、真实明暗主题、焦点 / 窄屏 / 设备缩放、系统剪贴板、Windows 原生 IME、动态插件 / Pro / 移动入口的完整运行验收。
- Worker / Rust / WASM 图片布局与批量几何、取消 / 版本 / lease / memory.grow / 复制次数和字节，以及同数据集首屏 / 滚动 / 缩放 / 编辑和大样本前后基准。

中央清单与后端需求由主任务同步。本节和 57 个新草稿不能作为 DOC-07 或完整 Univer / Pro 的验收完成证据。
