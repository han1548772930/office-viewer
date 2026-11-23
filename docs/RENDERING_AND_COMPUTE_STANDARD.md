# Canvas / WebGPU 渲染与 Rust/WASM 计算执行标准

更新日期：2026-11-23  
用户明确要求：编辑器使用 Canvas，优先 WebGPU；可提升性能的重计算在 Rust/WASM 层完成，注意零内存拷贝；使用五个指定的 WASM 库，WASM 可以操作 DOM。

本文件与[总标准](../PROJECT_EXECUTION_STANDARD.md)、[源码规则](UNIVER_REFERENCE_RULES.md)、[功能对齐清单](UNIVER_FEATURE_PARITY.md)共同执行。技术路径不能用来降低 Office 功能、布局、编辑与保真目标。

模块和依赖见 [完整工作台架构](../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)；具体库调用、缓冲所有权与边界复制见 [WASM 平台与内存契约](RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)。

## 1. 渲染路径

- Sheets 网格/行列头/选区/冻结区域/绘图，以及 Docs 页面/文字/表格/对象/选区，以 Canvas 为编辑内容的渲染载体。
- 默认优先 WebGPU：在 canvas 上使用 WebGPU 渲染后端。能力不可用、初始化失败或恢复失败时使用 Canvas 2D 后端；功能与数据模型一致。
- GPU 后端必须真正提交绘图命令、缓冲/纹理与管线。仅检测 navigator.gpu 后继续全部用 2D 绘制不能标记为 WebGPU 渲染完成。
- DOM 用于 Vue/antdv-next 应用壳、Ribbon、面板、弹层、最小文本/IME 输入桥和无障碍语义层。除 Canvas 绘制、命中测试、选区和拖拽等编辑器内核外，所有可替换交互控件均使用 `antdv-next`，包括画布内部或覆盖层中的单元格编辑、公式输入、文档输入、批注、内容控件和浮动操作区。不能逐单元格 DOM、组件库 Table 或整页 contenteditable 承担正文渲染。
- DOM 输入桥优先使用 `antdv-next` 的 Input / Input.TextArea，仅接收输入/组合输入并映射到逻辑范围，不能成为第二份文档模型。只有原生事件或输入法能力确实无法由组件承载时，才保留最小原生桥接，并记录技术依据。可访问性镜像由同一模型生成，不能另存内容。
- 主线程 WASM 通过 `web-sys` 操作 DOM 是允许的实现路径：Canvas 宿主、输入焦点、选区、事件与无障碍桥接由独立平台适配器负责，接收 Vue/`antdv-next` 提供的元素句柄并显式释放监听。解析/公式/布局等纯内核不依赖 DOM；Worker 只调用其环境可用的浏览器 API，不访问 Window/Document。
- 所有 `antdv-next` 控件通过统一 ConfigProvider 和主题 token 管理；Canvas 覆盖层、弹层和输入控件必须验证明暗主题、尺寸、禁用/悬浮/焦点状态及焦点恢复。
- 渲染后端、布局、文本度量、命中、逻辑选择与命令分别分层；Canvas2D 与 WebGPU 复用布局和绘制数据，不维护两套业务规则。

## 2. 后端与资源生命周期

- 用运行时能力探测，处理无 navigator.gpu、无适配器、设备请求失败、上下文创建失败、纹理/缓冲限制及设备丢失；不靠浏览器 UA 或固定版本表判断。
- 设备丢失后暂停旧管线，丢弃旧资源，在受控重试后重建或回退；主动 dispose 不触发无限恢复循环，失败原因可诊断。
- canvas 的上下文类型不能随意切换。切换后端时由 surface 宿主替换 canvas，并从权威模型/布局恢复，不能假设同一已绑定 canvas 可重新取另一类型上下文。
- 统一 devicePixelRatio、编辑器缩放与 CSS 坐标。backing store、光标、选择、图形、鼠标命中与 DOM 输入桥使用明确转换，避免重复缩放和高 DPI 模糊。
- 每次渲染只用当前尺寸的有效纹理，resize 后更新视口/裁剪；容器不可见或零尺寸时避免无效提交。
- 字体/媒体异步加载带版本；不能把旧度量或旧纹理应用到新文档。处理字体回退、组合字、双向文本、中文与 emoji。
- 统一管理字体/字形图集、图片纹理、buffer、pipeline 和缓存；限制缓存容量，淘汰/释放有策略。卸载工作台回收 GPU 与 DOM/Worker 资源。

## 3. 视口和帧优化

- 渲染按可视行列/页面、冻结分区与缓冲区域裁剪；不能把整本工作簿或所有文档页转成一个巨大 canvas。
- 用稀疏模型、行列前缀尺寸索引、可见区域查询、tile/页缓存与脏区失效；一次单元格修改不全量克隆/布局/绘制整个文件。
- GPU 数据按材质/纹理/裁剪批次组织，重复几何尽量复用；避免每帧重新编译 shader、建 pipeline、上传整本数据。
- 静态网格/纸张/文字与活动光标/选区/拖动预览可分层调度，保持坐标一致；同一帧合并重复请求。
- 使用 requestAnimationFrame 调度交互绘制；后台/静止时按需要暂停，不用固定间隔空转。
- 文字 shaping/行断/度量独立于后端；可用离屏 Canvas 生成字形图集再由 WebGPU 合成，但不能每帧整页重新栅格化成一张图片冒充完整 GPU 引擎。
- 首屏、滚动、编辑、缩放、冻结和窗口调整分别测量；只有帧率截图，没有输入/内存/上传指标不算优化验证。
- UI 功能验收默认不要求浏览器截图或像素级视觉比对；以渲染布局、文本度量、交互命中和运行时行为测试为主。只有用户明确要求视觉回归时才执行截图/像素检查，且不能替代功能验证。

## 4. 大计算优先 Rust/WASM

大型文件解析与写出、公式批量/增量重算、透视分组聚合、排序筛选、条件规则求值、图表数据统计/变换、查找索引与复杂布局热点优先评估并实现 Rust/WASM 内核。迁移的是实际算法和批量数据路径，不能仅用 WASM 包裹同一段同步 JavaScript。

- 大任务运行在 Web Worker；主线程负责交互、命令编排和绘制调度，不同步执行阻塞性 WASM。
- 小操作可留在 TypeScript。按基准确定阈值；线程启动、序列化与复制成本也计入，不以“全部改 WASM”代替性能设计。
- 每个计算域指定唯一权威引擎、输入 schema、结果类型和数值/排序规则；不能 Rust 与 TypeScript 各存一套互相冲突的计算逻辑。
- 解析、计算与排版按模块拆分；不把所有引擎塞进 office-parser 的单个函数或 parser.rs。
- 复用 Worker 与 WASM 实例，按 CPU/内存和前台优先级调度，控制并发；不能用硬件线程数无上限开启重型文件任务。
- SIMD、多线程或 GPU compute 是后续可测优化。先探测支持与部署条件，不让 SharedArrayBuffer/cross-origin isolation 成为普通浏览器基本功能的硬前提。
- WebGPU 优先解决绘制；大型业务计算优先遵循 Rust/WASM 决策，不为炫技将需要确定性的公式/聚合强行挪到低精度 GPU 路径。
- WASM 平台层实际使用 `wasm-bindgen-futures` 连接 Promise/Future 与可取消分块任务，`serde-wasm-bindgen` 处理控制 DTO，`web-sys` 处理浏览器/Worker API，`js-sys` 处理 TypedArray、Promise 和内存视图；浏览器生命周期与绑定测试使用 `wasm-bindgen-test`。纯原生内核的单元测试继续用 Rust 测试框架，不把浏览器绑定混入算法层。

## 5. 数据和内存

- JS↔WASM 使用批量 TypedArray/列式或紧凑结构；避免每单元格调用、每帧全量 JSON 或反复序列化整个文档。
- 热路径优先让 Rust 持有批量输入/结果与可复用 arena；同一 WASM 实例内通过 TypedArray 内存视图读写，避免额外 Rust Vec clone 和返回值 glue 再复制。每个视图声明 owner、范围、访问模式和 lease；视图期间不能释放/重分配/重入不兼容任务，`memory.grow` 后刷新视图。
- Worker 消息优先 transferable buffer；明确发送后所有权与 detached 状态。需要复用数据时显式保留权威副本，不能操作已转移 buffer。
- WASM linear-memory buffer 不能作为普通 transferable 转移。跨 Worker 的输出若需要独立 buffer，只复制有效结果区间一次再 transfer；普通 JS 输入进入 WASM memory 的必要复制、serde 对象转换及 GPU 上传分别记录次数/字节，不宣称整个链路零拷贝。可共享内存路径有单独 capability、同步与部署协议，基础功能不依赖它。
- 输入、计算缓存和输出设置内存预算；预估大小、分块/流式处理或给出明确错误，不能无限扩容后浏览器崩溃。
- WASM memory.grow 后旧 view 可能失效；适配层重新建立视图，不能缓存过期指针或越界访问。
- 避免隐式复制、无界中间向量与大范围快照；维护输入/输出复用策略和资源释放。
- 取消或失败不得返回部分结果假装成功；可以有显式分块结果协议，但应用到文档必须符合事务/版本要求。
- 排序稳定性、空值/错误值/类型比较、locale、日期系统、浮点误差与 Excel 语义要有固定测试；优化不能改变结果。

## 6. 任务协议和取消

每个任务至少标识 requestId、documentId、baseRevision、kind、输入格式/大小和取消状态；结果带匹配版本、统计与诊断。

- 进度只来自真实阶段/已处理量，不用定时器伪造完成百分比；高频事件合并限流。
- 文档关闭/替换、相关内容版本变化或新任务取代旧任务后，旧结果不得覆盖新内容。
- 同步 WASM 执行期间 Worker 不能及时处理普通 postMessage。因此取消必须采用分块并让出事件循环，或终止并重建 Worker 等实际策略，不能只发送 cancel 后宣称已取消。
- 超时、用户取消、Worker 崩溃、WASM trap、数据错误和资源不足分别编码；可重试任务不重复应用。
- 长时间计算不妨碍用户操作其他文档；前台输入/视口相关任务优先于后台索引和预热。

## 7. 性能验收

每次性能改动保存同一数据集的前后对照：设备、浏览器/后端、数据稀疏度、文件大小、首次/热运行、首屏、任务总耗时、跨边界传输、P50/P95 交互与帧耗时、主线程长任务、JS/WASM/GPU 内存。

- 至少覆盖普通文件、大表/多表、密集数值/公式、复杂样式/图形，以及长文档和跨页表格。
- 默认以 60Hz 环境的平滑滚动为目标；记录实测帧时间与掉帧，不凭机器性能承诺所有文件恒定帧率。
- 对可能超过 50ms 的主线程任务检查拆分或迁移；长计算有进度与可验证的取消路径。
- 比较 TypeScript、Worker/WASM 以及数据传输后的端到端表现；WASM 内核快而页面总耗时更慢不能记为优化成功。
- release 参数、SIMD 和线程策略用真实基准决定；体积优先的编译参数与算力优先的计算模块可以不同。
- WebGPU 与 Canvas2D 使用相同输入验证几何/布局/选区结果；测试无 GPU、初始化失败、设备丢失与 resize 后仍能编辑、撤销及保存。
- 性能和保真同时过关；减少数据、降低字体/公式精度、裁掉功能或漏画页面不能作为达标手段。

## 8. 当前迁移边界

2026-11-20 续作检查：代码现位于 `packages/office-studio/src`。Sheets/Docs 主路径已分别接入 `SpreadsheetCanvas.vue` 和 `DocumentCanvasViewport.vue`；布局、场景与 Canvas2D/WebGPU 后端分离，正文输入通过最小输入桥提交命令。不能再把现状概括为 DOM 网格或整页 contenteditable 草稿，也不能据此宣称全部绘制、复杂排版和大文件目标已完成。

`services/wasmAdapter.ts` 已通过 `parser/parser.worker.ts` 调用解析 WASM，使用 transferable ArrayBuffer，并在取消/超时/替换任务时终止 Worker；计算和写出拥有各自 Worker。各协议仍须分别验收版本、内存、过期结果及真实大样本行为；已有 Worker 或测试不等于性能优化达标。基准执行方式及验收口径见 [COMMON-11 基准记录](feature-evidence/COMMON-11-compute-benchmark.md)。

## 9. 技术依据

WebGPU 的适配器、设备、Canvas 配置、设备丢失和错误处理以 [W3C WebGPU 规范](https://gpuweb.github.io/gpuweb/)与 [GPUWeb 官方错误处理设计](https://github.com/gpuweb/gpuweb/blob/main/design/ErrorHandling.md)为依据。规范说明与本项目的后端选择/重试/回退策略分别记录，不能把示例中的 TODO 当成生产恢复逻辑。

<a id="wasm-platform-contract"></a>

## 10. WASM 平台与内存契约

更新日期：2026-11-23。用户要求使用 `wasm-bindgen-futures`、`serde-wasm-bindgen`、`web-sys`、`js-sys`、`wasm-bindgen-test`；能改善性能的重操作在 WASM 层实现，注意零内存拷贝；WASM 可以操作 DOM。

配套：[完整架构](../PROJECT_EXECUTION_STANDARD.md#workbench-architecture)、[渲染/计算标准](RENDERING_AND_COMPUTE_STANDARD.md)、[功能清单](UNIVER_FEATURE_PARITY.md)。以下是可执行契约，性能达标必须有实测，接口存在不等于验收。

### 10.1 五个库的实际职责

| 库 | 使用位置 | 必须验证的实际行为 |
| --- | --- | --- |
| `wasm-bindgen-futures` | WASM 平台 async、Promise/Future、分块任务宏任务让出 | Promise rejection 传播、让出后接收取消、释放任务；微任务不冒称普通消息取消 |
| `serde-wasm-bindgen` | 版本化控制 DTO、错误/诊断、类型化参数和结构化解析结果 | 字段/枚举/未知值校验、Map/object、null/undefined、安全整数；不称对象转换零拷贝 |
| `web-sys` | Worker/Window timer、浏览器能力、DOM/Canvas/焦点/输入/事件桥 | 主线程与 Worker 环境分开、句柄/监听/timer 生命周期、IME/焦点与主题宿主 |
| `js-sys` | TypedArray、ArrayBuffer/Memory view、Promise、Reflect 等平台对象 | view 与 copy 区别、byteOffset/length、detached/memory.grow、transfer 的所有权 |
| `wasm-bindgen-test` | 真实 wasm32 浏览器绑定/异步/DOM/TypedArray 契约测试 | 真实 JS↔WASM 行为和释放；不能用原生 Rust 单元测试替代 |

绑定和通用平台行为集中在 `crates/office-wasm-bridge`。解析、写出、计算通过依赖复用 serializer、宏任务让出和内存边界工具；不要求每个纯领域 crate 直接依赖全部浏览器库。`wasm-bindgen-test` 作为 dev-dependency，不进入生产运行依赖。版本由 Cargo 正常解析并生成锁文件，不能手改锁或把文档最新版本当成本地已安装版本。

### 10.2 主线程 WASM 和 DOM

- 允许主线程 WASM 使用 `web-sys` 操作 DOM/Canvas、焦点、input selection、事件、观测和无障碍桥；不规定所有 DOM 操作必须由 JavaScript 执行。
- Vue/`antdv-next` 管理可替换控件树、Form 和统一主题；平台桥接收其元素/宿主句柄，不另外创建一套按钮、Select 或文本编辑控件。
- 解析/公式/透视/纯布局内核只接收数据，不访问 DOM；Worker 层不读取 Window/Document，浏览器方法通过 capability 和目标环境区分。
- 输入桥映射 `beforeinput/input/composition` 与逻辑会话，不将 input.value 或 DOM 选区变成第二份持久模型。恢复焦点/选区确认 unit、session generation、model revision 和 IME 状态。
- 监听器/Closure/timer 由显式 owner 保存并在 dispose/Drop 移除，不用无界 `Closure::forget`。节点重建或编辑器卸载后旧句柄不能继续操作。

### 10.3 控制 DTO 与 bulk 数据

控制消息包含稳定协议、ID、版本、任务种类、参数、权限上下文、预算、诊断和计量；bulk 值/状态/索引/媒体单独走 TypedArray/buffer/arena。避免每单元格调用、每帧全量 JSON、以字符串承载大数值数组。

`serde-wasm-bindgen` DTO 显式选择 maps-as-objects 与 missing/null 语义，以匹配现有 parser/writer 契约；需要 BigInt 的字段单独版本化，不能把 unsafe integer 静默转 JS number。解析 DTO 仍经过 TypeScript schema 校验/归一化，结构化对象不是可信领域模型。

按用户要求直接重写 Rust 导出、Worker、客户端和类型，全部使用协议 2；不保留旧 WASM 产物兼容，不提供旧 API 或 JSON 字符串结果的回退。加载时校验唯一协议和所需接口，旧产物或缺失新接口返回明确版本/构建错误。测试后置阶段不生成 WASM；统一验证时先生成匹配的新产物。writer 的持久化快照 JSON 是当前输入格式，与旧 WASM 接口兼容无关，JSON/UTF-8 转换仍须计量。错误保留稳定 code/message/details，并区分解析、格式、预算、取消、过期、Worker trap 和平台失败。

### 10.4 可实现的零拷贝边界

| 数据路径 | 契约 | 必须计量的复制 |
| --- | --- | --- |
| 同一 WASM 内核的 buffer/arena | 借用 slice 或移动 Vec 所有权，保留批量缓存 | 业务自身需要的重排/中间结果单列 |
| WASM 内存 → 同线程 JS 临时读取 | TypedArray view，owner 保活；不 await、不重入会分配的 WASM | view 本身 0 字节复制；需要独立持久结果时另算 |
| 普通 JS TypedArray → WASM | 一次写入已分配 linear memory，或绑定一次边界复制 | 字节数与次数；传 Vec/&slice 的绑定不自动零拷贝 |
| Worker → 另一个 Worker/主线程 | 独立 ArrayBuffer transferable，发送后 detached | transfer 不复制；从 WASM view 生成独立区间通常复制一次 |
| WASM → GPU | 有效区间批量上传，复用 GPU buffer/texture | GPU upload 的字节与次数；不是零拷贝承诺 |
| serde DTO / JS 对象 / JSON | 结构化控制信息或持久化快照 | 记录对象转换/序列化；不可冒称零拷贝 |
| 共享内存 | 独立能力/部署/同步协议，保留普通浏览器路径 | 共享区无需 transfer；线程同步、拷入和 GPU 上传仍单列 |

禁止 transfer 整个 WASM `WebAssembly.Memory.buffer`、把 raw pointer 跨实例使用或将 view 指向的已释放内存当作独立结果。不同 WASM 实例的地址空间互相独立，两个实例共同使用 js-sys 并不共享 linear memory。

优先减少多余 Rust Vec clone、绑定 glue copy、全量 JSON/中间数组以及重复 Worker 传输。一个结果临时 view → 独立有效区间 copy → transfer 是明确的一次出站复制，不叫端到端零拷贝。不能通过丢弃源信息、减少功能或修改数值/布局语义降低复制量。

### 10.5 Buffer lease 生命周期

每个 lease 明确 `ownerId/generation/elementType/byteOffset/length/access/memoryEpoch`，含空 view、偏移与上界校验。实际 API 可通过受控同步 wrapper 表达等价信息，不要求所有字段都进入持久模型。

1. 分配/预留所有输入和输出所需空间后再建 view；预算和 `try_reserve` 失败给出稳定错误。
2. view 只在 owner 存活、地址范围和 generation 有效时可读/写；Rust 借用和 JS 访问不重叠为非法别名。
3. view 有效期间不 free/reallocate/consume owner，不进入会 memory.grow 的 WASM，也不跨 await、异步回调或任意 JS 重入保留 view。
4. 任意可能 grow 的调用后检查 buffer identity/epoch 并重新建 view；memory.grow(0) 也可能使普通 memory 旧 buffer detached。SharedArrayBuffer 旧 view 也不自动扩展长度。
5. 需要跨线程或长期持有结果时同步复制有效区间到独立 ArrayBuffer，再释放 owner；元数据、warning 等可能分配的调用放在获取 view 之前。
6. 错误/取消/过期/卸载路径释放 owner、任务、监听和缓存；重复释放/陈旧 handle/超界访问有显式拒绝，不凭调用习惯保证安全。

unsafe TypedArray::view 只出现在可审阅的边界适配中，写明所有者、有效期、重入和 growth 依据。外部 JS 总能误用 view，公开 API 必须有安全默认（独立 copy/同步 wrapper）并明确低层 lease 契约；不要对任意 JS 调用声称由 Rust borrow checker 全面保障。

### 10.6 异步、任务和取消

每域任务保存 `requestId/unitId/baseRevision/kind`，调度适配对应现有 `documentId/modelRevision` 等字段；结果匹配上下文后原子应用。进度来自真实阶段/工作量；CPU/内存和前台优先级控制并发，实例复用有生命周期。

`wasm-bindgen-futures::JsFuture` 用于真实 Promise，`spawn_local` 仅在 owner 可取消/可释放时使用。分块算法在浏览器/Worker 宏任务 timer 或等价调度后继续；单纯 `Promise.resolve().await` 只让微任务，不能使普通 cancel 消息及时处理。同步 parser/writer 可使用 terminate/recreate，不能给同步函数套 async 后宣称已可协作取消。

取消/超时/关闭/替换或 revision 改变后拒绝旧结果，失败不发布部分修改。不同域的预算/资源峰值和复制次数要纳入同一任务指标；WASM 内核快而总传输/页面更慢不算性能优化。

### 10.7 验证和当前阶段

原生 Rust 测试验证纯模型、排序/过滤/公式/布局和错误语义；`wasm-bindgen-test` 验证真实 TypedArray view/copy、memory.grow、结构化 DTO、Promise 让出/拒绝及 DOM 焦点/选区/释放。TypeScript 契约测试验证协议 2、拒绝不匹配产物、独立 transferable buffer、错误与过期结果；浏览器测试验证 IME/编辑/历史/重开、两种渲染后端和用户任务。

用户当前要求“先实现，后统一测试”，因此本阶段可以编写测试草稿和基准夹具，但不执行测试、类型检查、构建、Rust/WASM 检查、基准或重新启动 dev/浏览器。依赖锁文件正常生成属于依赖更新，不是功能/性能验证。统一验证时先生成匹配的 WASM，再执行必要检查和同数据集前后对照。

记录设备/浏览器/后端、版本、普通/大样本规模、冷热运行、首屏/滚动/编辑/缩放/任务、P50/P95、主线程长任务、JS/WASM/GPU 内存和所有复制/传输字节。未实测保持“未验收”，不说已优化或全链路零拷贝。

技术依据：官方 [wasm-bindgen Promise/Future](https://rustwasm.github.io/docs/wasm-bindgen/reference/js-promises-and-rust-futures.html)、[TypedArray view 示例与 growth 限制](https://rustwasm.github.io/docs/wasm-bindgen/examples/webgl.html)、[Closure 生命周期](https://rustwasm.github.io/docs/wasm-bindgen/reference/passing-rust-closures-to-js.html)、[serde-wasm-bindgen serializer](https://docs.rs/serde-wasm-bindgen/latest/serde_wasm_bindgen/struct.Serializer.html) 与 [wasm-bindgen-test](https://rustwasm.github.io/docs/wasm-bindgen/wasm-bindgen-test/index.html)。查阅日期 2026-11-23；固定本地版本以 Cargo.lock/源码为准。

### 10.8 主线程计算与 DOM 操作的边界

主线程 WASM 可以通过 web-sys 完成有界计算并操作 DOM。WASM 并不会使同步计算脱离浏览器主线程；长时间解析、布局或统计仍会阻塞输入和绘制。DOM 操作由独立平台适配层接收现有宿主句柄，纯领域与计算内核保持不依赖 DOM。

大型计算运行在 Worker/Rust/WASM；主线程 WASM 可负责 Canvas 宿主、焦点、文本选区、输入会话及短时命中。Vue/antdv-next 管理可替换 UI、表单和主题，Worker 不访问 Window/Document。异步让出、版本检查、句柄释放和输入法状态共同约束跨边界调用。

是否迁移小型交互由同一数据集的端到端基准决定，记录 WASM 调用、DOM 更新、P95 输入延迟和主线程长任务。存在 WASM 或 DOM 接口不代表已优化；没有真实测量时保留未验收。
