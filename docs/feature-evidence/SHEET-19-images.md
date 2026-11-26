# SHEET-19 图片、仿射几何与剪贴板证据

记录日期：2026-11-26。状态：**未验收**。本次补齐图片剪贴板、结构/尺寸投影、取消与作者数据保真；不将这些改动或测试声明数计为 SHEET-19 全部能力完成。文字/图片水印、打印、真实 Office 写回以及所有 Pro 能力仍按[完整功能清单](../UNIVER_FEATURE_PARITY.md)验收。

本阶段遵循“先实现，后统一测试”：只读取固定源码和现有实现、编辑代码及测试草稿。没有执行测试、类型检查、构建、Rust/WASM 检查、浏览器、基准或 `git diff --check`，也没有重启 dev。

## 1. 参考版本与模块归属

- 固定仓库：`E:/rust_test/univer`，v1.6.0，commit `eaf4cfa04fabdbcfc922c692f3450459f8a29546`；本次以 `git rev-parse HEAD` 读取确认提交。
- `features/sheet-images` 负责模型校验、显示框/四角范围、锚点、裁剪、资源和剪贴板计划；`api` 编排版本、权限及既有内容事务；Vue Canvas 管理输入与手势，控件使用 `antdv-next`。
- 图片与其他绘图共享 `domain/drawing-affine.ts` 的完整矩阵契约；本次没有另建几何引擎。小型位置计算保留 TypeScript；没有新增计算 Worker 或 WASM 协议，也没有性能优化达标声明。
- 文件解析/写出复用现有 parser/writer。系统 PNG 编码使用浏览器平台 API，不能当作 Office 导出或 WASM 计算完成证据。

固定源码主要来源（以下路径均相对上述固定仓库）：

| 来源 | 核对的契约 |
| --- | --- |
| `packages/sheets-drawing-ui/src/menu/schema.ts`、`image.menu.ts` | 图片/背景 Ribbon 的顺序、grid 位置、子项、工作表类型和权限状态 |
| `packages/sheets-drawing-ui/src/menu/worksheet-background-image.menu.ts`、`save-images.menu.ts` | 背景存在性、删除状态、单元格图片保存的宿主能力与选区条件 |
| `packages/sheets-drawing-ui/src/menu/drawing-popup-menu.controller.ts` | 单图六项、显隐条件、扩展位置、命令参数与弹层生命周期 |
| `packages/drawing-ui/src/views/image-popup-menu/ImagePopupMenu.tsx` | Sheets 通过一个下拉入口呈现上述菜单；More 是呈现状态，不另算六个功能 |
| `packages/drawing-ui/src/menu/{schema,group.menu,arrange.menu,align.menu}.ts` | 绘图右键三组、组/层级/对齐子项及选区条件 |
| `packages/sheets-drawing-ui/src/controllers/shortcuts/drawing.shortcut.ts` | 四方向移动、删除；公式栏、面板和形状文本编辑时不触发绘图快捷键 |
| `packages/sheets-drawing-ui/src/controllers/sheet-drawing-copy-paste.controller.ts` | 对象即时 cut、区域 copy 排除 None、区域包含判断、粘贴位置和新身份 |
| `packages/sheets-drawing/src/basics/drawing-reference.ts` 及其 `__tests__/drawing-reference.spec.ts` | 插删标记、删除区间坍缩、末单元格裁切、移动/复制语义 |
| `packages/sheets-drawing-ui/src/views/sheet-image-panel/{SheetDrawingPanel,SheetDrawingAnchor,SheetImageTransparency}.tsx`、`sheet-drawing-panel.util.ts` | 图片属性、三种锚点与不透明度的作者数据契约 |
| `packages/drawing-ui/src/utils/clipboard-image.ts` 及其测试 | 原生图片 MIME、PNG 转换、异常/平台回退与资源生命周期 |
| `packages/sheets-drawing/src/commands/commands/{set-sheet-drawing,remove-sheet-drawing}.command.ts` | 批次、命令/mutation、正反向历史及保存模型 |
| `packages/sheets-drawing/src/facade/__tests__/f-over-grid-image.spec.ts` | Facade 的编辑、几何与作者快照测试设计 |

公开入口：[Sheets 图片](https://docs.univer.ai/zh-CN/guides/sheets/features/images)、[Sheets 水印](https://docs.univer.ai/zh-CN/guides/sheets/features/watermark)。本记录的已核对证据来自固定源码；公开页面演示、没有公开源码的 Pro 行为及 Microsoft Office 文件样本仍需独立核实，不将链接列出等同于已经验证。

## 2. 入口清点

范围为 desktop 图片相关 OSS 声明。classic 顺序按 `order`，grid 与 classic 是同一集合的布局变体；未运行动态装配，不把静态声明当作某一窗口实际可见数量。全工作台和 Pro 的清点见[入口总表](TOP-RIBBON-PARITY-CHECKLIST-2026-11-21.md)和[工具栏对齐记录](../UNIVER_TOOLBAR_PARITY.md)。

| 分类 | 数量 | 计数边界 |
| --- | ---: | --- |
| 常驻 Ribbon 根 | 2 | 图片、工作表背景图 |
| 上述根的菜单子项 | 4 | 浮动图、单元格图、添加背景、删除背景 |
| 单元格/行头/列头上下文声明 | 3 | 同一保存命令在三处注册；只有 1 个命令语义 |
| 共用 DRAWING 右键根 | 3 | 组合、层级、对齐；不是常驻 Ribbon |
| 共用 DRAWING 右键子项 | 14 | 组合 2、层级 4、对齐/分布 8 |
| 单图浮动菜单固定项 | 6 | 编辑、删除、裁剪、水平翻转、垂直翻转、重置 |
| 独立 More 功能 | 0 | 复用同一浮动菜单，图标显隐不增加功能数量 |
| 绘图快捷键声明 | 5 | 4 个方向与 1 个删除声明（mac 使用 Backspace） |
| `FloatingObjectToolbarPosition.SHEET` 扩展 | 未核实完整数 | controller 动态拼接扩展 `BUTTON`；不得以固定六项覆盖插件入口 |
| Pro 图片/水印/扩展声明 | 未核实完整数 | 保持全范围目标，不推导为零或从交付中移除 |

### Ribbon 与单元格上下文

| 标签/入口 | 分组、位置和顺序 | 上游命令 ID | 显隐/禁用条件 | 来源 |
| --- | --- | --- | --- | --- |
| 图片 | INSERT / MEDIA；order 0；grid row 1 / column 1 / rowSpan 2 / showLabel | `sheet.menu.image`（菜单根） | Sheets 宿主；普通表检查工作簿/工作表/区域编辑权；chart sheet 浮动图检查工作簿/工作表编辑权 | `schema.ts`、`image.menu.ts` |
| 插入浮动图片 | 图片子菜单 order 0 | `sheet.command.insert-float-image` | 同图片根；chart sheet 允许浮动图 | 同上 |
| 插入单元格图片 | 图片子菜单 order 1 | `sheet.command.insert-cell-image` | Sheets 宿主；chart sheet 禁用；普通表检查选区编辑权 | 同上 |
| 工作表背景图 | VIEW / DISPLAY；order 0.5；grid row 1 / column 2 / rowSpan 2 / showLabel | `sheet.menu.worksheet-background-image`（菜单根） | Sheets 宿主、工作簿/工作表编辑权 | `schema.ts`、`worksheet-background-image.menu.ts` |
| 添加背景图 | 背景子菜单 order 0 | `sheet.command.add-worksheet-background-image` | 背景图编辑权限 | 同上 |
| 删除背景图 | 背景子菜单 order 1 | `sheet.command.delete-worksheet-background-image` | 背景图编辑权限，且背景存在 | 同上 |
| 保存选区单元格图片 | MAIN_AREA / COL_HEADER / ROW_HEADER 的 OTHERS；各 order 10 | `sheet.command.save-cell-images` | 上游需要 `showDirectoryPicker`，且选区含 cell image；本项目保存平台另有下载路径，差异需验收 | `schema.ts`、`save-images.menu.ts` |

### 绘图与单图上下文

以下右键位于 `DRAWING / OTHERS`，没有常驻 Ribbon grid 坐标。

| 入口 | 根/子项顺序 | 上游命令 ID | 状态 |
| --- | --- | --- | --- |
| 组合 / 组合、取消组合 | 根 order 1；子 0、1 | `drawing.operation.set-drawing-group`、`drawing.operation.cancel-drawing-group` | 根要求可组合绘图；组合需至少两个；取消组合需存在组 |
| 层级 / 置顶、前移、置底、后移 | 根 order 2；子 0、1、2、3 | `drawing.operation.set-drawing-arrange-front`、`...-forward`、`...-back`、`...-backward` | 无绘图焦点时隐藏 |
| 对齐 / 左、水平居中、右、上、垂直居中、下、水平分布、垂直分布 | 根 order 3；子 0–7 | `sheet.operation.set-drawing-align-left`、`...-center`、`...-right`、`...-top`、`...-middle`、`...-bottom`、`...-horizon`、`...-vertical` | 少于两个绘图时隐藏；命令模型再检查可执行性 |

单图 popup 固定数组依次如下。其 index 为分组索引，三个 index 2 项仍按数组顺序呈现；没有推导新的排序。

| 标签 | index | 上游命令 ID | 本项目对应 |
| --- | ---: | --- | --- |
| 编辑 | 0 | `sheet.operation.edit-sheet-image` | `sheet.operation.open-sheet-image-editor`；`SheetImageEditorDialog` → draft/patch → 统一图片更新 |
| 删除 | 1 | `sheet.command.remove-sheet-image` | `sheet.command.remove-sheet-images`；整批删除、单一历史事务 |
| 裁剪 | 2 | `sheet.operation.open-image-crop` | 同 ID；Canvas crop session → `sheet.command.set-sheet-image` 提交 |
| 水平翻转 | 2 | `sheet.command.toggle-flip-drawings`，`flipH: true` | 图片更新 patch 的 `flipX`；作者 crop 保留，渲染使用独立 flip |
| 垂直翻转 | 2 | 同命令，`flipV: true` | 图片更新 patch 的 `flipY` |
| 重置 | 3 | `sheet.operation.image-reset-size` | 现有 reset/update 路径；本次未新增 reset 实现或浏览器验收 |

popup 仅在单一图片/可支持绘图对象选择时出现；裁剪对象活动、多选、无对象、shape/OLE、配置 `disablePopup` 等情形隐藏；DOM 绘图禁用图片特有编辑/裁剪/翻转/重置。后续扩展来自 `FloatingObjectToolbarPosition.SHEET`。本项目的显式多图工具区、对齐与组合由共享 drawing 模块维护，本次不改该模块的 UI/命令。

## 3. 本次真实实现

| 变更 | 实际行为与数据闭环 | 限制/待验证 |
| --- | --- | --- |
| affine clipboard | 保存完整局部宽高及 `a/b/c/d/e/f`；显示框与四角 AABB 分别校验，支持 shear/reflection。粘贴仅平移 e/f 和兼容锚点，保持局部几何、原始 size/crop、flip、媒体与未知 metadata | 真实浏览器、重开和文件写回未执行 |
| 共享绘图几何更新 | `sheetImageWithDrawingGeometry` 在局部宽高与 a/b/c/d 不变时保留作者 size/crop/rotation，供剪贴板和 native owner 共用；geometry patch 比较原几何，真正的尺寸/方向更新仍走完整 bake | 共享宿主对齐/组变换及显式 geometry resize 的集成/浏览器验收未执行 |
| HTML 作者格式 | `data-office-image-clipboard-model` 携带版本化作者描述；`IMG` 用索引携带实际 source，描述中省略 source，避免重复 base64。新 Studio 可以在旧 broker 释放后恢复；旧完整-source 描述仍可读取 | 支持受控内部格式；普通外部 IMG 仍走外部图片准备；SVG 原生剪贴板与其他上游格式未在此补齐 |
| 验证和预算 | 版本、token、来源、重复图片 ID、媒体索引、矩阵、显示框、AABB、真实 data URL 尺寸/MIME/字节均校验；1000 张、单图 5 MiB、批次 32 MiB、序列化 64 MiB 预算 | 预算是同步安全边界，不是性能阈值或大样本达标证据 |
| range/object copy | 区域 copy 排除 None；检查完整 footprint 和可见行投影；坍缩 Both 在区域 copy 跳过，对象 copy 拒绝。显式对象 copy 支持三种锚点 | 区域 copy 与对象 copy 不共享 cut 时机，沿用上游可观察语义 |
| range/object paste | all / exceptBorders 带图片，其余特殊粘贴不带浮动图。range cut 仅在候选模型移走原 ID；对象 cut 立即形成删除事务，之后 paste 新建 ID。图片与单元格共同进入现有候选及历史 | 真实多区域/过滤/冻结/跨工作簿浏览器闭环待验收 |
| API 版本与权限 | copy/paste 捕获用户、工作簿实例、revision、history generation；校验调用方 expected 值并在发布前再次拒绝陈旧目标。来源权限按当前实际 footprint 或捕获 footprint，目标由图片事务检查完整 footprint | 权限与历史重入有草稿，未执行 |
| axis/cell/dimension 投影 | Position 按前后 marker 像素差平移矩阵；Both 按两 marker 投影更新局部宽高与中心，保留矩阵方向、size/crop 作者基准；None 固定。插删、移动、复制、cell shift 共用此逻辑 | 父任务在最终 filter/auto-height 后接 geometry-change，已被结构/显式绘图命令修改 placement 的对象跳过，避免重复投影；仍须统一执行集成验证 |
| 隐藏恢复 | Both marker 坍缩时返回零显示尺寸/footprint，保留可逆作者矩阵；取消隐藏后通过 geometry-change 恢复，避免写入 singular matrix | 真实冻结窗格与两个渲染后端尚未验证 |
| ratio crop | 中心平移使用完整 affine 世界方向及 source flip；保留完整未裁剪源框。属性 draft 已复用服务路径，未另建 rotation-only 算法 | 隐藏零尺寸拒绝；属性 resize/连续比例裁剪有草稿 |
| native PNG | 规范化资源；HTTP 使用实际下载媒体而非 URL/MIME 提示；解码和编码前后检查 `signal/current`；finally 清空临时 Canvas bitmap 与 decoded image source | 主任务负责写剪贴板及陈旧 fallback 接线；mock 平台测试不证明真实栅格结果或原生剪贴板 |

命令路径复用 `SHEET_IMAGE_COMMANDS`、`SHEET_IMAGE_OBJECT_COMMANDS`、既有 workbook 内容历史及 parser/writer；不新增组件内内容撤销栈。图片作者、媒体 resource、Canvas 布局和绘制命令保持分离。

## 4. 修改文件与接口

本子任务生产代码（相对 `packages/office-studio/src`）：

- `features/sheet-images/sheet-image-clipboard.ts`
- `features/sheet-images/sheet-image-object-clipboard.ts`
- `features/sheet-images/sheet-image-system-clipboard.ts`
- `features/sheet-images/sheet-image-structure.ts`
- `features/sheet-images/sheet-image-placement.ts`
- `features/sheet-images/sheet-image-crop.ts`
- `api/sheet-image-object-clipboard.ts`

新增/扩展接口：

```ts
captureSheetImageClipboardEntry(workbook, sheet, image): SheetImageClipboardEntry
translateSheetImageClipboardEntry(entry, axes, dx, dy, id): SheetFloatingImage
sheetImageObjectClipboardFromHtml(html: string): SheetImageObjectClipboard | undefined
transformSheetImagesForGeometryChange(images, { before, after }): SheetFloatingImage[]

interface SheetImageClipboardEncodingOptions {
  signal?: AbortSignal
  current?: () => boolean
}
sheetImageClipboardPngBlob(resource, options = {}): Promise<Blob>
```

API/UI HTML reader、系统 PNG 的 current/signal、最终 geometry-change 的事务接线由主任务完成；本子任务未覆盖中央 API、Facade、ExcelWorkbench、Canvas、共享 affine、drawing groups 或 Docs 文件。

## 5. 测试草稿、UI 与剩余验收

新增测试声明共 **66** 项，仅用 `rg` 静态清点，全部未执行。数量用于复查文件，不作为覆盖/完成证明。

| 文件（相对 `packages/office-studio/tests`） | 声明数 | 验证设计 |
| --- | ---: | --- |
| `sheet-image-affine-clipboard.contract.test.mjs` | 23 | 完整 affine、跨表平移、geometry 改变后的 Both 作者基准、HTML 恢复/实体/身份/预算、range copy/paste/cut、特殊粘贴 |
| `sheet-image-affine-structure-crop.contract.test.mjs` | 17 | 插删移动复制、cell shift、尺寸/隐藏恢复、共用 host 纯平移与真实 geometry resize、shear/reflection 裁剪、属性 resize/连续比例、对边固定、旋转、对齐、多图平移 |
| `sheet-image-system-png.contract.test.mjs` | 8 | 真实 1×1 PNG bytes、MIME/尺寸拒绝、取消/陈旧、Offscreen/DOM Canvas 释放、HTTP 实际媒体；编码结果使用平台 mock |
| `sheet-image-facade-history.contract.test.mjs` | 11 | 真实 Studio/Facade setter、整批原子拒绝、陈旧目标、即时 cut、broker 释放、footprint 权限、历史重入、内部/public snapshot 重开 |
| `sheet-image-workbench.browser.test.mjs` | 7 | 真实 ExcelWorkbench/Canvas、Ant 属性/Upload、reflected handle、HTML 跨 Studio paste、明暗主题/焦点、换表取消、文本剪切作用域 |

夹具：`fixtures/sheet-image-fixture.mjs`、`fixtures/sheet-images.html`。浏览器草稿等待 Canvas 实际 viewport 建立后再交互，主题断言等待弹层样式变化；没有截图或像素比较。

本次新增 UI 仅是浏览器夹具，使用统一 `ConfigProvider`、`Form`/`FormItem` 与 `Input`；真实图片属性/上传控件静态回读确认使用 `antdv-next` 的 Modal、Form、Input/TextArea、InputNumber、Select、Checkbox、Slider、Upload 和 Button。Canvas 的命中、选区与拖拽仍由编辑器内核处理。没有新增原生文本/IME 桥；尚未执行主题切换、焦点恢复或真实 Windows IME。

仍未验收：

1. 上述全部单元/契约/浏览器草稿、必要类型/前端/Rust/WASM 检查及本次跨模块 geometry 接线。
2. 真实 Office 图片样本（剪切裁剪、signed outcrop、反射/剪切矩阵、三种锚点、单元格图/背景图）导出和独立重开；图片关系、未知 metadata 与媒体保真。
3. WebGPU / Canvas2D 同输入几何、冻结/过滤/滚动/缩放、设备丢失与无 GPU 回退、原生系统图片/HTML 剪贴板、中文 IME。
4. 全部图片/水印/Pro/插件扩展入口到命令、模型、权限、撤销重做和保存重开的完整闭环；本次记录不能替代全工作台清点。
5. 同数据集普通/大样本的首屏、滚动、编辑、缩放、任务耗时、P50/P95、长任务、JS/WASM/GPU 内存、跨边界传输与复制字节。没有实测数据，保持未验收，不声称优化、零拷贝或性能达标。

HTML JSON/URI 编码、base64、媒体解码与 PNG 编码均存在分配/复制。若后续迁移批量媒体/几何计算，必须遵循[WASM 平台与内存契约](../RENDERING_AND_COMPUTE_STANDARD.md#wasm-platform-contract)，单列必要复制并验证取消/lease/释放，不转移 WASM 整个 linear-memory buffer。
