# Univer v1.6.0 · 桌面 Ribbon 源码批量清点

复算：`node --experimental-strip-types scripts/report-univer-ribbon-batch.mjs`。上游提交 `eaf4cfa04fabdbcfc922c692f3450459f8a29546`。

常规表按标签、再按行、组和列排列。格位是将列出的 OSS 注册项全部纳入、未求值 hidden$ 时，用固定版 `ribbon-grid-layout.ts` 得到的布局；不代表某次会话的可见按钮。Classic 列是组内 schema order，不是横向像素坐标。More 只重排已有项，右键、浮动条和子菜单不重复计入。

## sheets

常规根 63；上下文根 0。含可选扩展，不等于示例默认挂载数。

### START (30)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | HISTORY / 1 (1×1) | 0 | `UndoCommand.id` / `UndoMenuItemFactory` | `packages/ui/src/menu/schema.ts:40` |
| 1 | HISTORY / 2 (1×1) | 2 | `ActivateFormatPainterOperation.id` / `FormatPainterMenuItemFactory` | `packages/ui/src/menu/schema.ts:30` |
| 1 | HISTORY / 3 (2×1) | 4 | `PASTE_RIBBON_MENU_ID` / `PasteRibbonMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:233` |
| 1 | FORMAT / 1 (1×4) | 1 | `SetRangeFontFamilyCommand.id` / `FontFamilySelectorMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:260` |
| 1 | FORMAT / 5 (1×1) | 2 | `SetRangeFontSizeCommand.id` / `FontSizeSelectorMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:265` |
| 1 | FORMAT / 6 (1×1) | 2.1 | `SetRangeFontIncreaseCommand.id` / `FontSizeIncreaseMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:270` |
| 1 | FORMAT / 7 (1×1) | 2.2 | `SetRangeFontDecreaseCommand.id` / `FontSizeDecreaseMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:275` |
| 1 | LAYOUT / 1 (1×1) | 0 | `SetHorizontalTextAlignCommand.id` / `HorizontalAlignMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:325` |
| 1 | LAYOUT / 2 (1×1) | 6 | `SetTextWrapCommand.id` / `WrapTextMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:335` |
| 1 | LAYOUT / 3 (1×1) | 8 | `SetTextRotationCommand.id` / `TextRotateMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:345` |
| 1 | NUMBER / 1 (1×4) | 9 | `OpenNumfmtPanelOperator.id` / `FactoryOtherMenuItem` | `packages/sheets-numfmt-ui/src/menu/schema.ts:25` |
| 1 | OTHERS / 1 (2×1) | 0 | `AddRangeProtectionFromToolbarCommand.id` / `sheetPermissionToolbarMenuFactory` | `packages/sheets-ui/src/menu/schema.ts:373` |
| 1 | OTHERS / 2 (1×1) | 1 | `ToggleShortcutPanelOperation.id` / `ShortcutPanelMenuItemFactory` | `packages/ui/src/menu/schema.ts:52` |
| 2 | HISTORY / 1 (1×1) | 1 | `RedoCommand.id` / `RedoMenuItemFactory` | `packages/ui/src/menu/schema.ts:45` |
| 2 | HISTORY / 2 (1×1) | 3 | `ClearFormattingCommand.id` / `ClearFormattingMenuItemFactory` | `packages/ui/src/menu/schema.ts:35` |
| 2 | FORMAT / 1 (1×1) | 5 | `SetRangeBoldCommand.id` / `BoldMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:280` |
| 2 | FORMAT / 2 (1×1) | 6 | `SetRangeItalicCommand.id` / `ItalicMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:285` |
| 2 | FORMAT / 3 (1×1) | 7 | `SetRangeUnderlineCommand.id` / `UnderlineMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:290` |
| 2 | FORMAT / 4 (1×1) | 8 | `SetRangeStrickThroughCommand.id` / `StrikeThroughMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:295` |
| 2 | FORMAT / 5 (1×1) | 11 | `SetBorderBasicCommand.id` / `CellBorderSelectorMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:318` |
| 2 | FORMAT / 6 (1×1) | 10 | `SetBackgroundColorCommand.id` / `BackgroundColorSelectorMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:309` |
| 2 | FORMAT / 7 (1×1) | 9 | `SetRangeTextColorCommand.id` / `TextColorSelectorMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:300` |
| 2 | LAYOUT / 1 (1×1) | 3 | `SetVerticalTextAlignCommand.id` / `VerticalAlignMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:330` |
| 2 | LAYOUT / 2 (1×1) | 7 | `SetShrinkToFitCommand.id` / `ShrinkToFitMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:340` |
| 2 | LAYOUT / 3 (1×1) | 9 | `AddWorksheetMergeCommand.id` / `CellMergeMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:350` |
| 2 | NUMBER / 1 (1×1) | 9.1 | `SetPercentCommand.id` / `PercentMenuItem` | `packages/sheets-numfmt-ui/src/menu/schema.ts:30` |
| 2 | NUMBER / 2 (1×1) | 9.2 | `SetCurrencyCommand.id` / `CurrencySymbolIconMenuItem` | `packages/sheets-numfmt-ui/src/menu/schema.ts:35` |
| 2 | NUMBER / 3 (1×1) | 9.3 | `AddDecimalCommand.id` / `AddDecimalMenuItem` | `packages/sheets-numfmt-ui/src/menu/schema.ts:40` |
| 2 | NUMBER / 4 (1×1) | 9.4 | `SubtractDecimalCommand.id` / `SubtractDecimalMenuItem` | `packages/sheets-numfmt-ui/src/menu/schema.ts:45` |
| 2 | OTHERS / 2 (1×1) | 2 | `OpenFeatureSearchOperation.id` / `FeatureSearchMenuItemFactory` | `packages/ui/src/menu/schema.ts:56` |

### INSERT (9)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | MEDIA / 1 (2×1) | 0 | `SHEETS_IMAGE_MENU_ID` / `ImageMenuFactory` | `packages/sheets-drawing-ui/src/menu/schema.ts:42` |
| 1 | MEDIA / 2 (2×1) | 1 | `InsertHyperLinkToolbarOperation.id` / `insertLinkMenuToolbarFactory` | `packages/sheets-hyper-link-ui/src/menu/schema.ts:24` |
| 1 | MEDIA / 3 (1×1) | 2 | `ToggleSheetCommentPanelOperation.id` / `threadPanelMenuFactory` | `packages/sheets-thread-comment-ui/src/menu/schema.ts:24` |
| 1 | CELL / 1 (1×1) | 0 | `InsertQuickSheetDataValidationCommand.id` / `quickCheckboxMenuFactory` | `packages/sheets-data-validation-ui/src/menu/schema.ts:38` |
| 1 | CELL / 2 (1×1) | 2 | `QUICK_DATE_MENU_ID` / `quickDateMenuFactory` | `packages/sheets-data-validation-ui/src/menu/schema.ts:46` |
| 1 | CELL / 3 (1×1) | 4 | `AddIconSetConditionalRuleCommand.id` / `quickIconSetMenuFactory` | `packages/sheets-conditional-formatting-ui/src/menu/schema.ts:41` |
| 2 | MEDIA / 3 (1×1) | 2.1 | `ShowAddSheetCommentModalOperation.id` / `threadCommentMenuFactory` | `packages/sheets-thread-comment-ui/src/menu/schema.ts:29` |
| 2 | CELL / 1 (1×1) | 1 | `QUICK_DROPDOWN_MENU_ID` / `quickDropdownMenuFactory` | `packages/sheets-data-validation-ui/src/menu/schema.ts:42` |
| 2 | CELL / 2 (1×1) | 3 | `AddDataBarConditionalRuleCommand.id` / `quickDataBarMenuFactory` | `packages/sheets-conditional-formatting-ui/src/menu/schema.ts:37` |

### FORMULAS (12)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | BASIC / 1 (2×1) | 0 | `${InsertFunctionOperation.id}.common` / `InsertCommonFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:44` |
| 1 | BASIC / 2 (1×1) | 1 | `${InsertFunctionOperation.id}.financial` / `InsertFinancialFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:53` |
| 1 | BASIC / 3 (1×1) | 3 | `${InsertFunctionOperation.id}.text` / `InsertTextFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:61` |
| 1 | BASIC / 4 (1×1) | 5 | `${InsertFunctionOperation.id}.lookup` / `InsertLookupFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:69` |
| 1 | BASIC / 5 (1×1) | 7 | `${InsertFunctionOperation.id}.statistical` / `InsertStatisticalFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:77` |
| 1 | BASIC / 6 (1×1) | 9 | `${InsertFunctionOperation.id}.information` / `InsertInformationFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:85` |
| 1 | OTHERS / 1 (2×2) | 0 | `CalculationOptionsOperation.id` / `CalculationOptionsMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:95` |
| 2 | BASIC / 2 (1×1) | 2 | `${InsertFunctionOperation.id}.logical` / `InsertLogicalFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:57` |
| 2 | BASIC / 3 (1×1) | 4 | `${InsertFunctionOperation.id}.date` / `InsertDateFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:65` |
| 2 | BASIC / 4 (1×1) | 6 | `${InsertFunctionOperation.id}.math` / `InsertMathFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:73` |
| 2 | BASIC / 5 (1×1) | 8 | `${InsertFunctionOperation.id}.engineering` / `InsertEngineeringFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:81` |
| 2 | BASIC / 6 (1×1) | 10 | `${InsertFunctionOperation.id}.database` / `InsertDatabaseFunctionMenuItemFactory` | `packages/sheets-formula-ui/src/menu/schema.ts:89` |

### DATA (6)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | ORGANIZATION / 1 (1×1) | 0 | `SHEETS_SORT_MENU_ID` / `sortRangeMenuFactory` | `packages/sheets-sort-ui/src/menu/schema.ts:50` |
| 1 | ORGANIZATION / 2 (1×1) | 2 | `OpenTableSelectorOperation.id` / `sheetTableToolbarInsertMenuFactory` | `packages/sheets-table-ui/src/menu/schema.ts:25` |
| 1 | RULES / 1 (1×1) | 0 | `DATA_VALIDATION_MENU_ID` / `dataValidationMenuFactory` | `packages/sheets-data-validation-ui/src/menu/schema.ts:52` |
| 1 | OTHERS / 1 (2×2) | 0 | `OpenFindDialogOperation.id` / `FindReplaceMenuItemFactory` | `packages/find-replace/src/menu/schema.ts:24` |
| 2 | ORGANIZATION / 1 (1×1) | 1 | `SmartToggleSheetsFilterCommand.id` / `SmartToggleFilterMenuItemFactory` | `packages/sheets-filter-ui/src/menu/schema.ts:32` |
| 2 | RULES / 1 (1×1) | 1 | `OpenConditionalFormattingOperator.id` / `FactoryManageConditionalFormattingRule` | `packages/sheets-conditional-formatting-ui/src/menu/schema.ts:27` |

### VIEW (6)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | DISPLAY / 1 (2×1) | 0 | `SHEET_FROZEN_TOOLBAR_MENU_ID` / `SheetFrozenToolbarMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:382` |
| 1 | DISPLAY / 2 (2×1) | 0.5 | `WORKSHEET_BACKGROUND_IMAGE_MENU_ID` / `WorksheetBackgroundImageMenuFactory` | `packages/sheets-drawing-ui/src/menu/schema.ts:58` |
| 1 | DISPLAY / 3 (1×1) | 1 | `ZOOM_RATIO_MENU_ID` / `ZoomRatioMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:387` |
| 1 | VISIBILITY / 1 (1×1) | 0 | `ToggleGridlinesCommand.id` / `ToggleGridlinesMenuFactory` | `packages/sheets-ui/src/menu/schema.ts:394` |
| 2 | DISPLAY / 3 (1×1) | 2 | `ToggleCrosshairHighlightOperation.id` / `CrosshairHighlightMenuItemFactory` | `packages/sheets-crosshair-highlight/src/menu/schema.ts:24` |
| 2 | VISIBILITY / 1 (1×1) | 1 | `ToggleFullscreenOperation.id` / `FullscreenMenuItemFactory` | `packages/sheets-ui/src/menu/schema.ts:399` |

## docs

常规根 38；上下文根 7。含可选扩展，不等于示例默认挂载数。

### START (25)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | HISTORY / 1 (1×1) | 0 | `UndoCommand.id` / `UndoMenuItemFactory` | `packages/ui/src/menu/schema.ts:40` |
| 1 | HISTORY / 2 (1×1) | 2 | `ActivateFormatPainterOperation.id` / `FormatPainterMenuItemFactory` | `packages/ui/src/menu/schema.ts:30` |
| 1 | HISTORY / 3 (2×1) | 4 | `DocPasteSpecialCommand.id` / `PasteRibbonMenuFactory` | `packages/docs-ui/src/menu/schema.ts:239` |
| 1 | FORMAT / 1 (1×2) | 5.5 | `SetParagraphNamedStyleCommand.id` / `HeadingSelectorMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:283` |
| 1 | FORMAT / 3 (1×4) | 7 | `SetInlineFormatFontFamilyCommand.id` / `FontFamilySelectorMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:293` |
| 1 | FORMAT / 7 (1×2) | 6 | `SetInlineFormatFontSizeCommand.id` / `FontSizeSelectorMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:288` |
| 1 | LAYOUT / 1 (1×1) | 2 | `AlignOperationCommand.id` / `AlignMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:318` |
| 1 | LAYOUT / 2 (1×1) | 8 | `BulletListCommand.id` / `BulletListMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:328` |
| 1 | LAYOUT / 3 (2×1) | 10 | `OpenHeaderFooterPanelCommand.id` / `HeaderFooterMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:338` |
| 1 | OTHERS / 1 (2×1) | 0 | `OpenDocPermissionPanelOperation.id` / `DocPermissionMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:246` |
| 1 | OTHERS / 2 (1×1) | 1 | `ToggleShortcutPanelOperation.id` / `ShortcutPanelMenuItemFactory` | `packages/ui/src/menu/schema.ts:52` |
| 1 | OTHERS / 3 (1×1) | 3 | `OpenFindDialogOperation.id` / `DocsFindReplaceMenuItemFactory` | `packages/docs-find-replace/src/menu/schema.ts:24` |
| 2 | HISTORY / 1 (1×1) | 1 | `RedoCommand.id` / `RedoMenuItemFactory` | `packages/ui/src/menu/schema.ts:45` |
| 2 | HISTORY / 2 (1×1) | 3 | `ClearFormattingCommand.id` / `ClearFormattingMenuItemFactory` | `packages/ui/src/menu/schema.ts:35` |
| 2 | FORMAT / 1 (1×1) | 0 | `SetInlineFormatBoldCommand.id` / `BoldMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:253` |
| 2 | FORMAT / 2 (1×1) | 1 | `SetInlineFormatItalicCommand.id` / `ItalicMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:258` |
| 2 | FORMAT / 3 (1×1) | 2 | `SetInlineFormatUnderlineCommand.id` / `UnderlineMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:263` |
| 2 | FORMAT / 4 (1×1) | 3 | `SetInlineFormatStrikethroughCommand.id` / `StrikeThroughMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:268` |
| 2 | FORMAT / 5 (1×1) | 5 | `SetInlineFormatSuperscriptCommand.id` / `SuperscriptMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:278` |
| 2 | FORMAT / 6 (1×1) | 4 | `SetInlineFormatSubscriptCommand.id` / `SubscriptMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:273` |
| 2 | FORMAT / 7 (1×1) | 9 | `SetInlineFormatTextBackgroundColorCommand.id` / `BackgroundColorSelectorMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:307` |
| 2 | FORMAT / 8 (1×1) | 8 | `SetInlineFormatTextColorCommand.id` / `TextColorSelectorMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:298` |
| 2 | LAYOUT / 1 (1×1) | 7 | `OrderListCommand.id` / `OrderListMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:323` |
| 2 | LAYOUT / 2 (1×1) | 9 | `CheckListCommand.id` / `CheckListMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:333` |
| 2 | OTHERS / 2 (1×1) | 2 | `OpenFeatureSearchOperation.id` / `FeatureSearchMenuItemFactory` | `packages/ui/src/menu/schema.ts:56` |

### INSERT (10)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | MEDIA / 1 (2×1) | 4 | `DOC_BREAKS_MENU_ID` / `BreaksMenuFactory` | `packages/docs-ui/src/menu/schema.ts:359` |
| 1 | MEDIA / 2 (1×1) | 2 | `TABLE_MENU_ID` / `TableMenuFactory` | `packages/docs-ui/src/menu/schema.ts:345` |
| 1 | MEDIA / 3 (1×1) | 5 | `DOC_INSERT_EMOJI_MENU_ID` / `EmojiPickerMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:388` |
| 1 | MEDIA / 4 (1×1) | 0 | `DOCS_IMAGE_MENU_ID` / `ImageMenuFactory` | `packages/docs-drawing-ui/src/menu/schema.ts:41` |
| 1 | MEDIA / 5 (1×1) | 1 | `ShowDocHyperLinkEditPopupOperation.id` / `AddHyperLinkMenuItemFactory` | `packages/docs-hyper-link-ui/src/menu/schema.ts:25` |
| 1 | CELL / 1 (1×1) | 7 | `INSERT_TABLE_OF_CONTENTS_MENU_ID` / `InsertTableOfContentsMenuFactory` | `packages/docs-toc-ui/src/menu/schema.ts:51` |
| 1 | OTHERS / 1 (1×1) | 0 | `ToggleCommentPanelOperation.id` / `ToolbarDocCommentMenuItemFactory` | `packages/docs-thread-comment-ui/src/menu/schema.ts:33` |
| 2 | MEDIA / 3 (1×1) | 6 | `DOC_INSERT_SYMBOL_MENU_ID` / `SymbolPickerMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:393` |
| 2 | MEDIA / 5 (1×1) | 4 | `HorizontalLineCommand.id` / `HorizontalLineFactory` | `packages/docs-ui/src/menu/schema.ts:354` |
| 2 | OTHERS / 1 (1×1) | 0.1 | `StartAddCommentOperation.id` / `AddDocCommentMenuItemFactory` | `packages/docs-thread-comment-ui/src/menu/schema.ts:38` |

### VIEW (3)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | DISPLAY / 1 (2×1) | 0 | `DocOpenPageSettingCommand.id` / `PageSettingMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:400` |
| 1 | DISPLAY / 2 (1×1) | 1 | `DOC_ZOOM_RATIO_MENU_ID` / `DocZoomRatioMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:405` |
| 2 | DISPLAY / 2 (1×1) | 2 | `ToggleFullscreenOperation.id` / `DocFullscreenMenuItemFactory` | `packages/docs-ui/src/menu/schema.ts:410` |

### DOC_HEADER_FOOTER_RIBBON_TAB (5)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | insertion / 1 (2×1) | 0 | `DOC_HEADER_FOOTER_RIBBON_CONTROL.insertion` / `inline factory with label props.kind` | `packages/docs-ui/src/menu/header-footer-ribbon.ts:61` |
| 1 | navigation / 1 (2×1) | 0 | `DOC_HEADER_FOOTER_RIBBON_CONTROL.navigation` / `inline factory with label props.kind` | `packages/docs-ui/src/menu/header-footer-ribbon.ts:61` |
| 1 | options / 1 (2×1) | 0 | `DOC_HEADER_FOOTER_RIBBON_CONTROL.options` / `inline factory with label props.kind` | `packages/docs-ui/src/menu/header-footer-ribbon.ts:61` |
| 1 | position / 1 (2×1) | 0 | `DOC_HEADER_FOOTER_RIBBON_CONTROL.position` / `inline factory with label props.kind` | `packages/docs-ui/src/menu/header-footer-ribbon.ts:61` |
| 1 | close / 1 (2×1) | 0 | `CloseHeaderFooterCommand.id` / `() => ({
                        id: CloseHeaderFooterCommand.id,
                        type: MenuItemType.BUTTON,
                        icon: 'CloseIcon',
                        title: 'docs-ui.headerFooter.closeHeaderFooter',
                        tooltip: 'docs-ui.headerFooter.closeHeaderFooter',
                    })` | `packages/docs-ui/src/menu/header-footer-ribbon.ts:78` |

### DOC_TABLE_OF_CONTENTS_RIBBON_TAB (2)

| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |
|---:|---|---:|---|---|
| 1 | DOC_TABLE_OF_CONTENTS_RIBBON_GROUP / 1 (2×1) | 0 | `OpenTableOfContentsDialogOperation.id` / `UpdateTableOfContentsMenuFactory` | `packages/docs-toc-ui/src/menu/schema.ts:34` |
| 1 | DOC_TABLE_OF_CONTENTS_RIBBON_GROUP / 2 (2×1) | 1 | `DeleteTableOfContentsCommand.id` / `DeleteTableOfContentsMenuFactory` | `packages/docs-toc-ui/src/menu/schema.ts:39` |

## Pro 发布包

读取 60 个有 ES 入口的 1.6.0 包。下表按包列出 mergeMenu/appendRootMenu 注册调用中识别出的根声明次数，保留不同挂载分支及重复声明；不是唯一功能数，也不是桌面可见数。未解析列仅统计 schema 表达式，factory 的响应式条件需另追踪。

| 包 | 注册调用 | 常规根声明 | 上下文根声明 | 未解析表达式 |
|---|---:|---:|---:|---:|
| `@univerjs-pro/docs-callout-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/docs-chart-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/docs-code-ui` | 2 | 2 | 0 | 1 |
| `@univerjs-pro/docs-column-ui` | 2 | 0 | 0 | 0 |
| `@univerjs-pro/docs-history-ui` | 1 | 1 | 0 | 0 |
| `@univerjs-pro/docs-latex-ui` | 3 | 2 | 0 | 1 |
| `@univerjs-pro/docs-print` | 1 | 1 | 0 | 0 |
| `@univerjs-pro/docs-quote-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/docs-reference-ui` | 1 | 8 | 0 | 0 |
| `@univerjs-pro/docs-sdt-ui` | 2 | 1 | 13 | 0 |
| `@univerjs-pro/docs-shape-ui` | 3 | 4 | 12 | 0 |
| `@univerjs-pro/docs-table-ui` | 6 | 4 | 17 | 0 |
| `@univerjs-pro/embed-ui` | 3 | 0 | 0 | 5 |
| `@univerjs-pro/shape-editor-ui` | 2 | 1 | 0 | 1 |
| `@univerjs-pro/sheets-chart-ui` | 4 | 3 | 0 | 0 |
| `@univerjs-pro/sheets-history-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/sheets-outline-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/sheets-pivot-chart-ui` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/sheets-pivot-ui` | 2 | 2 | 0 | 3 |
| `@univerjs-pro/sheets-print` | 2 | 2 | 0 | 0 |
| `@univerjs-pro/sheets-shape-ui` | 3 | 3 | 13 | 0 |
| `@univerjs-pro/sheets-slicer-ui` | 2 | 2 | 4 | 0 |
| `@univerjs-pro/sheets-sparkline-ui` | 2 | 2 | 0 | 0 |

商业包的逐声明路径、order、gridLayout、注册语句行列及未解析分支可用 `node scripts/count-univer-pro-ribbon-schemas.mjs .research.local/univer-1.6.0/packages --bundle-only --inventory` 复查。没有执行商业 SDK。
