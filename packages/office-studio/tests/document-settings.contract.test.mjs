import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, cloneDocument } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { applyDocumentParagraphStylePatch, applyDocumentSectionSetupPatch, documentSectionSetup, validateDocumentParagraphStylePatch, validateDocumentSectionSetupPatch } from '../src/features/document/services/document-settings.ts'
import { saveDocumentSnapshot } from '../src/features/document/services/document-facade.ts'
import { DOCUMENT_MODERN_WIDTHS, DOCUMENT_PAPER_SIZES, documentModernWidth, documentPaperSize } from '../src/components/document/document-page-settings.ts'
import { layoutDocument } from '../src/features/document/layout/document-layout.ts'
import { sectionIdForBlock } from '../src/features/document/services/section-header-footer.ts'

const workbenchSource = readFileSync(new URL('../src/components/DocumentWorkbench.vue', import.meta.url), 'utf8')
const canvasSource = readFileSync(new URL('../src/components/document/DocumentCanvasViewport.vue', import.meta.url), 'utf8')

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
afterAll(() => vite.close())

function paragraph(blockId, offset = 0) { return { blockId, offset } }

test('paragraph setting updates cover alignment, four indent modes, spacing rules and wrap flags', () => {
  const block = createParagraph('Text')
  applyDocumentParagraphStylePatch(block, {
    align: 'justify', indentStart: 12, indentEnd: 7, indentFirstLine: 16, hanging: 4,
    spaceBefore: 8, spaceAfter: 14, spacingRule: 'exact', lineHeight: 18, wordWrap: false, snapToGrid: true,
  })
  assert.deepEqual(block.style, {
    ...block.style, align: 'justify', indentStart: 12, indentEnd: 7, indentFirstLine: 16, hanging: 4,
    spaceBefore: 8, spaceAfter: 14, spacingRule: 'exact', lineHeight: 18, wordWrap: false, snapToGrid: true,
  })
  assert.throws(() => validateDocumentParagraphStylePatch({ indentStart: -1 }), /non-negative/)
  assert.throws(() => validateDocumentParagraphStylePatch({ lineHeight: 0 }), /positive/)
  assert.throws(() => validateDocumentParagraphStylePatch({ spacingRule: 'auto', lineHeight: 6 }), /between 1 and 5/)
  assert.throws(() => validateDocumentParagraphStylePatch({ spacingRule: 'exact', lineHeight: 101 }), /between 1 and 100/)
  assert.doesNotThrow(() => applyDocumentParagraphStylePatch({ ...block, style: { ...block.style, spacingRule: 'exact', lineHeight: 18 } }, { align: 'center' }))
  assert.doesNotThrow(() => applyDocumentParagraphStylePatch({ ...block, style: { ...block.style, spacingRule: 'exact', lineHeight: 18 } }, { lineHeight: 20 }))
  assert.throws(() => validateDocumentParagraphStylePatch({ align: 'distributed' }), /alignment/)
})

test('paragraph settings flow through the normal document command mutation', () => {
  const model = createBlankDocument()
  const block = model.blocks[0]
  const selection = { anchor: paragraph(block.id), focus: paragraph(block.id) }
  executeDocumentCommand(model, selection, { id: 'doc.command.format-paragraph', patch: { align: 'center', wordWrap: false } })
  assert.equal(block.style.align, 'center')
  assert.equal(block.style.wordWrap, false)
})

test('paragraph patches reject unknown fields, wrong flags and out-of-range geometry without mutating the block', () => {
  const block = createParagraph('Keep model fidelity')
  const before = JSON.parse(JSON.stringify(block))
  for (const field of ['indent', 'indentStart', 'indentEnd', 'indentFirstLine', 'hanging', 'spaceBefore', 'spaceAfter']) {
    assert.doesNotThrow(() => validateDocumentParagraphStylePatch({ [field]: 0 }))
    assert.doesNotThrow(() => validateDocumentParagraphStylePatch({ [field]: 100 }))
    assert.throws(() => applyDocumentParagraphStylePatch(block, { [field]: 100.1 }), /between 0 and 100/)
    assert.throws(() => applyDocumentParagraphStylePatch(block, { [field]: NaN }), /finite/)
  }
  for (const patch of [{ unknown: 1 }, { metadata: 1 }, { wordWrap: 'true' }, { snapToGrid: 0 }, { align: undefined }, { indentStart: undefined }]) assert.throws(() => applyDocumentParagraphStylePatch(block, patch))
  assert.deepEqual(block, before)
  assert.throws(() => validateDocumentParagraphStylePatch(null), /must be an object/)
  assert.throws(() => validateDocumentParagraphStylePatch([]), /must be an object/)
  const fixed = { ...block, style: { ...block.style, spacingRule: 'exact', lineHeight: 24 } }
  assert.throws(() => applyDocumentParagraphStylePatch(fixed, { spacingRule: 'auto' }), /between 1 and 5/)
  assert.equal(fixed.style.spacingRule, 'exact')
  applyDocumentParagraphStylePatch(fixed, { spacingRule: 'auto', lineHeight: 1.5 })
  assert.equal(fixed.style.spacingRule, 'auto')
  assert.equal(fixed.style.lineHeight, 1.5)
})

test('section setup updates are validated, section scoped and preserved by snapshots', () => {
  const model = createBlankDocument()
  const first = createParagraph('Before')
  const section = { ...createParagraph(), type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  model.blocks = [first, section, createParagraph('After')]
  const selection = { anchor: paragraph(first.id), focus: paragraph(first.id) }
  executeDocumentCommand(model, selection, {
    id: 'doc.command.set-section-setup', sectionId: section.id,
    patch: { orientation: 'landscape', sectionStart: 'continuous', width: 900, height: 700, marginTop: 30, columnCount: 2, columnGap: 24, columnSeparator: 'line', pageNumberStart: 3 },
  })
  assert.deepEqual(documentSectionSetup(model, section.id), {
    width: 900, height: 700, orientation: 'landscape', marginTop: 30, marginRight: 72, marginBottom: 72, marginLeft: 72,
    columnCount: 2, columnGap: 24, columnSeparator: 'line', sectionStart: 'continuous', pageNumberStart: 3,
  })
  assert.deepEqual(saveDocumentSnapshot(cloneDocument(model)).blocks[1].sectionBreak.setup, section.sectionBreak.setup)
  assert.equal(section.sectionBreak.type, 'continuous')
  assert.throws(() => validateDocumentSectionSetupPatch({ columnCount: 0 }), /between 1 and 12/)
  assert.throws(() => validateDocumentSectionSetupPatch({ pageNumberStart: 0 }), /positive integer/)
  assert.throws(() => applyDocumentSectionSetupPatch(model, 'missing', { columnCount: 2 }), /Unknown document section/)
})

test('default section page setup remains compatible with legacy document page fields', () => {
  const model = createBlankDocument()
  applyDocumentSectionSetupPatch(model, 'first', { width: 800, height: 1100, orientation: 'portrait', marginLeft: 64, columnCount: 3, columnGap: 26, columnSeparator: 'line', pageNumberStart: 4 })
  assert.equal(model.page.width, 800)
  assert.equal(model.page.height, 1100)
  assert.equal(model.page.marginLeft, 64)
  assert.equal(documentSectionSetup(model, 'first').columnCount, 3)
  assert.equal(documentSectionSetup(model, 'first').columnGap, 26)
  assert.equal(documentSectionSetup(model, 'first').pageNumberStart, 4)
  assert.equal(saveDocumentSnapshot(cloneDocument(model)).sectionSetup.columnSeparator, 'line')
})

test('section settings preserve unused single-column gaps and enforce upstream numeric limits', () => {
  assert.doesNotThrow(() => validateDocumentSectionSetupPatch({ columnCount: 1, columnGap: 18, width: 1, height: 1 }))
  assert.doesNotThrow(() => validateDocumentSectionSetupPatch({ columnCount: 12, columnGap: 1000, sectionStart: 'unspecified' }))
  assert.throws(() => validateDocumentSectionSetupPatch({ columnGap: 1000.1 }), /1000/)
  assert.throws(() => validateDocumentSectionSetupPatch({ width: 0 }), /at least 1/)
  assert.throws(() => validateDocumentSectionSetupPatch({ height: 0 }), /at least 1/)
  const model = createBlankDocument()
  const section = { ...createParagraph(), type: 'sectionBreak', runs: [], sectionBreak: { type: 'continuous' } }
  model.blocks.push(section, createParagraph('Second section'))
  applyDocumentSectionSetupPatch(model, section.id, { sectionStart: 'unspecified' })
  assert.equal(documentSectionSetup(saveDocumentSnapshot(model), section.id).sectionStart, 'unspecified')
  assert.equal(sectionIdForBlock(model, section.id), section.id)
})

test('page setup presets preserve paper dimensions and modern content widths', () => {
  assert.deepEqual(DOCUMENT_PAPER_SIZES, {
    A3: { width: 1123, height: 1587 }, A4: { width: 794, height: 1124 }, A5: { width: 559, height: 794 },
    B4: { width: 944, height: 1344 }, B5: { width: 665, height: 944 }, Executive: { width: 696, height: 1008 },
    Folio: { width: 816, height: 1248 }, Legal: { width: 816, height: 1344 }, Letter: { width: 816, height: 1056 },
    Statement: { width: 528, height: 816 }, Tabloid: { width: 1056, height: 1632 },
  })
  assert.deepEqual(DOCUMENT_MODERN_WIDTHS, { narrow: 794, medium: 960, wide: 1123 })
  assert.equal(documentPaperSize({ width: 794, height: 1124, marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 }), 'A4')
  assert.equal(documentModernWidth({ width: 960, height: 1124, marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0 }), 'medium')
})

test('paragraph line spacing rules produce multiple, minimum and exact rendered line heights', () => {
  const model = createBlankDocument()
  const block = model.blocks[0]
  block.style = { ...block.style, lineHeight: 18, spacingRule: 'exact' }
  const exact = layoutDocument(model, { measure: text => text.length * 8 }).lines[0].height
  block.style = { ...block.style, lineHeight: 20, spacingRule: 'atLeast' }
  const atLeast = layoutDocument(model, { measure: text => text.length * 8 }).lines[0].height
  block.style = { ...block.style, lineHeight: 1.5, spacingRule: 'auto' }
  const automatic = layoutDocument(model, { measure: text => text.length * 8 }).lines[0].height
  assert.equal(exact, 18)
  assert.equal(atLeast, 20)
  assert.ok(automatic > atLeast)
})

test('paragraph and section settings operations open their real panels and commit through document commands', () => {
  assert.match(workbenchSource, /ribbonVariant: 'classic'/)
  assert.match(workbenchSource, /<DocumentSectionSettingsPanel[^>]*:section-options="sectionOptions"[^>]*@select-section="selectSection"/)
  assert.match(workbenchSource, /case 'sidebar\.operation\.doc-paragraph-setting-panel': paragraphSettingsPanelOpen\.value = true/)
  assert.match(workbenchSource, /case 'sidebar\.operation\.doc-section-setting-panel': sectionSettingsPanelOpen\.value = true/)
  assert.match(workbenchSource, /<DocumentParagraphSettingsPanel[^>]*:style="paragraphSettingsStyle"[^>]*@apply="applyParagraphSettings"/)
  assert.match(workbenchSource, /<DocumentSectionSettingsPanel[^>]*:section-id="headerFooterSectionId"[^>]*:setup="activeSectionSetup"[^>]*@apply="applySectionSettings"/)
  assert.match(workbenchSource, /execute\(\{ id: 'doc\.command\.format-paragraph', patch \}, false\)/)
  assert.match(workbenchSource, /execute\(\{ id: 'doc\.command\.set-section-setup', sectionId, patch \}, false\)/)
})

test('document context menu uses antdv-next Menu for clipboard, layout and table actions with edit permission checks', () => {
  assert.match(canvasSource, /if \(!point\) return[\s\S]*emit\('context-menu', \{ x: event\.clientX, y: event\.clientY, point \}\)/)
  assert.match(workbenchSource, /commandId: 'sidebar\.operation\.doc-paragraph-setting-panel'/)
  assert.match(workbenchSource, /commandId: 'sidebar\.operation\.doc-section-setting-panel'/)
  assert.match(workbenchSource, /if \(supportsDocumentTableOfContents\(model\.value\) && selection\.value\.focus\.row == null\) layoutActions\.push/)
  assert.match(workbenchSource, /disabled: !canEditContextSelection\.value/)
  assert.match(workbenchSource, /import \{[^}]*Menu[^}]*\} from 'antdv-next'/)
  assert.match(workbenchSource, /<Menu mode="vertical" :selectable="false" :items="documentContextMenuItems" @click="selectDocumentContextMenuItem" \/>/)
  assert.match(workbenchSource, /commandId: 'doc\.command\.copy'/)
  assert.match(workbenchSource, /commandId: 'doc\.command\.cut'/)
  assert.match(workbenchSource, /commandId: 'doc\.command\.paste'/)
  assert.match(workbenchSource, /commandId: 'doc\.command\.paste-special'/)
  assert.match(workbenchSource, /children: pasteSpecialActions\.map\(action => \(\{ key: action\.id, label: action\.title, disabled: action\.disabled \}\)\)/)
  assert.match(workbenchSource, /\['source', 'destination', 'text'\] as const/)
  assert.match(workbenchSource, /disabled: !canEditContextSelection\.value \|\| typeof navigator === 'undefined' \|\| !navigator\.clipboard/)
  assert.match(workbenchSource, /textStyle: \{ \.\.\.marksAtPoint\(model\.value, selection\.value\.focus\), \.\.\.pendingMarks\.value \}/)
  assert.match(workbenchSource, /commandId: 'doc\.command\.delete'/)
  assert.match(workbenchSource, /function selectDocumentContextMenuItem\(info: \{ key: string \| number \}\)/)
  assert.doesNotMatch(workbenchSource, /class="document-context-menu"[^>]*role="menu"/)
  assert.match(workbenchSource, /@context-menu="openDocumentContextMenu"/)
})
