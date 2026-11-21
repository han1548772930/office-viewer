import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { test, afterAll } from 'vitest'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createParagraph, cloneDocument } from '../src/features/document/model/document-model.ts'
import { executeDocumentCommand } from '../src/features/document/commands/document-commands.ts'
import { applyDocumentParagraphStylePatch, applyDocumentSectionSetupPatch, documentSectionSetup, validateDocumentParagraphStylePatch, validateDocumentSectionSetupPatch } from '../src/features/document/services/document-settings.ts'
import { saveDocumentSnapshot } from '../src/features/document/services/document-facade.ts'

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

test('section setup updates are validated, section scoped and preserved by snapshots', () => {
  const model = createBlankDocument()
  const first = createParagraph('Before')
  const section = { ...createParagraph(), type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  model.blocks = [first, section, createParagraph('After')]
  const selection = { anchor: paragraph(first.id), focus: paragraph(first.id) }
  executeDocumentCommand(model, selection, {
    id: 'doc.command.set-section-setup', sectionId: section.id,
    patch: { orientation: 'landscape', width: 900, height: 700, marginTop: 30, columnCount: 2, columnGap: 24, columnSeparator: 'line', pageNumberStart: 3 },
  })
  assert.deepEqual(documentSectionSetup(model, section.id), {
    width: 900, height: 700, orientation: 'landscape', marginTop: 30, marginRight: 72, marginBottom: 72, marginLeft: 72,
    columnCount: 2, columnGap: 24, columnSeparator: 'line', pageNumberStart: 3,
  })
  assert.deepEqual(saveDocumentSnapshot(cloneDocument(model)).blocks[1].sectionBreak.setup, section.sectionBreak.setup)
  assert.throws(() => validateDocumentSectionSetupPatch({ columnCount: 0 }), /between 1 and 12/)
  assert.throws(() => validateDocumentSectionSetupPatch({ pageNumberStart: 0 }), /positive integer/)
  assert.throws(() => applyDocumentSectionSetupPatch(model, 'missing', { columnCount: 2 }), /Unknown document section/)
})

test('default section page setup remains compatible with legacy document page fields', () => {
  const model = createBlankDocument()
  applyDocumentSectionSetupPatch(model, 'first', { width: 800, height: 1100, orientation: 'portrait', marginLeft: 64 })
  assert.equal(model.page.width, 800)
  assert.equal(model.page.height, 1100)
  assert.equal(model.page.marginLeft, 64)
  assert.equal(documentSectionSetup(model, 'first').columnCount, 1)
})

test('paragraph and section settings operations open their real panels and commit through document commands', () => {
  assert.match(workbenchSource, /case 'sidebar\.operation\.doc-paragraph-setting-panel': paragraphSettingsPanelOpen\.value = true/)
  assert.match(workbenchSource, /case 'sidebar\.operation\.doc-section-setting-panel': sectionSettingsPanelOpen\.value = true/)
  assert.match(workbenchSource, /<DocumentParagraphSettingsPanel[^>]*:style="paragraphSettingsStyle"[^>]*@apply="applyParagraphSettings"/)
  assert.match(workbenchSource, /<DocumentSectionSettingsPanel[^>]*:section-id="headerFooterSectionId"[^>]*:setup="activeSectionSetup"[^>]*@apply="applySectionSettings"/)
  assert.match(workbenchSource, /execute\(\{ id: 'doc\.command\.format-paragraph', patch \}\)/)
  assert.match(workbenchSource, /execute\(\{ id: 'doc\.command\.set-section-setup', sectionId, patch \}\)/)
})

test('document context menu exposes the ordered LAYOUT operations and applies table scope and edit permission rules', () => {
  assert.match(canvasSource, /if \(!point\) return[\s\S]*emit\('context-menu', \{ x: event\.clientX, y: event\.clientY, point \}\)/)
  assert.match(workbenchSource, /commandId: 'sidebar\.operation\.doc-paragraph-setting-panel'/)
  assert.match(workbenchSource, /commandId: 'sidebar\.operation\.doc-section-setting-panel'/)
  assert.match(workbenchSource, /if \(selection\.value\.focus\.row == null\) layoutActions\.push/)
  assert.match(workbenchSource, /disabled: !canEditContextSelection\.value/)
  assert.match(workbenchSource, /role="group" :aria-label="createRibbonTranslator\(officeLocale, 'doc'\)\('doc\.context\.layout-group'/)
  assert.match(workbenchSource, /@context-menu="openDocumentContextMenu"/)
})
