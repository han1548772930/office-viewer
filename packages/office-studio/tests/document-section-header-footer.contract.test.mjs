import assert from 'node:assert/strict'
import { test, after } from 'node:test'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { applySectionHeaderFooter, documentSectionIds, resolveSectionHeaderFooter } from '../src/features/document/services/section-header-footer.ts'
import { createParagraph } from '../src/features/document/model/document-model.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom' })
const { layoutDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
const { executeDocumentCommand } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
const { sceneFromDocument } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')
after(() => vite.close())

function modelWithSection() {
  const model = createBlankDocument()
  const first = model.blocks[0]
  const sectionBreak = { ...createParagraph(), type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } }
  const pageBreak = { ...createParagraph(), type: 'pageBreak', runs: [] }
  model.blocks = [first, sectionBreak, createParagraph(), pageBreak, createParagraph(), { ...pageBreak, id: `${pageBreak.id}-second` }, createParagraph()]
  return { model, sectionBreak }
}

test('section headers and footers inherit previous content and can be unlinked transactionally', () => {
  const { model, sectionBreak } = modelWithSection()
  applySectionHeaderFooter(model, 'first', { header: 'Company', footer: 'Confidential' })
  assert.deepEqual(documentSectionIds(model), ['first', sectionBreak.id])
  assert.equal(resolveSectionHeaderFooter(model, sectionBreak.id).header, 'Company')
  applySectionHeaderFooter(model, sectionBreak.id, { linkToPrevious: false, header: 'Quarterly report' })
  assert.equal(resolveSectionHeaderFooter(model, sectionBreak.id).header, 'Quarterly report')
  assert.equal(resolveSectionHeaderFooter(model, sectionBreak.id).footer, 'Confidential')
  assert.equal(resolveSectionHeaderFooter(model, 'first').header, 'Company')
})

test('section header/footer commands are persisted by document snapshots and layout selects first/even content', () => {
  const { model, sectionBreak } = modelWithSection()
  const target = { blockId: model.blocks[0].id, offset: 0 }
  executeDocumentCommand(model, { anchor: target, focus: target }, {
    id: 'doc.command.set-section-header-footer',
    sectionId: 'first',
    patch: { header: 'Main header', footer: 'Page footer', differentFirstPage: true, firstPageHeader: 'Cover', differentOddEven: true, evenPageHeader: 'Even header', evenPageFooter: 'Even footer' },
  })
  const reopened = JSON.parse(JSON.stringify(model))
  const layout = layoutDocument(reopened, { measure: text => text.length * 7 })
  assert.equal(layout.pageHeaderFooters[0].header, 'Cover')
  assert.equal(layout.pageHeaderFooters[0].footer, '')
  assert.equal(layout.pageHeaderFooters[3].header, 'Even header')
  assert.equal(layout.pageHeaderFooters[3].footer, 'Even footer')
  const scene = sceneFromDocument(reopened, layout, { width: 900, height: 5000, scrollLeft: 0, scrollTop: 0, zoom: 100 }, {
    selection: { anchor: target, focus: target }, focused: false, caretVisible: false, composing: false, images: new Map(),
  })
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === 'Cover'))
  assert.ok(scene.commands.some(command => command.kind === 'text' && command.text === 'Even header'))
  assert.equal(sectionBreak.type, 'sectionBreak')
})

test('page and total-page fields recalculate from the paginated document after command execution', () => {
  const { model } = modelWithSection()
  const target = { blockId: model.blocks[0].id, offset: 0 }
  const selection = { anchor: target, focus: target }
  executeDocumentCommand(model, selection, { id: 'doc.command.insert-page-field', sectionId: 'first', story: 'footer', field: 'PAGE' })
  executeDocumentCommand(model, selection, { id: 'doc.command.insert-page-field', sectionId: 'first', story: 'footer', field: 'NUMPAGES' })

  const snapshot = JSON.parse(JSON.stringify(model))
  const firstLayout = layoutDocument(snapshot, { measure: text => text.length * 7 })
  assert.equal(firstLayout.pageHeaderFooters[0].footer, ' 1 4')
  assert.equal(firstLayout.pageHeaderFooters.at(-1).footer, ' 4 4')

  snapshot.blocks.splice(1, 0, { ...createParagraph(), type: 'pageBreak', runs: [] })
  const expandedLayout = layoutDocument(snapshot, { measure: text => text.length * 7 })
  assert.equal(expandedLayout.pageHeaderFooters[0].footer, ' 1 5')
  assert.equal(expandedLayout.pageHeaderFooters.at(-1).footer, ' 5 5')
})

test('invalid section ids and malformed patch values fail before mutating the model', () => {
  const { model } = modelWithSection()
  const before = JSON.stringify(model)
  assert.throws(() => applySectionHeaderFooter(model, 'missing', { header: 'x' }), /Unknown document section/)
  assert.throws(() => applySectionHeaderFooter(model, 'first', { header: 5 }), /Invalid section header/)
  assert.equal(JSON.stringify(model), before)
})
