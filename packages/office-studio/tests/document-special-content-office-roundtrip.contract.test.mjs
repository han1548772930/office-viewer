import assert from 'node:assert/strict'
import { test, after } from 'node:test'
import { createServer } from 'vite'
import { createBlankDocument } from '../src/domain/model.ts'
import { createCodeBlock } from '../src/features/document/code-blocks/code-blocks.ts'
import { createLatexFormulaBlock, createLatexFormulaRun } from '../src/features/document/latex/latex-formula.ts'

const vite = await createServer({ configFile: './vite.config.ts', server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
const { documentHtml } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
const { normalizeDocument } = await vite.ssrLoadModule('/packages/office-studio/src/parser/documentNormalizer.ts')
const { resolveSectionHeaderFooter } = await vite.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
after(() => vite.close())

function envelope(blocks) {
  return {
    schemaVersion: 1, parserVersion: 'contract', format: 'docx', fileName: 'special.docx', size: 0,
    model: 'document', payload: { modelType: 'document', data: { documentId: 'special', blocks } },
    diagnostics: { warnings: [], errors: [], stats: {}, capabilities: [] },
  }
}

test('internal HTML snapshot preserves code language, source, line numbers, and formula modes', () => {
  const model = createBlankDocument()
  const code = createCodeBlock('const x = 1\nreturn x', 'typescript', true)
  const formula = createLatexFormulaRun('x^2', 'inline')
  const equation = createLatexFormulaBlock('\\frac{1}{2}')
  model.blocks = [code, { ...model.blocks[0], runs: [formula] }, equation]
  const html = documentHtml(model)
  assert.match(html, /data-office-code-block="v1" data-office-language="typescript" data-office-line-numbers="true"/u)
  assert.match(html, /const x = 1\nreturn x/u)
  assert.match(html, /data-office-latex="v1" data-office-latex-mode="inline" data-office-latex-source="x\^2"/u)
  assert.ok(html.includes('data-office-latex="v1" data-office-latex-mode="block" data-office-latex-source="\\frac{1}{2}"'))
})

test('DOCX tagged code paragraphs and formula controls normalize to editable model metadata', () => {
  const code = normalizeDocument(envelope([{
    id: 'code', type: 'paragraph', content: 'const x = 1\nreturn x',
    details: { inlines: [{ type: 'text', text: 'const x = 1\nreturn x' }], bookmarks: [{ name: 'office_code_v1_typescript_1' }] },
    children: [],
  }])).blocks[0]
  assert.deepEqual(code.metadata.codeBlock, { source: 'const x = 1\nreturn x', language: 'typescript', lineNumbers: true })
  assert.equal(code.runs.map((run) => run.text).join(''), code.metadata.codeBlock.source)

  const formula = normalizeDocument(envelope([{
    id: 'formula', type: 'paragraph', content: 'x^2',
    details: { inlines: [{ type: 'form_field', text: 'x^2' }] },
    children: [{ details: { field: { tag: 'office-latex:v1:inline' } } }],
  }])).blocks[0]
  assert.equal(formula.runs[0].text, '\uFFFC')
  assert.deepEqual(formula.runs[0].metadata.latexFormula, { latex: 'x^2', mode: 'inline', version: 1 })
})

test('DOCX section references normalize header variants and linked section inheritance', () => {
  const model = normalizeDocument({
    ...envelope([
      { id: 'body-a', type: 'paragraph', content: 'A', details: { inlines: [{ type: 'text', text: 'A' }] }, sourceRefs: [{ story: 'body' }] },
      { id: 'docx:sample-section-break-1', type: 'text', content: '', details: { type: 'section_break_details', sectionBreak: { type: 'nextPage' } }, sourceRefs: [{ story: 'body' }] },
      { id: 'body-b', type: 'paragraph', content: 'B', details: { inlines: [{ type: 'text', text: 'B' }] }, sourceRefs: [{ story: 'body' }] },
      { id: 'docx:sample-section-break-2', type: 'text', content: '', details: { type: 'section_break_details', sectionBreak: { type: 'continuous' } }, sourceRefs: [{ story: 'body' }] },
      { id: 'body-c', type: 'paragraph', content: 'C', details: { inlines: [{ type: 'text', text: 'C' }] }, sourceRefs: [{ story: 'body' }] },
      { id: 'header-main', type: 'header', content: 'Main header', sourceRefs: [{ story: 'header', partName: 'word/header1.xml' }] },
      { id: 'footer-main', type: 'footer', content: 'Main footer', sourceRefs: [{ story: 'footer', partName: 'word/footer1.xml' }] },
      { id: 'header-cover', type: 'header', content: 'Cover header', sourceRefs: [{ story: 'header', partName: 'word/header2.xml' }] },
      { id: 'header-even', type: 'header', content: 'Even header', sourceRefs: [{ story: 'header', partName: 'word/header3.xml' }] },
    ]),
    payload: { modelType: 'document', data: {
      documentId: 'sample',
      sectionLayout: [
        { startNodeIndex: 0, headerSections: [{ partName: 'word/header1.xml', referenceType: 'default' }], footerSections: [{ partName: 'word/footer1.xml', referenceType: 'default' }] },
        { startNodeIndex: 2, sectionPropertiesXml: '<w:sectPr><w:titlePg/></w:sectPr>', headerSections: [
          { partName: 'word/header2.xml', referenceType: 'first' },
          { partName: 'word/header3.xml', referenceType: 'even' },
        ], footerSections: [] },
        { startNodeIndex: 4, headerSections: [], footerSections: [] },
      ],
      evenAndOddHeaders: true,
      blocks: envelope([
        { id: 'body-a', type: 'paragraph', content: 'A', details: { inlines: [{ type: 'text', text: 'A' }] }, sourceRefs: [{ story: 'body' }] },
        { id: 'docx:sample-section-break-1', type: 'text', content: '', details: { type: 'section_break_details', sectionBreak: { type: 'nextPage' } }, sourceRefs: [{ story: 'body' }] },
        { id: 'body-b', type: 'paragraph', content: 'B', details: { inlines: [{ type: 'text', text: 'B' }] }, sourceRefs: [{ story: 'body' }] },
        { id: 'docx:sample-section-break-2', type: 'text', content: '', details: { type: 'section_break_details', sectionBreak: { type: 'continuous' } }, sourceRefs: [{ story: 'body' }] },
        { id: 'body-c', type: 'paragraph', content: 'C', details: { inlines: [{ type: 'text', text: 'C' }] }, sourceRefs: [{ story: 'body' }] },
        { id: 'header-main', type: 'header', content: 'Main header', sourceRefs: [{ story: 'header', partName: 'word/header1.xml' }] },
        { id: 'footer-main', type: 'footer', content: 'Main footer', sourceRefs: [{ story: 'footer', partName: 'word/footer1.xml' }] },
        { id: 'header-cover', type: 'header', content: 'Cover header', sourceRefs: [{ story: 'header', partName: 'word/header2.xml' }] },
        { id: 'header-even', type: 'header', content: 'Even header', sourceRefs: [{ story: 'header', partName: 'word/header3.xml' }] },
      ]).payload.data.blocks,
    } },
  })
  assert.deepEqual(model.sectionHeaderFooters.first, {
    header: 'Main header', footer: 'Main footer', firstPageHeader: '', firstPageFooter: '',
    evenPageHeader: '', evenPageFooter: '', pageFields: [], differentFirstPage: false, differentOddEven: true, linkToPrevious: false,
  })
  assert.deepEqual(model.sectionHeaderFooters['docx:sample-section-break-1'], {
    header: 'Main header', footer: 'Main footer', firstPageHeader: 'Cover header', firstPageFooter: '',
    evenPageHeader: 'Even header', evenPageFooter: '', pageFields: [], differentFirstPage: true, differentOddEven: true, linkToPrevious: false,
  })
  assert.deepEqual(model.sectionHeaderFooters['docx:sample-section-break-2'], {
    header: '', footer: '', firstPageHeader: '', firstPageFooter: '',
    evenPageHeader: '', evenPageFooter: '', pageFields: [], differentFirstPage: false, differentOddEven: true, linkToPrevious: true,
  })
  assert.equal(resolveSectionHeaderFooter(model, 'docx:sample-section-break-2').header, 'Main header')
  assert.equal(resolveSectionHeaderFooter(model, 'docx:sample-section-break-2').evenPageHeader, 'Even header')
})
