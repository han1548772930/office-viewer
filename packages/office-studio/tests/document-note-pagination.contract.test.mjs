import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// DOC-12 oversized-decoration contract drafts. No tests, browser, Office sample or performance acceptance were run.
let server, domain, model, stories, notes, settings, layout, hosts, scene, codec
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-note-pagination', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  stories = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story-blocks.ts')
  notes = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-editing.ts')
  settings = await server.ssrLoadModule('/packages/office-studio/src/features/document/notes/document-note-settings.ts')
  layout = await server.ssrLoadModule('/packages/office-studio/src/features/document/layout/document-layout.ts')
  hosts = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-host.ts')
  scene = await server.ssrLoadModule('/packages/office-studio/src/features/document/rendering/document-scene.ts')
  codec = await server.ssrLoadModule('/packages/office-studio/src/api/document-body-codec.ts')
})
afterAll(async () => { await server?.close() })
const copy = value => JSON.parse(JSON.stringify(value)), measurer = { measure: text => [...text].length * 6 }
function paragraph(id, text = '') {
  const block = model.createParagraph(text); block.id = id; block.runs[0].id = id + '-run'; block.runs[0].fontSize = 15
  block.style = { ...block.style, indent: 0, lineHeight: 1, spacingRule: 'auto', spaceBefore: 0, spaceAfter: 0, wordWrap: true }
  return block
}
const story = blocks => stories.projectDocumentStoryBlocks({ blocks, runs: [], paragraphs: [] }, 'display').content
const paragraphs = (prefix, count) => Array.from({ length: count }, (_, index) => paragraph(prefix + index, prefix + index))
function fixture(kind, separatorBlocks, noticeBlocks = [], continuationBlocks = paragraphs('CONT', 8), noteBlocks = paragraphs('NOTE', 18), columnCount = 1) {
  const value = domain.createBlankDocument(); value.id = 'note-pagination'
  value.page = { width: 260, height: 180, marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20 }
  value.blocks = [paragraph('reference', 'R')]
  const point = { blockId: 'reference', offset: 1 }
  notes.insertDocumentNote(value, { anchor: point, focus: point }, kind, { noteId: 'note', content: story(noteBlocks) })
  settings.setDocumentNoteSettings(value, kind, { separator: story(separatorBlocks), continuationNotice: story(noticeBlocks),
    continuationSeparator: story(continuationBlocks), ...(kind === 'footnote' ? { columnCount } : {}) })
  return value
}
function textBlocks(blocks) {
  return blocks.flatMap(block => block.type === 'table'
    ? model.tableForBlock(block).rows.flatMap(row => row.cells.flatMap(cell => textBlocks(cell.blocks ?? [])))
    : block.type === 'image' || block.type === 'divider' ? [] : [block])
}
function assertCompleteParagraphs(blocks, contents) {
  const lines = contents.flatMap(content => content.lines).filter(line => !line.repeatedHeader)
  for (const block of textBlocks(blocks)) {
    const actual = lines.filter(line => (line.target.cellBlockId ?? line.target.blockId) === block.id).sort((a, b) => a.start - b.start)
    assert.ok(actual.length, 'visible paragraph ' + block.id)
    let offset = 0
    for (const line of actual) {
      assert.equal(line.start, offset, 'one continuous source span ' + block.id); offset = line.end
      assert.ok(line.height > 0 && line.width > 0)
      if (line.clip) {
        assert.ok(line.clip.width > 0 && line.clip.height > 0)
        assert.ok(line.y >= line.clip.y - 1e-6 && line.y + line.height <= line.clip.y + line.clip.height + 1e-6, 'whole visible line ' + block.id)
      }
    }
    assert.equal(offset, model.runText(block.runs).length, 'complete source tail ' + block.id)
  }
}
function decorationInvocations(result, key) {
  const groups = []
  for (const content of result.noteDecorations.filter(content => content.kind === key)) {
    if (!content.continuedBefore) groups.push([])
    assert.ok(groups.length); groups.at(-1).push(content)
  }
  return groups
}
function assertOnlyRealReferences(value, result, kind, before) {
  assert.deepEqual(value, before)
  assert.deepEqual(Object.keys(kind === 'footnote' ? value.footnotes : value.endnotes), ['note'])
  assert.equal(result.lines.flatMap(line => line.glyphs).filter(glyph => glyph.note?.role === 'reference').length, 1)
  assert.equal(result.noteDecorations.flatMap(content => content.lines).flatMap(line => line.glyphs).some(glyph => glyph.note), false)
  assert.ok(result.diagnostics.some(item => item.code === 'DOCUMENT_NOTE_DECORATION_PAGINATED' && item.noteKind === kind))
  assert.ok(result.pages.length < 200, 'bounded progress for this fixed fixture')
  assert.ok(result.noteDecorations.every(content => [...content.lines, ...content.cells, ...content.images].every(item => item.host?.kind === 'noteDecoration')))
}

test('default note separator fields use the capped first width and full continuation width at the font baseline', () => {
  for (const kind of ['footnote', 'endnote']) for (const width of [200, 360, 800]) {
    const value = domain.createBlankDocument()
    value.id = 'default-separator-' + kind + '-' + width
    value.page = { width, height: 260, marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20 }
    value.blocks = [paragraph('reference', 'R')]
    const point = { blockId: 'reference', offset: 1 }
    notes.insertDocumentNote(value, { anchor: point, focus: point }, kind, { noteId: 'note', content: story(paragraphs('NOTE', 24)) })
    const before = copy(value), result = layout.layoutDocument(value, measurer)
    const decorations = result.noteDecorations.filter(content => content.kind !== 'continuationNotice')
    assert.ok(decorations.some(content => content.kind === 'separator'))
    assert.ok(decorations.some(content => content.kind === 'continuationSeparator'))
    for (const content of decorations) {
      assert.equal(content.dividers.length, 1)
      const divider = content.dividers[0], line = content.lines[0], em = 10 * 4 / 3
      assert.equal(divider.noteSeparator, true)
      assert.equal(divider.width, content.kind === 'separator' ? Math.min(192, content.width) : content.width)
      assert.ok(Math.abs(divider.height - em / 20) < 1e-9)
      assert.ok(Math.abs(divider.y + divider.height / 2 - (line.baseline - em / 4)) < 1e-9)
      assert.ok(Math.abs(content.height - line.height) < 1e-9, 'default paragraph adds no trailing spacing')
      assert.equal(line.target.blockId, 'default-note-decoration:' + kind + ':' + content.kind + ':paragraph')
      assert.equal(line.start, 0); assert.equal(line.end, 0)
    }
    assert.deepEqual(value, before)
    const selection = { anchor: point, focus: point }
    for (const zoom of [50, 100]) {
      const rendered = scene.sceneFromDocument(value, result, { width: 1200, height: result.height + 100, scrollLeft: 0, scrollTop: 0, zoom },
        { selection, focused: false, caretVisible: false, composing: false, images: new Map() })
      const thickness = 10 * 4 / 3 / 20 * zoom / 100
      assert.equal(rendered.commands.filter(command => command.kind === 'line' && Math.abs(command.width - thickness) < 1e-9).length, decorations.length)
    }
  }
})

test('unmarked custom fractional note rules preserve their existing width, placement, and one-pixel stroke', () => {
  for (const fraction of [1 / 3, 1 / 2, 1]) {
    const separator = paragraph('custom-separator')
    separator.metadata.noteSeparator = { version: 1, widthFraction: fraction, color: '#123456' }
    const value = fixture('footnote', [separator], [], [], [paragraph('NOTE', 'Note')]), before = copy(value)
    const result = layout.layoutDocument(value, measurer), content = result.noteDecorations.find(content => content.kind === 'separator')
    const line = content.lines[0], divider = content.dividers[0]
    assert.equal(divider.width, content.width * fraction)
    assert.equal(divider.height, 1); assert.equal(divider.noteSeparator, undefined)
    assert.equal(divider.y, line.y + Math.max(1, line.height - 2)); assert.equal(divider.color, '#123456')
    assert.deepEqual(value, before)
  }
})

test('public separator FIELD bodies retain their identities through save and reopen and use the same layout and scene', () => {
  for (const kind of ['footnote', 'endnote']) {
    const reopened = {}
    for (const source of [
      { decoration: 'separator', fieldType: 'FOOTNOTE_SEPARATOR', fontSize: 12 },
      { decoration: 'continuationSeparator', fieldType: 'FOOTNOTE_CONTINUATION_SEPARATOR', fontSize: 18 },
    ]) {
      const paragraphId = 'public-' + kind + '-' + source.decoration
      const range = { rangeId: paragraphId + '-field', rangeType: 1, startIndex: 0, endIndex: 0, wholeEntity: true,
        properties: { fieldType: source.fieldType } }
      const body = { dataStream: '\uFFFC\r\n', paragraphs: [{ paragraphId, startIndex: 1,
        paragraphStyle: { spaceAbove: { v: 0 }, spaceBelow: { v: 0 }, textStyle: { fs: source.fontSize } } }],
        sectionBreaks: [{ sectionId: paragraphId + '-section', startIndex: 2 }], customRanges: [range] }
      const original = copy(body), decoded = codec.decodeDocumentBody(body)
      assert.deepEqual(body, original)
      assert.equal(decoded.content.blocks.length, 1)
      assert.equal(decoded.content.blocks[0].metadata.noteSeparator.fieldType, source.fieldType)
      assert.deepEqual(decoded.content.blocks[0].metadata.univerNoteSeparatorRangeSource, range)
      const saved = codec.encodeDocumentBody(decoded.content, undefined, decoded.contentControls)
      assert.equal(saved.body.dataStream, body.dataStream)
      assert.deepEqual(saved.body.customRanges, [range])
      const restored = codec.decodeDocumentBody(saved.body)
      const block = restored.content.blocks[0]
      assert.equal(block.id, paragraphId)
      assert.equal(model.runText(block.runs), '')
      assert.equal(block.runs[0].fontSize, source.fontSize)
      assert.equal(block.style.spaceBefore, 0); assert.equal(block.style.spaceAfter, 0)
      assert.equal(block.metadata.noteSeparator.fieldType, source.fieldType)
      assert.deepEqual(block.metadata.univerNoteSeparatorRangeSource, range)
      reopened[source.decoration] = restored.content
    }
    for (const width of [200, 360, 800]) {
      const value = domain.createBlankDocument()
      value.id = 'public-separator-' + kind + '-' + width
      value.page = { width, height: 260, marginTop: 20, marginBottom: 20, marginLeft: 20, marginRight: 20 }
      value.blocks = [paragraph('reference', 'R')]
      const point = { blockId: 'reference', offset: 1 }
      notes.insertDocumentNote(value, { anchor: point, focus: point }, kind, { noteId: 'note', content: story(paragraphs('NOTE', 24)) })
      settings.setDocumentNoteSettings(value, kind, reopened)
      const before = copy(value), result = layout.layoutDocument(value, measurer)
      const decorations = result.noteDecorations.filter(content => content.kind !== 'continuationNotice')
      assert.ok(decorations.some(content => content.kind === 'separator'))
      assert.ok(decorations.some(content => content.kind === 'continuationSeparator'))
      for (const content of decorations) {
        const block = value.noteSettings[kind][content.kind].blocks[0]
        const em = block.runs[0].fontSize * 4 / 3
        assert.equal(content.lines.length, 1); assert.equal(content.dividers.length, 1)
        const line = content.lines[0], divider = content.dividers[0]
        const host = { kind: 'noteDecoration', noteKind: kind, decoration: content.kind }
        assert.equal(divider.noteSeparator, true)
        assert.equal(divider.width, content.kind === 'separator' ? Math.min(192, content.width) : content.width)
        assert.ok(Math.abs(divider.height - em / 20) < 1e-9)
        assert.ok(Math.abs(divider.y + divider.height / 2 - (line.baseline - em / 4)) < 1e-9)
        assert.equal(line.target.blockId, block.id); assert.equal(line.start, 0); assert.equal(line.end, 0)
        assert.deepEqual(line.host, host); assert.deepEqual(divider.host, host)
        assert.deepEqual(block.metadata.univerNoteSeparatorRangeSource, reopened[content.kind].blocks[0].metadata.univerNoteSeparatorRangeSource)
      }
      assert.deepEqual(value, before)
      for (const zoom of [50, 100]) {
        const rendered = scene.sceneFromDocument(value, result, { width: 1200, height: result.height + 100, scrollLeft: 0, scrollTop: 0, zoom },
          { selection: { anchor: point, focus: point }, focused: false, caretVisible: false, composing: false, images: new Map() })
        for (const decoration of ['separator', 'continuationSeparator']) {
          const contents = decorations.filter(content => content.kind === decoration)
          const divider = contents[0].dividers[0]
          const lines = rendered.commands.filter(command => command.kind === 'line'
            && Math.abs(command.width - divider.height * zoom / 100) < 1e-9
            && Math.abs(command.x2 - command.x1 - divider.width * zoom / 100) < 1e-9 && command.y1 === command.y2)
          assert.equal(lines.length, contents.length)
        }
      }
    }
  }
})

test('small legacy separator stories use their real decoration host for layout IDs and input geometry', () => {
  for (const kind of ['footnote', 'endnote']) {
    const value = fixture(kind, [paragraph('small-separator', 'Separator')], [], [], [paragraph('small-note', 'Note')])
    const legacy = paragraph('legacy-source', 'Legacy separator')
    value.noteSettings[kind].separator = { runs: copy(legacy.runs), paragraphs: [{ startOffset: 0, endOffset: 16, style: copy(legacy.style) }] }
    const before = copy(value), host = { kind: 'noteDecoration', noteKind: kind, decoration: 'separator' }
    const scoped = hosts.documentModelForBlockHost(value, host), result = layout.layoutDocument(value, measurer)
    const content = result.noteDecorations.filter(content => content.noteKind === kind && content.kind === 'separator')
    assert.equal(content.length, 1); assert.ok(content[0].lines.length > 0)
    for (const line of content[0].lines) { assert.deepEqual(line.host, host); assert.equal(line.target.blockId, scoped.blocks[0].id) }
    assert.equal((result.diagnostics ?? []).some(item => item.code === 'DOCUMENT_NOTE_DECORATION_PAGINATED'), false)
    assert.deepEqual(value, before); assert.deepEqual(Object.keys(kind === 'footnote' ? value.footnotes : value.endnotes), ['note'])
  }
})

test('a separator ending exactly at a complete page advances into the note without recursively adding more separators', () => {
  const value = fixture('footnote', paragraphs('SEP', 13), [], [], [paragraph('only-note', 'Note')]), before = copy(value)
  const result = layout.layoutDocument(value, measurer), separators = decorationInvocations(result, 'separator')
  assert.equal(separators.length, 1); assert.ok(separators[0].length >= 2)
  assert.equal(result.noteDecorations.filter(content => content.kind === 'continuationSeparator').length, 0)
  assert.equal(result.noteDecorations.filter(content => content.kind === 'continuationNotice').length, 0)
  assertCompleteParagraphs(value.noteSettings.footnote.separator.blocks, separators[0])
  const note = result.auxiliaryContent.filter(content => content.host.kind === 'note')
  assertCompleteParagraphs(value.footnotes.note.content.blocks, note)
  assert.ok(note[0].page > separators[0].at(-1).page); assertOnlyRealReferences(value, result, 'footnote', before)
})

test('oversized notice and continuation separator finish their complete story before the next real note fragments', () => {
  for (const kind of ['footnote', 'endnote']) {
    const value = fixture(kind, paragraphs('SEP', 13), paragraphs('NOTICE', 12)), before = copy(value), result = layout.layoutDocument(value, measurer)
    const note = result.auxiliaryContent.filter(content => content.host.kind === 'note' && content.host.noteKind === kind)
    const source = (kind === 'footnote' ? value.footnotes : value.endnotes).note.content
    assertCompleteParagraphs(source.blocks, note)
    for (const key of ['separator', 'continuationSeparator', 'continuationNotice']) {
      const groups = decorationInvocations(result, key)
      assert.ok(groups.length > 0)
      assert.ok(groups.length <= note.flatMap(content => content.lines).length, 'repeat only after actual note progress')
      for (const group of groups) assertCompleteParagraphs(value.noteSettings[kind][key].blocks, group)
    }
    assertOnlyRealReferences(value, result, kind, before)
  }
})

test('oversized rich table and image separators retain every cell paragraph, image and table strip in both note kinds', () => {
  for (const kind of ['footnote', 'endnote']) {
    const table = { ...paragraph('separator-table'), type: 'table', runs: [], metadata: { table: { columnWidths: [180], headerRowCount: 0,
      rows: Array.from({ length: 8 }, (_, row) => ({ id: 'row-' + row, cells: [{ id: 'cell-' + row, blocks: [paragraph('cell-p-' + row,
        ('Row' + row + '😀 ').repeat(12))], style: {}, rowSpan: 1, colSpan: 1 }] })) } } }
    const image = { ...paragraph('separator-image'), type: 'image', runs: [], metadata: { image: {
      src: 'data:image/png;base64,AA==', width: 100, height: 300, alt: 'Complete separator bitmap',
    } } }
    const value = fixture(kind, [paragraph('separator-head', 'Head'), table, image, paragraph('separator-tail', 'Tail')], [], [], [paragraph('only-note', 'Note')])
    const before = copy(value), result = layout.layoutDocument(value, measurer), groups = decorationInvocations(result, 'separator')
    assert.equal(groups.length, 1); assertCompleteParagraphs(value.noteSettings[kind].separator.blocks, groups[0])
    const images = groups[0].flatMap(content => content.images), fragments = groups[0].flatMap(content => content.tableFragments)
    assert.equal(images.length, 1); assert.ok(images[0].width > 0 && images[0].height > 0)
    assert.ok(Math.abs(images[0].height / images[0].width - 3) <= 1e-6)
    assert.ok(images[0].clip.y + images[0].clip.height >= images[0].y + images[0].height - 1e-6)
    const ordered = fragments.sort((a, b) => a.sourceStart - b.sourceStart)
    assert.ok(ordered.length > 1); assert.equal(ordered[0].sourceStart, 0)
    for (let index = 1; index < ordered.length; index++) assert.ok(Math.abs(ordered[index].sourceStart - ordered[index - 1].sourceEnd) <= 1e-6)
    assert.equal(ordered.at(-1).continuedAfter, false); assertOnlyRealReferences(value, result, kind, before)
  }
})

test('multi-column footnotes keep oversized decorations full-width and skip parity blank pages', () => {
  const value = fixture('footnote', paragraphs('SEP', 13), [], [], paragraphs('NOTE', 22), 2)
  value.blocks.push({ ...paragraph('next-section'), type: 'sectionBreak', sectionBreak: { type: 'oddPage' } }, paragraph('after', 'After'))
  const before = copy(value), result = layout.layoutDocument(value, measurer), blank = result.pages.filter(page => page.isParityBlank)
  assert.ok(blank.length)
  for (const page of blank) assert.equal([...result.auxiliaryContent, ...result.noteDecorations].some(content => content.page === page.index), false)
  for (const content of result.noteDecorations) assert.ok(Math.abs(content.width - result.pages[content.page].content.width) <= 1e-6)
  const note = result.auxiliaryContent.filter(content => content.host.kind === 'note')
  assert.ok(note.some(content => content.lines.some(line => line.column === 1)))
  assert.ok(note.every(content => content.columns.every(column => column.width < result.pages[content.page].content.width)))
  assertCompleteParagraphs(value.footnotes.note.content.blocks, note); assertOnlyRealReferences(value, result, 'footnote', before)
})

test('section-end endnotes resume later body text after all oversized separator and note content', () => {
  const value = fixture('endnote', paragraphs('SEP', 13), paragraphs('NOTICE', 12)), current = settings.getDocumentNoteSettings(value, 'endnote')
  settings.setDocumentNoteSettings(value, 'endnote', { ...current, position: 'sectEnd' })
  value.blocks.push({ ...paragraph('next-section'), type: 'sectionBreak', sectionBreak: { type: 'continuous', setup: { columnCount: 2, columnGap: 12 } } }, paragraph('after', 'After'))
  const before = copy(value), result = layout.layoutDocument(value, measurer), after = result.lines.find(line => line.target.blockId === 'after')
  const preceding = [...result.noteDecorations.flatMap(content => content.lines), ...result.auxiliaryContent.filter(content => content.host.kind === 'note').flatMap(content => content.lines)]
    .sort((a, b) => a.page - b.page || a.y - b.y).at(-1)
  assert.ok(after.page > preceding.page || after.page === preceding.page && after.y >= preceding.y + preceding.height - 1e-6)
  assertOnlyRealReferences(value, result, 'endnote', before)
})
