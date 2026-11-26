import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

let server
let commands
let modelUtils
let domain
let numbering
let lists
let controls
let stories
let headerFooter
let blockHosts
let storyFormatting
let toolbar

beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  commands = await server.ssrLoadModule('/packages/office-studio/src/features/document/commands/document-commands.ts')
  modelUtils = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  numbering = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-numbering.ts')
  lists = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-lists.ts')
  controls = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/content-controls.ts')
  stories = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-story.ts')
  headerFooter = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/section-header-footer.ts')
  blockHosts = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-block-host.ts')
  storyFormatting = await server.ssrLoadModule('/packages/office-studio/src/features/document/services/document-story-formatting.ts')
  toolbar = await server.ssrLoadModule('/packages/office-studio/src/components/document/document-toolbar-state.ts')
})

afterAll(async () => { await server?.close() })

function paragraph(id, text) { return { ...modelUtils.createParagraph(text), id } }
function documentWith(blocks) { return { ...domain.createBlankDocument(), id: 'named-style-doc', blocks } }
function selection(from, to) { return { anchor: { blockId: from, offset: 0 }, focus: { blockId: to, offset: 1 } } }
function documentWithStory(story = 'header', text = 'One\nTail') {
  const model = documentWith([paragraph('body', 'Body unchanged')])
  model.blocks[0].metadata.sectionId = 'first'
  model.sectionHeaderFooters = { first: { ...headerFooter.EMPTY_SECTION_HEADER_FOOTER, [story]: text, pageFields: [],
    storyContent: { [story]: { runs: [modelUtils.createRun(text, { italic: true, fontFamily: 'Georgia' })],
      paragraphs: stories.documentStoryParagraphs(text, model.blocks[0].style) } } } }
  return model
}

test('named paragraph styles update semantic heading levels, visual runs, and spacing across selected paragraphs', () => {
  const first = paragraph('first', 'Title')
  const second = paragraph('second', 'Section')
  const model = documentWith([first, second])

  commands.executeDocumentCommand(model, selection('first', 'second'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2',
  })

  assert.deepEqual(model.blocks.map((block) => ({ type: block.type, level: block.level, namedStyle: block.metadata.namedStyleType })), [
    { type: 'heading', level: 2, namedStyle: 'heading2' },
    { type: 'heading', level: 2, namedStyle: 'heading2' },
  ])
  assert.deepEqual(first.runs.map(({ text, bold, fontSize }) => ({ text, bold, fontSize })), [{ text: 'Title', bold: true, fontSize: 18 }])
  assert.equal(first.style.spaceBefore, 24)
  assert.equal(first.style.spaceAfter, 13.3)
})

test('title and subtitle remain outside the heading outline and persist in JSON snapshots', () => {
  const model = documentWith([paragraph('title', 'Document'), paragraph('subtitle', 'Summary')])
  commands.executeDocumentCommand(model, selection('title', 'title'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'title',
  })
  commands.executeDocumentCommand(model, selection('subtitle', 'subtitle'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'subtitle',
  })

  const restored = JSON.parse(JSON.stringify(model))
  assert.deepEqual(restored.blocks.map((block) => [block.type, block.level, block.metadata.namedStyleType]), [
    ['paragraph', 0, 'title'],
    ['paragraph', 0, 'subtitle'],
  ])
  assert.equal(restored.blocks[0].runs[0].fontSize, 26)
  assert.equal(restored.blocks[1].runs[0].color, '#999999')
})

test('normal style resets heading semantics and visual defaults', () => {
  const heading = paragraph('heading', 'Body again')
  commands.executeDocumentCommand(documentWith([heading]), selection('heading', 'heading'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading1',
  })
  commands.executeDocumentCommand(documentWith([heading]), selection('heading', 'heading'), {
    id: 'doc.command.set-paragraph-named-style', namedStyle: 'normal',
  })

  assert.equal(heading.type, 'paragraph')
  assert.equal(heading.level, 0)
  assert.equal(heading.metadata.namedStyleType, 'normal')
  assert.deepEqual(heading.runs.map(({ bold, fontSize, color }) => ({ bold, fontSize, color })), [
    { bold: false, fontSize: 11, color: '#1f2937' },
  ])
})

test('all eight named styles preserve paragraph identity, indentation and unrelated character properties', () => {
  for (const [namedStyle, style] of Object.entries(modelUtils.DOCUMENT_NAMED_PARAGRAPH_STYLES)) {
    const block = paragraph('stable-paragraph', 'Text')
    Object.assign(block.style, { indent: 2, indentStart: 35, indentEnd: -5, indentFirstLine: 9, hanging: 3, align: 'right', wordWrap: true })
    Object.assign(block.runs[0], { italic: true, underline: true, fontFamily: 'Georgia', href: 'https://example.com' })
    const runId = block.runs[0].id
    const model = documentWith([block])
    commands.executeDocumentCommand(model, selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle })
    assert.equal(block.id, 'stable-paragraph')
    assert.equal(block.runs[0].id, runId)
    assert.equal(block.metadata.namedStyleType, namedStyle)
    assert.deepEqual([block.type, block.level], [style.level ? 'heading' : 'paragraph', style.level])
    assert.deepEqual([block.style.indent, block.style.indentStart, block.style.indentEnd, block.style.indentFirstLine, block.style.hanging], [2, 35, -5, 9, 3])
    assert.equal(block.style.align, 'right')
    assert.equal(block.style.wordWrap, true)
    assert.deepEqual([block.runs[0].text, block.runs[0].italic, block.runs[0].underline, block.runs[0].fontFamily, block.runs[0].href],
      ['Text', true, true, 'Georgia', 'https://example.com'])
  }
})

test('numbered and bullet paragraphs retain list identity and a list depth independent of heading depth', () => {
  for (const [type, preset] of [['numbered', 'ORDER_LIST_2'], ['bullet', 'BULLET_LIST_3']]) {
    const block = { ...paragraph(`list-${type}`, 'List item'), type, level: 2, list: { id: 'stable-list', preset } }
    const model = documentWith([block])
    const list = structuredClone(block.list)
    commands.executeDocumentCommand(model, selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading5' })
    assert.deepEqual([block.type, block.level, block.metadata.numberingLevel], ['heading', 5, 2])
    assert.deepEqual(block.list, list)
    assert.equal(numbering.blockHasNumbering(block), true)
    assert.equal(numbering.blockNumberingLevelIndex(block), 2)
    commands.executeDocumentCommand(model, selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'normal' })
    assert.deepEqual([block.type, block.level], ['paragraph', 0])
    assert.deepEqual(block.list, list)
    assert.equal(numbering.blockNumberingLevelIndex(block), 2)
  }
})

test('legacy lists gain retained preset identity and imported numbering keeps its custom definition', () => {
  for (const [type, preset] of [['numbered', 'ORDER_LIST'], ['bullet', 'BULLET_LIST']]) {
    const block = { ...paragraph(`legacy-${type}`, 'Legacy item'), type, level: 3 }
    commands.executeDocumentCommand(documentWith([block]), selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'title' })
    assert.equal(block.list.preset, preset)
    assert.equal(numbering.blockNumberingLevelIndex(block), 3)
    const listId = block.list.id
    modelUtils.applyNamedParagraphStyle(block, 'heading1')
    assert.equal(block.list.id, listId)
  }
  const definition = numbering.numberingDefinitionFromParts({ abstractNumId: 7, levels: [
    { ilvl: 0, format: 'upperRoman', text: '%1)', start: 4 },
    { ilvl: 1, format: 'decimal', text: '%1.%2', start: 2 },
  ] }, { numId: 9, abstractNumId: 7 })
  const block = { ...paragraph('imported-list', 'Imported item'), type: 'numbered', level: 1,
    metadata: { numberingDefinition: definition, numberingListId: 'imported-num-9' } }
  const original = structuredClone(block.metadata)
  commands.executeDocumentCommand(documentWith([block]), selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle: 'subtitle' })
  assert.equal(block.list, undefined)
  assert.deepEqual(block.metadata.numberingDefinition, original.numberingDefinition)
  assert.equal(numbering.blockNumberingId(block), 'imported-num-9')
  assert.equal(numbering.blockNumberingLevelIndex(block), 1)
})

test('named styles retain checklist ownership, checked state and nesting across every preset', () => {
  const block = { ...modelUtils.createChecklistParagraph('Task', true, 'task-list'), id: 'task', level: 2 }
  const model = documentWith([block])
  for (const namedStyle of Object.keys(modelUtils.DOCUMENT_NAMED_PARAGRAPH_STYLES)) {
    commands.executeDocumentCommand(model, selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle })
    assert.equal(block.type, 'checklist')
    assert.equal(block.level, 2)
    assert.deepEqual(block.checklist, { listId: 'task-list', checked: true })
    assert.equal(block.metadata.namedStyleType, namedStyle)
  }
})

test('a lock outside the selected word and a denied final paragraph reject the entire named-style mutation', () => {
  for (const denied of ['content-control', 'paragraph']) {
    const first = paragraph('first', 'First'), last = paragraph('last', 'Last locked')
    const model = documentWith([first, last])
    if (denied === 'content-control') controls.createDocumentContentControl(model,
      { id: 'locked-tail', type: 'text', blockId: last.id, startOffset: 5, endOffset: 11, lock: 'contentLocked' })
    if (denied === 'paragraph') model.permissionRules = [{ id: 'deny-tail', scope: 'paragraph', targetId: last.id, action: 'edit', effect: 'deny', userIds: ['local-user'] }]
    const before = structuredClone(model)
    const state = commands.documentNamedParagraphStyleState(model, selection(first.id, last.id))
    assert.equal(state.available, false)
    assert.ok(state.reason)
    assert.throws(() => commands.executeDocumentCommand(model, selection(first.id, last.id),
      { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2' }))
    assert.deepEqual(model, before)
  }
})

test('a half-open range leaves the final paragraph and its locked content unchanged', () => {
  const first = paragraph('first', 'Selected'), last = paragraph('last', 'Locked')
  const model = documentWith([first, last])
  controls.createDocumentContentControl(model, { id: 'untouched-lock', type: 'text', blockId: last.id, startOffset: 0, endOffset: 6, lock: 'contentLocked' })
  const original = structuredClone(last)
  const forward = { anchor: { blockId: first.id, offset: 1 }, focus: { blockId: last.id, offset: 0 } }
  assert.deepEqual(commands.documentNamedParagraphStyleState(model, forward), { available: true })
  commands.executeDocumentCommand(model, { anchor: forward.focus, focus: forward.anchor },
    { id: 'doc.command.set-paragraph-named-style', namedStyle: 'title' })
  assert.equal(model.blocks[0].metadata.namedStyleType, 'title')
  assert.deepEqual(model.blocks[1], original)
})

test('table structural selections apply named styles only to the selected editable cell paragraphs', () => {
  const table = { ...paragraph('table', ''), type: 'table', runs: [], metadata: { table: modelUtils.createTable(1, 2) } }
  const first = table.metadata.table.rows[0].cells[0].blocks[0], last = table.metadata.table.rows[0].cells[1].blocks[0]
  first.runs = [modelUtils.createRun('Selected cell')]
  last.runs = [modelUtils.createRun('Other cell')]
  const original = structuredClone(last)
  const point = { blockId: table.id, row: 0, col: 0, cellBlockId: first.id, offset: 0 }
  const range = { anchor: point, focus: point, tableSelection: { tableId: table.id, kind: 'cell', range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 } } }
  commands.executeDocumentCommand(documentWith([table]), range, { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading4' })
  assert.equal(first.metadata.namedStyleType, 'heading4')
  assert.deepEqual(last, original)
})

test('legacy story named-style state follows the normalized start paragraph through page-field object offsets', () => {
  for (const story of stories.DOCUMENT_STORY_KEYS) {
    const model = documentWithStory(story), section = model.sectionHeaderFooters.first, content = section.storyContent[story]
    content.paragraphs[0].metadata = { namedStyleType: 'title' }
    content.paragraphs[1].metadata = { namedStyleType: 'heading4' }
    section.pageFields = [{ story, field: 'PAGE', offset: 3, cachedResult: '12' }]
    const reversed = { sectionId: 'first', story, start: 9, end: 1 }
    const before = structuredClone(model)
    assert.equal(storyFormatting.documentStorySelectionFormatting(model, reversed).namedStyle, 'title')
    assert.equal(toolbar.documentStoryToolbarState(model, reversed).namedStyle, 'title')
    assert.equal(toolbar.documentStoryToolbarState(model, { ...reversed, start: 5, end: 5 }).namedStyle, 'heading4')
    assert.deepEqual(model, before)
  }
})

test('legacy story ranges reuse the existing named-style host command while retaining fields and other paragraphs', () => {
  for (const [namedStyle, style] of Object.entries(modelUtils.DOCUMENT_NAMED_PARAGRAPH_STYLES)) {
    const model = documentWithStory(), section = model.sectionHeaderFooters.first
    section.footer = 'Other story'
    section.pageFields = [{ story: 'header', field: 'PAGE', offset: 3, instruction: 'PAGE', cachedResult: '12' }]
    const host = { kind: 'story', sectionId: 'first', story: 'header' }
    const target = blockHosts.documentBlockHostSelectionFromLegacyRange(model, 'first', 'header', 1, 5)
    const before = blockHosts.documentModelForBlockHost(model, host)
    const untouched = structuredClone(before.blocks[1]), body = structuredClone(model.blocks)
    assert.equal(commands.documentCommandSupportsBlockHost({ id: 'doc.command.set-paragraph-named-style' }), true)
    assert.equal(commands.documentNamedParagraphStyleState(before, target).available, true)
    assert.deepEqual(commands.executeDocumentCommand(model, target, { id: 'doc.command.set-paragraph-named-style', namedStyle, host }), target)
    const after = blockHosts.documentModelForBlockHost(model, host), field = model.sectionHeaderFooters.first.pageFields[0]
    assert.deepEqual([after.blocks[0].metadata.namedStyleType, after.blocks[0].type, after.blocks[0].level],
      [namedStyle, style.level ? 'heading' : 'paragraph', style.level])
    assert.deepEqual(after.blocks[1], untouched)
    assert.deepEqual(model.blocks, body)
    assert.equal(model.sectionHeaderFooters.first.footer, 'Other story')
    assert.deepEqual([field.field, field.offset, field.instruction, field.cachedResult, field.textStyle.fontSize, field.textStyle.bold],
      ['PAGE', 3, 'PAGE', '12', style.fontSize, style.bold])
    const restored = JSON.parse(JSON.stringify(model))
    assert.equal(toolbar.documentStoryToolbarState(restored, { sectionId: 'first', story: 'header', start: 1, end: 3 }).namedStyle, namedStyle)
  }
})

test('whole-paragraph story locks, paragraph permissions and previous-section links reject named styles atomically', () => {
  for (const denied of ['content-lock', 'paragraph-permission', 'linked-story']) {
    const model = documentWithStory('header', 'One rest\nTail')
    if (denied === 'content-lock') model.sectionHeaderFooters.first.storyContent.header.locks = [{ id: 'locked-rest', startOffset: 4, endOffset: 8, lock: 'contentLocked' }]
    if (denied === 'linked-story') {
      model.blocks.push({ ...modelUtils.createParagraph(), id: 'second', type: 'sectionBreak', runs: [], sectionBreak: { type: 'nextPage' } },
        paragraph('second-body', 'Other section'))
      model.sectionHeaderFooters.second = { ...structuredClone(model.sectionHeaderFooters.first), linkToPrevious: true }
    }
    const host = { kind: 'story', sectionId: denied === 'linked-story' ? 'second' : 'first', story: 'header' }
    if (denied === 'paragraph-permission') model.permissionRules = [{ id: 'deny-story-paragraph', scope: 'paragraph',
      targetId: blockHosts.documentModelForBlockHost(model, host).blocks[0].id, action: 'edit', effect: 'deny', userIds: ['local-user'] }]
    const target = blockHosts.documentBlockHostSelectionFromLegacyRange(model, host.sectionId, host.story, 1, 2)
    const before = structuredClone(model)
    if (denied !== 'linked-story') assert.equal(commands.documentNamedParagraphStyleState(blockHosts.documentModelForBlockHost(model, host), target).available, false)
    if (denied === 'linked-story') assert.throws(() => blockHosts.documentModelForBlockHost(model, host, true), /请先取消“链接到前一节”/)
    assert.throws(() => commands.executeDocumentCommand(model, target, { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2', host }))
    assert.deepEqual(model, before)
  }
})

test('list buttons retain every named style and cancel back to the original paragraph without touching the final boundary', () => {
  for (const [namedStyle, style] of Object.entries(modelUtils.DOCUMENT_NAMED_PARAGRAPH_STYLES)) {
    for (const type of ['bullet', 'numbered', 'checklist']) {
      const block = paragraph('styled', 'Text'), last = paragraph('last', 'Unselected')
      Object.assign(block.style, { indentStart: 35, indentFirstLine: 9, align: 'right' })
      Object.assign(block.runs[0], { italic: true, fontFamily: 'Georgia', href: 'https://example.com' })
      if (type !== 'checklist') {
        modelUtils.applyNamedParagraphStyle(last, 'heading5')
        last.list = { id: 'adjacent-list', preset: type === 'bullet' ? 'BULLET_LIST_3' : 'ORDER_LIST_2' }
        last.metadata.numberingLevel = 3
      }
      const model = documentWith([block, last])
      commands.executeDocumentCommand(model, selection(block.id, block.id), { id: 'doc.command.set-paragraph-named-style', namedStyle })
      const original = structuredClone(block), untouched = structuredClone(last)
      const target = { anchor: { blockId: block.id, offset: 1 }, focus: { blockId: last.id, offset: 0 } }
      commands.executeDocumentCommand(model, target, { id: 'doc.command.set-block-type', type })
      assert.deepEqual([block.metadata.namedStyleType, block.type, block.level],
        [namedStyle, type === 'checklist' ? 'checklist' : style.level > 0 ? 'heading' : type, type === 'checklist' ? 0 : style.level])
      assert.equal(numbering.blockNumberingLevelIndex(block), 0)
      if (type !== 'checklist') assert.equal(numbering.blockNumberingId(block), 'adjacent-list')
      assert.deepEqual(block.runs, original.runs)
      assert.deepEqual(block.style, original.style)
      assert.equal(toolbar.documentToolbarState(model, target).namedStyle, namedStyle)
      assert.deepEqual(last, untouched)
      commands.executeDocumentCommand(model, target, { id: 'doc.command.set-block-type', type })
      assert.deepEqual(block, original)
      assert.deepEqual(last, untouched)
    }
  }
})

test('all twelve list presets preserve independent heading depth and cancellation also handles imported and unnamed headings', () => {
  const definition = numbering.numberingDefinitionFromParts({ abstractNumId: 4,
    levels: Array.from({ length: 4 }, (_, ilvl) => ({ ilvl, format: 'upperRoman', text: `%${ilvl + 1})`, start: 1 })) },
  { numId: 7, abstractNumId: 4 })
  for (const source of ['preset', 'imported', 'unnamed']) {
    for (const preset of [undefined, ...Object.values(lists.listPickerPresets).flat()]) {
      const block = paragraph('heading', 'Numbered heading')
      modelUtils.applyNamedParagraphStyle(block, 'heading2')
      block.metadata.numberingLevel = 3
      if (source !== 'imported') block.list = { id: 'existing-list', preset: 'ORDER_LIST_2' }
      if (source === 'imported') Object.assign(block.metadata, { numberingDefinition: structuredClone(definition), numberingListId: 'imported-list' })
      if (source === 'unnamed') delete block.metadata.namedStyleType
      const original = structuredClone(block), model = documentWith([block]), target = selection(block.id, block.id)
      if (preset) commands.executeDocumentCommand(model, target, { id: 'doc.command.change-list-type', listType: preset })
      assert.deepEqual([block.type, block.level, block.metadata.namedStyleType, numbering.blockNumberingLevelIndex(block)],
        ['heading', 2, original.metadata.namedStyleType, 3])
      if (preset) assert.equal(lists.blockListPreset(block), preset)
      if (preset) assert.equal(block.metadata.numberingDefinition, undefined)
      assert.equal(lists.paragraphListMarkers(model.blocks).has(block.id), true)
      commands.executeDocumentCommand(model, target, { id: 'doc.command.set-block-type', type: preset ? lists.listKind(preset) : 'numbered' })
      assert.deepEqual([block.type, block.level, block.metadata.namedStyleType], ['heading', 2, original.metadata.namedStyleType])
      assert.equal(numbering.blockHasNumbering(block), false)
      assert.equal(block.metadata.numberingLevel, undefined)
      assert.deepEqual(block.runs, original.runs)
      assert.deepEqual(block.style, original.style)
    }
  }
})

test('story list toggles keep named headings, page-field caches, body text and the next paragraph through the existing host command', () => {
  for (const type of ['bullet', 'numbered']) {
    const model = documentWithStory(), section = model.sectionHeaderFooters.first
    section.footer = 'Other story'
    section.pageFields = [{ story: 'header', field: 'PAGE', offset: 3, instruction: 'PAGE', cachedResult: '12' }]
    const host = { kind: 'story', sectionId: 'first', story: 'header' }
    const target = blockHosts.documentBlockHostSelectionFromLegacyRange(model, 'first', 'header', 1, 5)
    commands.executeDocumentCommand(model, target, { id: 'doc.command.set-paragraph-named-style', namedStyle: 'heading2', host })
    const original = blockHosts.documentModelForBlockHost(model, host), body = structuredClone(model.blocks)
    commands.executeDocumentCommand(model, target, { id: 'doc.command.set-block-type', type, host })
    const listed = blockHosts.documentModelForBlockHost(model, host)
    assert.deepEqual([listed.blocks[0].type, listed.blocks[0].level, listed.blocks[0].metadata.namedStyleType, numbering.blockNumberingLevelIndex(listed.blocks[0])],
      ['heading', 2, 'heading2', 0])
    assert.equal(lists.paragraphListMarkers(listed.blocks).has(listed.blocks[0].id), true)
    commands.executeDocumentCommand(model, target, { id: 'doc.command.set-block-type', type, host })
    assert.deepEqual(blockHosts.documentModelForBlockHost(model, host).blocks, original.blocks)
    assert.deepEqual(model.blocks, body)
    assert.equal(model.sectionHeaderFooters.first.footer, 'Other story')
    const field = model.sectionHeaderFooters.first.pageFields[0]
    assert.deepEqual([field.field, field.offset, field.instruction, field.cachedResult], ['PAGE', 3, 'PAGE', '12'])
    assert.equal(toolbar.documentStoryToolbarState(JSON.parse(JSON.stringify(model)), { sectionId: 'first', story: 'header', start: 1, end: 3 }).namedStyle, 'heading2')
  }
})
