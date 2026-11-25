import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Drafts for unified verification; this implementation batch does not run them.
let server, fields, topology, domain, model
beforeAll(async () => {
  server = await createServer({ cacheDir: 'node_modules/.vite/document-text-field-topology', optimizeDeps: { noDiscovery: true },
    server: { middlewareMode: true, hmr: false, ws: false }, appType: 'custom', logLevel: 'error' })
  fields = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-fields.ts')
  topology = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-text-field-topology.ts')
  domain = await server.ssrLoadModule('/packages/office-studio/src/domain/model.ts')
  model = await server.ssrLoadModule('/packages/office-studio/src/features/document/model/document-model.ts')
})
afterAll(async () => { await server?.close() })

function fixture() {
  const reference = id => ({ ...model.createRun(''), id, metadata: { noteReference: { kind: 'footnote', noteId: id } } })
  const block = { ...model.createParagraph(), id: 'p', runs: [model.createRun('AB'), reference('one'), reference('two'), model.createRun('CDEF')] }
  return { ...domain.createBlankDocument(), blocks: [block] }
}
function field(id, start, end, parent, sequence) {
  return { id, field: 'REF', instruction: `REF ${id}`, cachedResult: id,
    startOffset: start[0], startNoteReferenceCount: start[1], endOffset: end[0], endNoteReferenceCount: end[1],
    metadata: { source: { ...(parent === undefined ? {} : { parentFieldId: parent }),
      ...(sequence === undefined ? {} : { beginSequence: sequence }) } } }
}
function ranges(definitions) {
  const value = fixture()
  fields.setDocumentTextFields(value.blocks[0], definitions)
  return fields.documentTextFieldRanges(value)
}
const parents = definitions => topology.documentTextFieldTopologyParents(ranges(definitions))

test('explicit parents order equal result ranges and point children independently of collection order', () => {
  for (const [start, end] of [[[0, 0], [6, 2]], [[2, 1], [2, 2]], [[2, 1], [2, 1]]]) {
    assert.deepEqual(parents([field('inner', start, end, 'outer'), field('outer', start, end, null)]), [1, undefined])
  }
  assert.deepEqual(parents([field('outer', [0, 0], [6, 2], null), field('start-child', [0, 0], [0, 0], 'outer'),
    field('end-child', [6, 2], [6, 2], 'outer'), field('next-root', [6, 2], [6, 2], null)]), [undefined, 0, 0, undefined])
})

test('adjacent native ranges at one display offset remain separate siblings', () => {
  const definitions = [field('outer', [2, 0], [2, 2], null), field('first', [2, 0], [2, 1], 'outer'),
    field('second', [2, 1], [2, 2], 'outer'), field('point-child', [2, 2], [2, 2], 'outer'),
    field('point-root', [2, 2], [2, 2], null), field('following', [2, 2], [3, 2], null)]
  assert.deepEqual(parents(definitions), [undefined, 0, 0, 0, undefined, undefined])
  assert.deepEqual(parents([field('first', [2, 0], [2, 1]), field('second', [2, 1], [2, 2])]), [undefined, undefined])
})

test('geometry infers unspecified parents while explicit null retains root identity', () => {
  assert.deepEqual(parents([field('outer', [0, 0], [4, 2]), field('inner', [1, 0], [3, 2]),
    field('next', [4, 2], [6, 2])]), [undefined, 0, undefined])
  assert.deepEqual(parents([field('outer', [0, 0], [6, 2], undefined, 10),
    field('inner', [0, 0], [6, 2], undefined, 20)]), [undefined, 0])
  assert.deepEqual(parents([field('later', [0, 0], [6, 2], undefined, 20),
    field('earlier', [0, 0], [6, 2], undefined, 10)]), [1, undefined])
  assert.throws(() => parents([field('outer', [0, 0], [6, 2], null), field('root-inside', [1, 0], [3, 2], null)]), /兄弟/)
})

test('parent identity, native containment, cycles and source sequence are validated', () => {
  for (const definitions of [
    [field('self', [0, 0], [1, 0], 'self')],
    [field('missing', [0, 0], [1, 0], 'absent')],
    [field('first', [2, 0], [2, 2], 'second'), field('second', [2, 0], [2, 2], 'first')],
    [field('outer', [2, 0], [2, 1], null), field('outside', [2, 0], [2, 2], 'outer')],
    [field('outer', [2, 0], [2, 2], null), field('first', [2, 0], [2, 2], 'outer'), field('second', [2, 1], [3, 2], 'outer')],
  ]) assert.throws(() => parents(definitions))
  for (const parent of ['', ' ', 7, true, {}, []]) assert.throws(() => parents([field('invalid', [0, 0], [1, 0], parent)]), /parentFieldId/)
  for (const sequence of [-1, 0.5, NaN, '1', null, Number.MAX_SAFE_INTEGER + 1]) {
    assert.throws(() => parents([field('invalid', [0, 0], [1, 0], null, sequence)]), /beginSequence/)
  }
  const malformed = field('source', [0, 0], [1, 0]); malformed.metadata.source = []
  assert.throws(() => parents([malformed]), /source/)
  assert.throws(() => parents([field('first', [0, 0], [4, 2]), field('crossing', [3, 2], [6, 2])]), /交叉/)
})

test('the model rejects native endpoints outside their real paragraph geometry', () => {
  assert.throws(() => ranges([field('bad-rank', [2, 3], [2, 3], null)]), /序号/)
  assert.throws(() => ranges([field('backwards', [2, 2], [2, 1], null)]), /端点顺序/)
  const value = fixture(); value.blocks[0].runs[1].text = '12'
  assert.throws(() => fields.setDocumentTextFields(value.blocks[0], [field('inside-label', [3, 0], [4, 1], null)]), /字符内部/)
})

test('root depth zero permits 513 fields through depth 512 and rejects depth 513', () => {
  const definitions = Array.from({ length: 513 }, (_, index) => field(`f${index}`, [2, 0], [2, 0], index ? `f${index - 1}` : null))
  const result = parents(definitions)
  assert.equal(result.length, 513)
  assert.equal(result[0], undefined)
  assert.equal(result[512], 511)
  assert.throws(() => parents([...definitions, field('f513', [2, 0], [2, 0], 'f512')]), /512/)
  const inferred = Array.from({ length: 513 }, (_, index) => field(`g${index}`, [0, 0], [6, 2], undefined, index))
  assert.equal(parents(inferred)[512], 511)
  assert.throws(() => parents([...inferred, field('g513', [0, 0], [6, 2], undefined, 513)]), /512/)
})
