import assert from 'node:assert/strict'
import { afterAll, beforeAll, test } from 'vitest'
import { createServer } from 'vite'

// Deferred UI capability and request-scope contracts, 2026-11-22. Not executed.
let server, capabilities
beforeAll(async () => {
  server = await createServer({ server: { middlewareMode: true, hmr: false }, appType: 'custom', logLevel: 'error' })
  capabilities = await server.ssrLoadModule('/packages/office-studio/src/components/document/document-header-footer-capabilities.ts')
})
afterAll(async () => { await server?.close() })

const kinds = ['insertion', 'navigation', 'options', 'position']
const actions = [
  { action: 'insert-field', field: 'PAGE' },
  { action: 'insert-date-time', format: 'short', automatic: true },
  { action: 'refresh-fields' },
  { action: 'navigate', area: 'footer' },
  { action: 'set-options', patch: { linkToPrevious: false } },
  { action: 'set-options', patch: { differentFirstPage: true, differentOddEven: true } },
  { action: 'set-options', patch: { marginHeader: 24, marginFooter: 32 } },
  { action: 'set-options', patch: { marginHeader: 24, differentOddEven: true } },
]

for (const [name, state, controls, deniedActions] of [
  ['editable independent story', { disabled: false, sectionDisabled: false, contentDisabled: false }, [false, false, false, false], [false, false, false, false, false, false, false, false]],
  ['read-only story with editable section', { disabled: true, sectionDisabled: false }, [true, false, true, false], [true, true, true, false, true, true, false, true]],
  ['denied section', { disabled: true, sectionDisabled: true }, [true, false, true, true], [true, true, true, false, true, true, true, true]],
  ['linked story with editable section', { disabled: false, sectionDisabled: false, contentDisabled: true }, [true, false, false, false], [true, true, true, false, false, false, false, false]],
  ['legacy denied state', { disabled: true }, [true, false, true, true], [true, true, true, false, true, true, true, true]],
]) test(`${name} keeps navigation, content and section geometry permissions distinct`, () => {
  assert.deepEqual(kinds.map(kind => capabilities.headerFooterControlDisabled(kind, state)), controls)
  assert.deepEqual(actions.map(action => capabilities.headerFooterActionDisabled(state, action)), deniedActions)
})

test('a popup request requires its captured scope and rejects a newer editable target', () => {
  const opened = { requestScope: 7 }, newer = { requestScope: 8 }
  assert.equal(capabilities.headerFooterRequestScopeMatches(opened, { requestScope: 7 }), true)
  assert.equal(capabilities.headerFooterRequestScopeMatches(newer, { requestScope: 7 }), false)
  assert.equal(capabilities.headerFooterRequestScopeMatches(newer, undefined), false)
  assert.equal(capabilities.headerFooterRequestScopeMatches(newer, {}), false)
  assert.equal(capabilities.headerFooterRequestScopeMatches({}, {}), true)
  assert.equal(capabilities.headerFooterRequestScopeMatches({}, undefined), false)
})

test('serialized contextual actions retain explicit false, zero and their request scope', () => {
  for (const action of actions) assert.deepEqual(capabilities.parseHeaderFooterAction(JSON.stringify({ ...action, requestScope: 0 })), { ...action, requestScope: 0 })
  const action = { action: 'set-options', requestScope: 19, patch: { linkToPrevious: false, marginHeader: 0 } }
  const parsed = capabilities.parseHeaderFooterAction(JSON.stringify(action))
  assert.deepEqual(parsed, action)
  parsed.patch.marginHeader = 20
  assert.equal(action.patch.marginHeader, 0)
})

test('malformed contextual actions cannot reach a command or replace a popup scope', () => {
  for (const raw of [undefined, {}, '', '{', 'null', '[]', '7', '"insert-field"']) assert.equal(capabilities.parseHeaderFooterAction(raw), undefined)
  for (const action of [
    { action: 'unknown' },
    { action: 'navigate', area: 'body' },
    { action: 'insert-field', field: 'DATE' },
    { action: 'insert-date-time', format: 'unsupported', automatic: true },
    { action: 'insert-date-time', format: 'short', automatic: 'true' },
    { action: 'refresh-fields', requestScope: -1 },
    { action: 'refresh-fields', requestScope: '7' },
    { action: 'refresh-fields', requestScope: null },
    { action: 'set-options' },
    { action: 'set-options', patch: null },
    { action: 'set-options', patch: [] },
    { action: 'set-options', patch: { unknown: true } },
    { action: 'set-options', patch: { linkToPrevious: 1 } },
    { action: 'set-options', patch: { marginHeader: '12' } },
    { action: 'set-options', patch: { marginFooter: null } },
  ]) assert.equal(capabilities.parseHeaderFooterAction(JSON.stringify(action)), undefined, JSON.stringify(action))
})
