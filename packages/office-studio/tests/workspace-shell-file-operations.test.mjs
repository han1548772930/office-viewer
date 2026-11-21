import assert from 'node:assert/strict'
import { test } from 'vitest'
import { createWorkspaceFileOperations, workspaceFileNameError } from '../../../examples/workbench/src/workspace-file-operations.ts'

function deferred() {
  let resolve
  let reject
  const promise = new Promise((res, rej) => { resolve = res; reject = rej })
  return { promise, resolve, reject }
}

function fixture(overrides = {}) {
  const calls = { notices: [], downloads: [], savepoints: [], tasks: [], importsCancelled: 0, writes: [] }
  let current = { kind: 'workbook', id: 'book-1', revision: 2, generation: 1, name: 'Original.xlsx', text: 'original' }
  const services = {
    importFile: async () => true,
    cancelImport: () => { calls.importsCancelled++ },
    capture: (kind = current.kind) => ({
      kind,
      model: { text: current.text },
      fileName: current.name,
      savepoint: { kind, unitId: current.id, revision: current.revision, generation: current.generation },
    }),
    write: async (snapshot, options) => {
      calls.writes.push({ snapshot, options })
      return { bytes: new Uint8Array([0x50, 0x4b]), warnings: [] }
    },
    download: (bytes, snapshot) => { calls.downloads.push({ bytes, snapshot }) },
    markSaved: (savepoint) => { calls.savepoints.push(savepoint); return true },
    onExportTask: (task) => { calls.tasks.push(task) },
    notify: (type, message) => { calls.notices.push({ type, message }) },
    ...overrides,
  }
  return { files: createWorkspaceFileOperations(services), calls, setCurrent: (value) => { current = { ...current, ...value } } }
}

test('file rename accepts Office names and rejects empty, reserved and unsafe download names', () => {
  assert.equal(workspaceFileNameError('  季度预算 2026.xlsx  '), undefined)
  assert.equal(workspaceFileNameError('项目说明（最终）.docx'), undefined)
  assert.equal(workspaceFileNameError('a'.repeat(255)), undefined)
  for (const name of ['', ' \t ', '.', '..', '...', 'a'.repeat(256), 'a/b', 'a\\b', 'a:b', 'a?b', 'a*b', 'a\nb', '<a>', 'a|b', 'a"b']) assert.ok(workspaceFileNameError(name), name)
})

test('save downloads the captured file and marks its exact savepoint after a document switch', async () => {
  const writer = deferred()
  const app = fixture({ write: () => writer.promise })
  const pending = app.files.exportFile('save')
  app.setCurrent({ kind: 'document', id: 'doc-2', revision: 9, name: 'Other.docx', text: 'new edits' })
  writer.resolve({ bytes: new Uint8Array([1, 2, 3]), warnings: [] })
  assert.equal(await pending, true)
  assert.equal(app.calls.downloads[0].snapshot.fileName, 'Original.xlsx')
  assert.deepEqual(app.calls.downloads[0].snapshot.model, { text: 'original' })
  assert.deepEqual(app.calls.savepoints, [{ kind: 'workbook', unitId: 'book-1', revision: 2, generation: 1 }])
  assert.equal(app.calls.tasks.at(-1), null)
})

test('an Office compatibility warning keeps the source dirty while reporting the downloaded file', async () => {
  const app = fixture({ write: async () => ({ bytes: new Uint8Array([1]), warnings: [{ code: 'UNSUPPORTED_PART', message: '某一嵌入对象未写出' }] }) })
  assert.equal(await app.files.exportFile('save'), true)
  assert.equal(app.calls.downloads.length, 1)
  assert.equal(app.calls.savepoints.length, 0)
  assert.equal(app.calls.notices.at(-1).type, 'info')
  assert.match(app.calls.notices.at(-1).message, /某一嵌入对象未写出/)
})

test('exporting a copy does not change the document savepoint', async () => {
  const app = fixture()
  assert.equal(await app.files.exportFile('export'), true)
  assert.equal(app.calls.downloads.length, 1)
  assert.equal(app.calls.savepoints.length, 0)
  assert.match(app.calls.notices.at(-1).message, /导出文件已下载/)
})

test('download failure leaves the savepoint intact and clears the task state', async () => {
  const app = fixture({ download: () => { throw new Error('浏览器下载被拒绝') } })
  assert.equal(await app.files.exportFile('save'), false)
  assert.equal(app.calls.savepoints.length, 0)
  assert.deepEqual(app.calls.notices.at(-1), { type: 'error', message: '浏览器下载被拒绝' })
  assert.equal(app.calls.tasks.at(-1), null)
})

test('cancelling an export aborts work and suppresses delayed progress and success', async () => {
  const writer = deferred()
  let options
  const app = fixture({ write: (_snapshot, value) => { options = value; return writer.promise } })
  const pending = app.files.exportFile()
  options.onProgress({ message: '正在写出ZIP' })
  assert.equal(app.calls.tasks.at(-1).message, '正在写出ZIP')
  assert.equal(app.files.cancelExport(), true)
  assert.equal(options.signal.aborted, true)
  const taskCount = app.calls.tasks.length
  options.onProgress({ message: '过期进度' })
  writer.resolve({ bytes: new Uint8Array([1]), warnings: [] })
  assert.equal(await pending, false)
  assert.equal(app.calls.tasks.length, taskCount)
  assert.equal(app.calls.downloads.length, 0)
  assert.equal(app.calls.savepoints.length, 0)
  assert.deepEqual(app.calls.notices, [{ type: 'info', message: '导出任务已取消' }])
})

test('a superseded writer cannot download or hide the newer writer task', async () => {
  const older = deferred()
  const newer = deferred()
  const options = []
  const app = fixture({ write: (_snapshot, value) => { options.push(value); return options.length === 1 ? older.promise : newer.promise } })
  const first = app.files.exportFile()
  app.setCurrent({ name: 'Latest.xlsx', revision: 3 })
  const second = app.files.exportFile()
  assert.equal(options[0].signal.aborted, true)
  older.resolve({ bytes: new Uint8Array([1]), warnings: [] })
  assert.equal(await first, false)
  assert.equal(app.calls.tasks.at(-1).fileName, 'Latest.xlsx')
  assert.equal(app.calls.downloads.length, 0)
  newer.resolve({ bytes: new Uint8Array([2]), warnings: [] })
  assert.equal(await second, true)
  assert.equal(app.calls.downloads.length, 1)
  assert.equal(app.calls.downloads[0].snapshot.fileName, 'Latest.xlsx')
})

test('cancelled or unapplied imports cannot produce a misleading success notice', async () => {
  const parse = deferred()
  const app = fixture({ importFile: () => parse.promise })
  const pending = app.files.importFile({ name: 'Cancelled.xlsx' })
  app.files.cancelImport()
  parse.resolve(true)
  assert.equal(await pending, false)
  assert.equal(app.calls.importsCancelled, 1)
  assert.equal(app.calls.notices.length, 0)
  const skipped = fixture({ importFile: async () => false })
  assert.equal(await skipped.files.importFile({ name: 'Superseded.xlsx' }), false)
  assert.equal(skipped.calls.notices.length, 0)
})

test('an older import rejection cannot replace a newer successful import notice', async () => {
  const older = deferred()
  let count = 0
  const app = fixture({ importFile: () => ++count === 1 ? older.promise : Promise.resolve(true) })
  const pending = app.files.importFile({ name: 'Old.xlsx' })
  assert.equal(await app.files.importFile({ name: 'Current.docx' }), true)
  older.reject(new Error('stale error'))
  assert.equal(await pending, false)
  assert.deepEqual(app.calls.notices, [{ type: 'success', message: 'Current.docx 已解析并载入' }])
})

test('disposing aborts both operations without emitting delayed state or success', async () => {
  const writer = deferred()
  let signal
  const app = fixture({ write: (_snapshot, options) => { signal = options.signal; return writer.promise } })
  const pending = app.files.exportFile()
  app.files.dispose()
  app.files.dispose()
  assert.equal(signal.aborted, true)
  writer.resolve({ bytes: new Uint8Array([1]), warnings: [] })
  assert.equal(await pending, false)
  assert.equal(app.calls.importsCancelled, 1)
  assert.equal(app.calls.downloads.length, 0)
  assert.equal(app.calls.notices.length, 0)
  assert.equal(await app.files.exportFile(), false)
  assert.equal(await app.files.importFile({ name: 'After.xlsx' }), false)
})
