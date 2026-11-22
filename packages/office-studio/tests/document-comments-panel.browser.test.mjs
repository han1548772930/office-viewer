import assert from 'node:assert/strict'
import { afterAll, beforeAll, beforeEach, test } from 'vitest'
import { chromium, expect } from '@playwright/test'
import vue from '@vitejs/plugin-vue'
import { createServer } from 'vite'

const fixturePath = '/__document-comments-panel-test__'
const fixture = `<!doctype html><html lang="en"><head><meta charset="utf-8"><style>body{margin:0}#app{height:100vh}</style></head><body><div id="app"></div><script type="module">
import { createApp, h, ref } from 'vue'
import { ConfigProvider } from 'antdv-next'
import DocumentCommentsPanel from '/packages/office-studio/src/components/document/DocumentCommentsPanel.vue'
const comment = { id: 'comment-1', anchor: [], referenceText: 'Selected text', messages: [{ id: 'message-1', authorId: 'alice', authorName: 'Alice', text: 'Original comment', createdAt: '2026-11-21T00:00:00.000Z' }], resolved: false, createdAt: '2026-11-21T00:00:00.000Z', updatedAt: '2026-11-21T00:00:00.000Z' }
const comments = ref({ [comment.id]: comment })
const createResult = ref(false)
const replyResult = ref(false)
const editResult = ref(false)
window.commentPanelFixture = { comments, createResult, replyResult, editResult, calls: [] }
createApp({ render: () => h(ConfigProvider, null, { default: () => h(DocumentCommentsPanel, {
  comments: comments.value, currentUserId: 'alice', canComment: true, canCreate: true, canStartCreate: true,
  selectedCommentId: 'comment-1',
  onSelect: (commentId) => window.commentPanelFixture.calls.push(['select', commentId]),
  onCreateComment: async (text) => { window.commentPanelFixture.calls.push(['create', text]); if (!createResult.value) return false; comments.value = { ...comments.value, 'comment-2': { ...comment, id: 'comment-2', messages: [{ ...comment.messages[0], id: 'message-2', text }] } }; return true },
  onReplyComment: async (commentId, text, mentions) => { window.commentPanelFixture.calls.push(['reply', commentId, text, mentions]); if (!replyResult.value) return false; comments.value = { ...comments.value, [commentId]: { ...comments.value[commentId], messages: [...comments.value[commentId].messages, { id: 'message-3', authorId: 'alice', authorName: 'Alice', text, createdAt: '2026-11-21T00:01:00.000Z', mentions }] } }; return true },
  onEditComment: async (commentId, messageId, text) => { window.commentPanelFixture.calls.push(['edit', commentId, messageId, text]); if (!editResult.value) return false; comments.value = { ...comments.value, [commentId]: { ...comments.value[commentId], messages: comments.value[commentId].messages.map(message => message.id === messageId ? { ...message, text } : message) } }; return true },
}) }) }).mount('#app')
</script></body></html>`

let server
let browser
let page
let baseURL
const errors = []

beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: process.cwd(),
    plugins: [vue(), {
      name: 'document-comments-panel-fixture',
      configureServer(vite) {
        vite.middlewares.use(fixturePath, async (_request, response, next) => {
          try {
            response.setHeader('Content-Type', 'text/html; charset=utf-8')
            response.end(await vite.transformIndexHtml(fixturePath, fixture))
          } catch (error) {
            next(error)
          }
        })
      },
    }],
    server: { host: '127.0.0.1', port: 0 },
    logLevel: 'error',
  })
  await server.listen()
  const address = server.httpServer.address()
  if (!address || typeof address === 'string') throw new Error('Vite did not allocate an HTTP port')
  baseURL = `http://127.0.0.1:${address.port}${fixturePath}`
  browser = await chromium.launch({ channel: 'msedge', headless: true })
}, 60_000)

afterAll(async () => {
  await page?.close()
  await browser?.close()
  await server?.close()
})

beforeEach(async () => {
  errors.length = 0
  page = await browser.newPage({ viewport: { width: 520, height: 800 } })
  page.on('pageerror', error => errors.push(error.message))
  await page.goto(baseURL)
  await expect(page.getByRole('heading', { name: 'Comments', exact: true })).toBeVisible()
})

test('create, reply and edit keep drafts after failed mutations and clear them after success', async () => {
  const create = page.getByRole('textbox', { name: 'Add a comment', exact: true })
  await create.fill('New comment')
  await create.press('Control+Enter')
  await expect(create).toHaveValue('New comment')
  await page.evaluate(() => { window.commentPanelFixture.createResult.value = true })
  await create.press('Control+Enter')
  await expect(create).toHaveValue('')

  await page.getByRole('button', { name: 'Reply', exact: true }).first().click()
  const reply = page.getByRole('textbox', { name: 'Reply', exact: true })
  await reply.fill('Reply text')
  await reply.press('Control+Enter')
  await expect(reply).toHaveValue('Reply text')
  await page.evaluate(() => { window.commentPanelFixture.replyResult.value = true })
  await reply.press('Control+Enter')
  await expect(reply).toHaveValue('')

  await page.getByRole('button', { name: 'More', exact: true }).first().click()
  await page.getByRole('menuitem', { name: 'Edit', exact: true }).click()
  const edit = page.getByRole('textbox', { name: 'Edit comment', exact: true })
  await edit.fill('Edited comment')
  await edit.press('Control+Enter')
  await expect(edit).toHaveValue('Edited comment')
  await page.evaluate(() => { window.commentPanelFixture.editResult.value = true })
  await edit.press('Control+Enter')
  await expect(edit).toHaveCount(0)
  await expect(page.getByText('Edited comment', { exact: true })).toBeVisible()
  assert.deepEqual(await page.evaluate(() => window.commentPanelFixture.calls.map(([kind]) => kind)), ['create', 'create', 'select', 'reply', 'reply', 'edit', 'edit'])
  assert.deepEqual(errors, [])
})

test('comment threads use antdv-next Card while retaining the source card activation behavior', async () => {
  const thread = page.getByRole('article').filter({ hasText: 'Original comment' })
  await expect(thread).toHaveClass(/ant-card/)
  await expect(thread).toHaveAttribute('id', 'document-comment-comment-1')
  await page.getByText('Original comment', { exact: true }).click()
  assert.deepEqual(await page.evaluate(() => window.commentPanelFixture.calls), [['select', 'comment-1']])
  assert.deepEqual(errors, [])
})
