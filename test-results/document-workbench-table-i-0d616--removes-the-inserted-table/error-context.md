# Instructions

- Following Playwright test failed.
- Explain why, be concise, respect Playwright best practices.
- Provide a snippet of code with the fix, if possible.

# Test info

- Name: document-workbench.spec.ts >> table insertion modal applies dimensions and Ctrl+Z removes the inserted table
- Location: tests\browser\document-workbench.spec.ts:61:1

# Error details

```
Test timeout of 60000ms exceeded.
```

```
Error: locator.click: Test timeout of 60000ms exceeded.
Call log:
  - waiting for locator('.ant-modal').locator('.ant-modal-footer button').filter({ hasText: '确定' })

```

# Page snapshot

```yaml
- generic [ref=e1]:
  - generic [ref=e3]:
    - banner [ref=e4]:
      - link "Office Studio 工作台" [ref=e5] [cursor=pointer]:
        - /url: "#"
        - generic [ref=e9]: Office Studio
      - navigation "工作台" [ref=e10]:
        - button "表格" [ref=e11] [cursor=pointer]
        - button "文档" [ref=e15] [cursor=pointer]
      - generic "未命名文档" [ref=e20]
      - generic [ref=e21]:
        - button "新建工作簿" [ref=e22] [cursor=pointer]:
          - generic [ref=e24]: 新建
        - button "打开 Office 文件" [ref=e25] [cursor=pointer]:
          - generic [ref=e28]: 打开
        - button "撤销" [disabled] [ref=e30]
        - button "重做" [disabled] [ref=e31]
        - button "导出 Office 文件" [ref=e32] [cursor=pointer]:
          - generic [ref=e36]: 导出
        - button "保存 Office 文件" [ref=e37] [cursor=pointer]:
          - generic [ref=e42]: 保存
    - main "文档编辑工作区" [ref=e43]:
      - generic [ref=e44]:
        - generic [ref=e45]:
          - generic [ref=e50]:
            - textbox "文档名称" [ref=e51]: 未命名文档
            - generic [ref=e52]: 已保存
          - generic [ref=e56]:
            - button "打印文档" [ref=e57] [cursor=pointer]
            - button "保存" [ref=e62] [cursor=pointer]
            - button "打开文档" [ref=e68] [cursor=pointer]
        - navigation "文档工具栏" [ref=e74]:
          - radiogroup "segmented control" [ref=e75]:
            - generic [ref=e76]:
              - generic [ref=e77] [cursor=pointer]:
                - radio "开始"
                - generic "开始" [ref=e78]
              - generic [ref=e79] [cursor=pointer]:
                - radio "插入" [checked]
                - generic "插入" [ref=e80]
              - generic [ref=e81] [cursor=pointer]:
                - radio "布局"
                - generic "布局" [ref=e82]
              - generic [ref=e83] [cursor=pointer]:
                - radio "视图"
                - generic "视图" [ref=e84]
          - textbox "查找内容" [ref=e89]
        - generic [ref=e90]:
          - button "表格" [ref=e91] [cursor=pointer]
          - button "图片" [ref=e95] [cursor=pointer]
          - button "链接" [ref=e101] [cursor=pointer]
          - button "分页符" [ref=e106] [cursor=pointer]
          - button "分隔线" [ref=e109] [cursor=pointer]
          - button "内容控件" [ref=e112] [cursor=pointer]
        - main [ref=e117]:
          - complementary [ref=e118]:
            - generic [ref=e119]: 文档大纲
            - paragraph [ref=e120]: 文档中暂无标题
            - generic [ref=e121]: 页面
            - button "第 1 页" [ref=e122] [cursor=pointer]
            - button "新建页面" [ref=e128] [cursor=pointer]
          - generic [ref=e131]:
            - textbox "文档画布编辑区" [ref=e132]
            - textbox "文档文字输入" [ref=e133]
            - generic [ref=e134]: webgpu
        - generic [ref=e135]:
          - generic [ref=e136]:
            - generic [ref=e137]: 第 1 页，共 1 页
            - generic [ref=e139]: 0 个字符
          - generic [ref=e140]:
            - button "上一页" [disabled] [ref=e141]
            - button "下一页" [disabled] [ref=e142]
            - button "缩小" [ref=e144] [cursor=pointer]
            - slider "缩放比例" [ref=e146]: "100"
            - button "放大" [ref=e147] [cursor=pointer]
            - generic [ref=e149]: 100%
  - generic [ref=e150]:
    - dialog "插入表格":
      - generic [ref=e151]:
        - button "Close" [ref=e152] [cursor=pointer]:
          - generic "Close" [ref=e153]:
            - img "close" [ref=e154]
        - generic [ref=e157]: 插入表格
        - generic [ref=e160]:
          - generic [ref=e161]:
            - text: 行数
            - generic [ref=e162]:
              - spinbutton "行数 Increase Value Decrease Value" [ref=e163]: "2"
              - generic:
                - button "Increase Value" [ref=e164] [cursor=pointer]:
                  - img "up" [ref=e165]
                - button "Decrease Value" [ref=e168] [cursor=pointer]:
                  - img "down" [ref=e169]
          - generic [ref=e172]:
            - text: 列数
            - generic [ref=e173]:
              - spinbutton "列数 Increase Value Decrease Value" [active] [ref=e174]: "3"
              - generic:
                - button "Increase Value" [ref=e175] [cursor=pointer]:
                  - img "up" [ref=e176]
                - button "Decrease Value" [ref=e179] [cursor=pointer]:
                  - img "down" [ref=e180]
        - generic [ref=e183]:
          - button "取 消" [ref=e184] [cursor=pointer]
          - button "确 定" [ref=e186] [cursor=pointer]
```

# Test source

```ts
  1  | import { expect, test } from '@playwright/test'
  2  | 
  3  | type WorkbenchSnapshot = {
  4  |   text: string
  5  |   fontFamilies: string[]
  6  |   blocks: string[]
  7  |   controls: Array<{ type: string; startOffset: number; endOffset: number }>
  8  | }
  9  | 
  10 | async function snapshot(page: import('@playwright/test').Page): Promise<WorkbenchSnapshot> {
  11 |   return page.evaluate(() => {
  12 |     const root = document.querySelector<HTMLElement>('.document-workbench')
  13 |     const state = root?.__vueParentComponent?.setupState as { model?: { blocks?: Array<{ type: string; runs?: Array<{ text: string; fontFamily: string }> }>; contentControls?: Array<{ type: string; startOffset: number; endOffset: number }> } } | undefined
  14 |     const model = state?.model
  15 |     const runs = model?.blocks?.flatMap((block) => block.runs ?? []) ?? []
  16 |     return {
  17 |       text: runs.map((run) => run.text).join(''),
  18 |       fontFamilies: runs.map((run) => run.fontFamily),
  19 |       blocks: model?.blocks?.map((block) => block.type) ?? [],
  20 |       controls: model?.contentControls?.map(({ type, startOffset, endOffset }) => ({ type, startOffset, endOffset })) ?? [],
  21 |     }
  22 |   })
  23 | }
  24 | 
  25 | test.beforeEach(async ({ page }) => {
  26 |   await page.goto('/')
  27 |   await page.locator('.workspace-tabs .workspace-tab').nth(1).click()
  28 |   await expect(page.locator('.document-workbench')).toBeVisible()
  29 |   await page.getByRole('textbox', { name: '文档画布编辑区' }).waitFor()
  30 | })
  31 | 
  32 | test('an antdv-next font selection retains the Canvas text range', async ({ page }) => {
  33 |   const bridge = page.getByLabel('文档文字输入')
  34 |   await bridge.focus()
  35 |   await bridge.pressSequentially('选区保留')
  36 |   await bridge.press('Control+A')
  37 |   const font = page.locator('.document-workbench .font-select')
  38 |   await font.click()
  39 |   await page.locator('.ant-select-dropdown:visible .ant-select-item').filter({ hasText: 'Georgia' }).click()
  40 |   const result = await snapshot(page)
  41 |   expect(result.text).toBe('选区保留')
  42 |   expect(result.fontFamilies).toContain('Georgia')
  43 | })
  44 | 
  45 | test('Canvas input bridge commits Chinese composition as one undoable edit', async ({ page }) => {
  46 |   const bridge = page.getByLabel('文档文字输入')
  47 |   await bridge.focus()
  48 |   await bridge.evaluate((element) => {
  49 |     element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }))
  50 |     element.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: '中文输入' }))
  51 |     element.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '中文输入' }))
  52 |   })
  53 |   await expect.poll(async () => (await snapshot(page)).text).toBe('中文输入')
  54 |   await bridge.focus()
  55 |   await bridge.press('Control+Z')
  56 |   await expect.poll(async () => (await snapshot(page)).text).toBe('')
  57 |   await page.locator('.document-workbench .tool-button[aria-label="重做"]').click()
  58 |   await expect.poll(async () => (await snapshot(page)).text).toBe('中文输入')
  59 | })
  60 | 
  61 | test('table insertion modal applies dimensions and Ctrl+Z removes the inserted table', async ({ page }) => {
  62 |   await page.locator('.ribbon-segmented .ant-segmented-item').filter({ hasText: '插入' }).click()
  63 |   await page.locator('main[aria-label="文档编辑工作区"] .ribbon-action').filter({ hasText: '表格' }).click()
  64 |   const dialog = page.locator('.ant-modal')
  65 |   await expect(dialog).toBeVisible()
  66 |   await dialog.getByRole('spinbutton').nth(0).fill('2')
  67 |   await dialog.getByRole('spinbutton').nth(1).fill('3')
> 68 |   await dialog.locator('.ant-modal-footer button').filter({ hasText: '确定' }).click()
     |                                                                              ^ Error: locator.click: Test timeout of 60000ms exceeded.
  69 |   await expect(dialog).toBeHidden()
  70 |   await expect.poll(async () => (await snapshot(page)).blocks).toContain('table')
  71 |   await page.locator('.document-workbench .tool-button[aria-label="撤销"]').click()
  72 |   await expect.poll(async () => (await snapshot(page)).blocks).not.toContain('table')
  73 | })
  74 | 
```