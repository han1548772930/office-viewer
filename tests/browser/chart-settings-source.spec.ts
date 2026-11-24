import { expect, test, type Locator, type Page } from '@playwright/test'

// P24 workbench drafts only. No runner, server, browser, pixel, theme/IME,
// Office, Worker/WASM or performance acceptance has run in this phase.
// Each full range has 40,000 samples × two fields = 80,002 source positions.
const SAMPLE_COUNT = 40_000
type ReadKind = 'source' | 'statistics'
type ReadMode = 'hold' | 'memory'
test.setTimeout(120_000)

async function setup(page: Page, title: string, options: { chart?: boolean; twoSources?: boolean; pareto?: boolean } = {}) {
  await page.setViewportSize({ width: 1440, height: 900 })
  await page.goto('/')
  await expect(page.getByRole('grid', { name: '工作表画布' })).toBeVisible()
  await page.evaluate(async ({ title, options, count }) => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const neededRows = count + 4, availableRows = sheet.getMaxRows()
    if (availableRows < neededRows) sheet.insertRowsAfter(availableRows - 1, neededRows - availableRows)
    const samples = Array.from({ length: count }, (_value, index) => [
      options.pareto ? index < 2 ? 'A' : index === 2 ? 'B' : index === 3 ? 'C' : 'zero' : `Source A ${index + 1}`,
      options.pareto ? index === 0 ? 6 : index === 1 ? -1 : index === 2 ? 4 : index === 3 ? 2 : 0
        : index === 1 ? null : index === count - 1 ? 100 : 1,
    ])
    const range = `A1:B${count + 1}`
    sheet.getRange(range).setValues([[options.pareto ? 'Cause' : 'Sample', 'Measurements A'], ...samples])
    if (options.twoSources) sheet.getRange(`D1:E${count + 1}`).setValues([['Sample', 'Measurements B'],
      ...Array.from({ length: count }, (_value, index) => [`Source B ${index + 1}`, index === 1 ? null : index === count - 1 ? 200 : 2])])
    sheet.getRange(range).activate()
    if (options.chart !== false) {
      const builder = sheet.newChart(options.pareto ? 'pareto' : 'boxplot').setSource(range).setCategoryField(0).setValueFields([1])
        .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(title).setLegend({ visible: false })
        .setYAxis({ scaleType: 'linear', min: 0, max: 220 }).setSeries(1, { color: 'green' })
      builder.setOptions(options.pareto ? { invalidValueStrategy: 'zero',
        aggregation: { aggregate: true, topN: { count: 1, includeOthers: true, othersLabel: 'Other causes' } },
        pareto: { includeZeroValues: false } } : { boxplot: { quartileMethod: 'inclusive', whiskerMethod: 'iqr' } })
      await sheet.insertChart(builder.build())
    }
  }, { title, options, count: SAMPLE_COUNT })
  if (options.chart !== false) await expect(page.getByRole('button', { name: title, exact: true })).toHaveAttribute('data-rendered', 'true')
}

/** Hold after the real service has returned to exercise the UI's own stale
 * completion guard, independently of the real provider's cancellation. */
async function trackReads(page: Page) {
  await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const state = { calls: [] as Array<any>, next: undefined as { kind: ReadKind; mode: ReadMode } | undefined, synchronousDataReads: 0 }
    ;(globalThis as any).__chartSettingsReadDraft = state
    const prototype = Object.getPrototypeOf(sheet)
    for (const [kind, method] of [['source', 'getChartSettingsData'], ['statistics', 'getChartSettingsStatistics']] as const) {
      const original = prototype[method]
      prototype[method] = async function(config: any, options: any = {}) {
        const mode = state.next?.kind === kind ? state.next.mode : undefined
        if (mode) state.next = undefined
        const call: any = { kind, phase: 'pending', aborted: options.signal?.aborted ?? false, range: { ...config.source.range }, type: config.type,
          includeZeroValues: config.options?.pareto?.includeZeroValues, pivotTableId: config.pivotTableId }
        state.calls.push(call)
        options.signal?.addEventListener('abort', () => { call.aborted = true }, { once: true })
        try {
          const result = await original.call(this, config, mode === 'memory' ? { ...options, memoryBudgetBytes: 4096 } : options)
          const data = kind === 'source' ? result : result.preparedData
          call.sourceCount = data?.categories.length; call.firstCategory = data?.categories[0]
          call.materializedCategories = kind === 'statistics' && result.pareto?.categories.length < 16
            ? [...result.pareto.categories] : undefined
          if (mode === 'hold') { call.phase = 'held'; await new Promise<void>(resolve => { call.release = resolve }) }
          call.phase = 'returned'
          return result
        } catch (reason) {
          call.phase = 'failed'; call.error = reason instanceof Error ? reason.message : String(reason); call.code = (reason as any)?.code
          throw reason
        }
      }
    }
    const chart = sheet.getCharts()[0]
    if (chart) Object.getPrototypeOf(chart).getData = function() {
      state.synchronousDataReads++
      throw new Error('The chart settings UI must use the asynchronous draft readers')
    }
  })
}
async function armRead(page: Page, kind: ReadKind, mode: ReadMode) {
  await page.evaluate(value => { (globalThis as any).__chartSettingsReadDraft.next = value }, { kind, mode })
}
async function reads(page: Page) {
  return page.evaluate(() => {
    const state = (globalThis as any).__chartSettingsReadDraft
    return { synchronousDataReads: state.synchronousDataReads, calls: state.calls.map(({ release: _release, ...call }: any) => call) }
  })
}
async function releaseRead(page: Page, index: number) {
  await page.evaluate(value => { (globalThis as any).__chartSettingsReadDraft.calls[value].release() }, index)
}
async function waitReady(panel: Locator) { await expect(panel.getByText('正在读取图表数据源…', { exact: true })).toHaveCount(0) }
async function editor(page: Page, title: string) {
  await page.getByRole('button', { name: title, exact: true }).click({ position: { x: 24, y: 18 } })
  const panel = page.getByRole('dialog', { name: '图表设置', exact: true })
  await expect(panel).toBeVisible()
  return panel
}
async function seriesPoints(page: Page, panel: Locator, raw = true) {
  await panel.getByText('系列、标记与数据标签', { exact: true }).click()
  await panel.getByText('单独设置系列', { exact: true }).click()
  const header = raw ? '单个原始样本' : '单个数据项目'
  await panel.getByText(header, { exact: true }).click()
  return panel.locator('.ant-collapse-item').filter({ has: page.getByText(header, { exact: true }) }).last()
}
async function number(panel: Locator, label: string, value: string) {
  const control = panel.getByRole('spinbutton', { name: label, exact: true })
  await control.fill(value); await control.press('Tab')
}
function selected(page: Page, panel: Locator, label: string) {
  return panel.locator('.ant-form-item').filter({ has: page.getByRole('combobox', { name: label, exact: true }) }).last().locator('.ant-select-selection-item')
}
async function choose(page: Page, panel: Locator, label: string, value: string) {
  await panel.getByRole('combobox', { name: label, exact: true }).click()
  await page.getByRole('option', { name: value, exact: true }).click()
}
async function applyAndClose(panel: Locator) {
  await panel.getByRole('button', { name: '应用', exact: true }).click()
  const close = panel.getByRole('button', { name: '关闭', exact: true }).last()
  await expect(close).toBeEnabled(); await expect(panel).toBeVisible()
  await close.click(); await expect(panel).toBeHidden()
}
async function snapshot(page: Page, reopen = false) {
  return page.evaluate(async createCopy => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), current = officeAPI.getActiveWorkbook()!
    const workbook = createCopy ? officeAPI.createWorkbook({ ...current.save(), id: crypto.randomUUID() }, { makeCurrent: false }) : current
    const chart = workbook.getActiveSheet()!.getCharts()[0]!, info = chart.getInfo()
    const [data, layout, commands] = await Promise.all([chart.getComputedData(), chart.getComputedLayout(), chart.getComputedRenderCommands()])
    return { revision: info.revision, source: info.source, pivotTableId: info.pivotTableId, options: chart.getOptions(),
      raw: { count: data.categories.length, first: data.categories[0], last: data.categories.at(-1),
        firstValue: data.series[0]?.values[0], secondValue: data.series[0]?.values[1], lastValue: data.series[0]?.values.at(-1) },
      categories: layout.categories, paint: JSON.stringify(commands) }
  }, reopen)
}
async function history(page: Page, action: 'undo' | 'redo') {
  await page.evaluate(async operation => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    if (operation === 'undo') workbook.undo(); else workbook.redo()
  }, action)
}

test('an over-budget Boxplot source supports its last raw sample, detached edits, one apply, history and internal reopen', async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message))
  const title = 'Boxplot full settings source draft'
  await setup(page, title); const before = await snapshot(page)
  expect(before.raw).toMatchObject({ count: SAMPLE_COUNT, last: 'Source A 40000', secondValue: null, lastValue: 100 })
  await trackReads(page)
  const panel = await editor(page, title); await waitReady(panel)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source A 40000 · 100')
  await number(points, '项目标记大小', '18'); await number(points, '项目填充不透明度', '0.35')
  await choose(page, points, '项目标记形状', '菱形')
  const draft = await snapshot(page); expect(draft.revision).toBe(before.revision); expect(draft.options).toEqual(before.options)
  const read = await reads(page); expect(read.calls.map((call: any) => call.kind)).toEqual(['source']); expect(read.synchronousDataReads).toBe(0)
  await applyAndClose(panel)
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1); expect(applied.raw).toEqual(before.raw)
  expect(applied.options.series![1]!.dataPoints![39999]).toMatchObject({ size: 18, shape: 'diamond', fillOpacity: 0.35 })
  expect(applied.paint).not.toBe(before.paint)
  const restored = await snapshot(page, true)
  expect(restored.options).toEqual(applied.options); expect(restored.raw).toEqual(applied.raw); expect(restored.paint).toBe(applied.paint)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
  await history(page, 'redo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(applied.options)
  expect((await reads(page)).synchronousDataReads).toBe(0); expect(errors).toEqual([])
})

test('a held old source cannot overwrite a later over-budget range binding or its last sample label', async ({ page }) => {
  const title = 'Boxplot changed settings binding draft'
  await setup(page, title, { twoSources: true }); const before = await snapshot(page)
  await trackReads(page); await armRead(page, 'source', 'hold')
  const panel = await editor(page, title)
  await expect.poll(async () => (await reads(page)).calls[0]?.phase).toBe('held')
  await panel.getByRole('textbox', { name: '图表数据区域', exact: true }).fill('D1:E40001')
  await expect.poll(async () => (await reads(page)).calls[0]?.aborted).toBe(true)
  await expect.poll(async () => (await reads(page)).calls.at(-1)?.phase).toBe('returned'); await waitReady(panel)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source B 40000 · 200')
  await releaseRead(page, 0); await expect.poll(async () => (await reads(page)).calls[0]?.phase).toBe('returned')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source B 40000 · 200')
  expect((await reads(page)).calls).toHaveLength(2); expect((await reads(page)).synchronousDataReads).toBe(0)
  await panel.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(panel).toBeHidden()
  expect((await snapshot(page)).source).toEqual(before.source); expect((await snapshot(page)).revision).toBe(before.revision)
})

test('closing cancels a pending full-source read and reopening starts a fresh read before the old provider returns', async ({ page }) => {
  const title = 'Boxplot source close and reopen draft'
  await setup(page, title); await trackReads(page); await armRead(page, 'source', 'hold')
  let panel = await editor(page, title)
  await expect.poll(async () => (await reads(page)).calls[0]?.phase).toBe('held')
  await expect(panel.getByRole('button', { name: '应用', exact: true })).toBeDisabled()
  await panel.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(panel).toBeHidden()
  await expect.poll(async () => (await reads(page)).calls[0]?.aborted).toBe(true)
  panel = await editor(page, title); await waitReady(panel)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source A 40000 · 100')
  await releaseRead(page, 0); await expect.poll(async () => (await reads(page)).calls[0]?.phase).toBe('returned')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source A 40000 · 100')
  expect((await reads(page)).calls).toHaveLength(2); expect((await reads(page)).synchronousDataReads).toBe(0)
  await panel.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(panel).toBeHidden()
})

test('a real source memory rejection is visible and retry loads the same untouched author draft', async ({ page }) => {
  const title = 'Boxplot source memory retry draft'
  await setup(page, title); const before = await snapshot(page)
  await trackReads(page); await armRead(page, 'source', 'memory')
  const panel = await editor(page, title)
  await expect.poll(async () => (await reads(page)).calls[0]?.phase).toBe('failed')
  const failed = (await reads(page)).calls[0]
  expect(failed.code).toBe('COMPUTE_MEMORY_LIMIT'); await expect(panel.getByText(failed.error, { exact: true }).first()).toBeVisible()
  expect((await snapshot(page)).options).toEqual(before.options); expect((await snapshot(page)).revision).toBe(before.revision)
  await panel.getByRole('button', { name: '重新读取数据源', exact: true }).click()
  await expect.poll(async () => (await reads(page)).calls.at(-1)?.phase).toBe('returned'); await waitReady(panel)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source A 40000 · 100')
  expect((await reads(page)).calls).toHaveLength(2); expect((await reads(page)).synchronousDataReads).toBe(0)
  await panel.getByRole('button', { name: '关闭', exact: true }).last().click(); await expect(panel).toBeHidden()
  expect((await snapshot(page)).options).toEqual(before.options)
})

test('the insert modal reads a full source and stores the last raw-sample style before any FChart exists', async ({ page }) => {
  await setup(page, 'Unused insertion fixture title', { chart: false }); await trackReads(page)
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '插入', exact: true }).click()
  await page.getByRole('button', { name: '图表', exact: true }).click()
  const panel = page.getByRole('dialog', { name: '插入图表', exact: true }); await expect(panel).toBeVisible()
  await choose(page, panel, '图表类型', '箱线图')
  await panel.getByRole('textbox', { name: '图表数据区域', exact: true }).fill('A1:B40001')
  await waitReady(panel)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000'); await number(points, '项目标记大小', '17')
  await expect(selected(page, points, '选择数据项目')).toHaveText('40000 · Source A 40000 · 100')
  await expect.poll(async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    return officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts().length
  })).toBe(0)
  await panel.getByRole('button', { name: '插入', exact: true }).click(); await expect(panel).toBeHidden()
  const applied = await snapshot(page)
  expect(applied.raw.count).toBe(SAMPLE_COUNT); expect(applied.options.series![1]!.dataPoints![39999]!.size).toBe(17)
  const copy = await snapshot(page, true); expect(copy.raw).toEqual(applied.raw); expect(copy.options).toEqual(applied.options)
  await history(page, 'undo')
  await expect.poll(async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name)
    return officeAPI.getActiveWorkbook()!.getActiveSheet()!.getCharts().length
  })).toBe(0)
})

test('Pareto panels share one Rust materialization and reject held zero/Top N statistics after the draft changes', async ({ page }) => {
  const title = 'Pareto settings materialization draft'
  await setup(page, title, { pareto: true }); const before = await snapshot(page)
  expect(before.categories).toEqual(['Other causes', 'A'])
  await trackReads(page)
  const panel = await editor(page, title); await waitReady(panel)
  expect((await reads(page)).calls.map((call: any) => call.kind)).toEqual(['statistics'])
  await panel.getByText('图型专属设置', { exact: true }).click()
  await panel.getByText('帕累托柱形样式', { exact: true }).click(); await panel.getByText('帕累托累计线样式', { exact: true }).click()
  const bar = panel.locator('.chart-pareto-settings > .ant-collapse-item').filter({ has: page.getByText('帕累托柱形样式', { exact: true }) })
  const line = panel.locator('.chart-pareto-settings > .ant-collapse-item').filter({ has: page.getByText('帕累托累计线样式', { exact: true }) })
  await number(bar, '选择帕累托柱形数据项目序号', '1'); await number(line, '选择帕累托累计线数据项目序号', '1')
  await expect(selected(page, bar, '选择帕累托柱形数据项目')).toHaveText('1 · Other causes')
  await expect(selected(page, line, '选择帕累托累计线数据项目')).toHaveText('1 · Other causes')
  const points = await seriesPoints(page, panel, false)
  await number(points, '数据项目序号', '1'); await expect(selected(page, points, '选择数据项目')).toHaveText('1 · Other causes')
  await number(bar, '柱形项目不透明度', '0.4'); await number(line, '累计线项目标记大小', '12')
  expect((await reads(page)).calls).toHaveLength(1); expect((await snapshot(page)).options).toEqual(before.options)
  await armRead(page, 'statistics', 'hold'); await panel.getByRole('switch', { name: '保留零权重分类', exact: true }).click()
  await expect.poll(async () => (await reads(page)).calls[1]?.phase).toBe('held')
  await panel.getByText('分类聚合与前 N 项', { exact: true }).click()
  await panel.getByRole('switch', { name: '显示前 N 项', exact: true }).click()
  await expect.poll(async () => (await reads(page)).calls[1]?.aborted).toBe(true)
  await expect.poll(async () => (await reads(page)).calls[2]?.phase).toBe('returned'); await waitReady(panel)
  await number(bar, '选择帕累托柱形数据项目序号', '4')
  await expect(selected(page, bar, '选择帕累托柱形数据项目')).toHaveText('4 · zero')
  await releaseRead(page, 1); await expect.poll(async () => (await reads(page)).calls[1]?.phase).toBe('returned')
  await expect(selected(page, bar, '选择帕累托柱形数据项目')).toHaveText('4 · zero')
  expect((await reads(page)).calls.every((call: any) => call.kind === 'statistics')).toBe(true)
  expect((await reads(page)).synchronousDataReads).toBe(0)
  await applyAndClose(panel)
  const applied = await snapshot(page)
  expect(applied.revision).toBe(before.revision + 1); expect(applied.categories).toEqual(['A', 'B', 'C', 'zero'])
  expect(applied.options.pareto!.barStyle!.dataPoints![0]!.fillOpacity).toBe(0.4)
  expect(applied.options.pareto!.cumulativeLineStyle!.dataPoints![0]!.size).toBe(12)
  const copy = await snapshot(page, true); expect(copy.options).toEqual(applied.options); expect(copy.paint).toBe(applied.paint)
  await history(page, 'undo'); await expect.poll(async () => (await snapshot(page)).options).toEqual(before.options)
})

test('the actual PivotChart panel loads a large pivot output and persists a far-sample draft with its pivot association', async ({ page }) => {
  await setup(page, 'Unused pivot fixture title', { chart: false })
  const pivotId = await page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!, sheet = workbook.getActiveSheet()!
    const pivot = await workbook.addPivotTable({ unitId: workbook.getId(), subUnitId: sheet.getSheetId(), sheetName: sheet.getSheetName(),
      range: { startRow: 0, startColumn: 0, endRow: 40_000, endColumn: 1 } }, 'existing',
    { unitId: workbook.getId(), subUnitId: sheet.getSheetId(), row: 0, col: 3 })
    await pivot.addField(0, 'row'); await pivot.addField(1, 'value')
    return pivot.getPivotTableId()
  })
  await trackReads(page)
  await page.getByTestId('workbench-ribbon').getByRole('tab', { name: '数据', exact: true }).click()
  const open = page.getByRole('button', { name: '透视图', exact: true }); await expect(open).toBeEnabled(); await open.click()
  const panel = page.getByRole('dialog', { name: '透视图', exact: true }); await expect(panel).toBeVisible()
  await choose(page, panel, '图表类型', '箱线图'); await waitReady(panel)
  await expect.poll(async () => (await reads(page)).calls.at(-1)?.sourceCount).toBeGreaterThanOrEqual(SAMPLE_COUNT)
  const points = await seriesPoints(page, panel)
  await number(points, '原始样本序号', '40000')
  await expect(selected(page, points, '选择数据项目')).toContainText('40000 · Source A')
  await number(points, '项目标记大小', '17')
  await panel.getByRole('button', { name: '创建透视图', exact: true }).click()
  await expect(panel.getByText('已创建并关联透视图。', { exact: true })).toBeVisible()
  await expect(panel.getByRole('button', { name: '关闭', exact: true })).toBeEnabled()
  await panel.getByRole('button', { name: '关闭', exact: true }).click(); await expect(panel).toBeHidden()
  const applied = await snapshot(page)
  expect(applied.pivotTableId).toBe(pivotId)
  expect(Object.values(applied.options.series ?? {}).some(style => style.dataPoints?.[39999]?.size === 17)).toBe(true)
  const copy = await snapshot(page, true); expect(copy.pivotTableId).toBe(pivotId); expect(copy.options).toEqual(applied.options)
  await history(page, 'undo')
  await expect.poll(async () => page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), workbook = officeAPI.getActiveWorkbook()!
    return { charts: workbook.getActiveSheet()!.getCharts().length, pivots: workbook.getPivotTables().length }
  })).toEqual({ charts: 0, pivots: 1 })
})
