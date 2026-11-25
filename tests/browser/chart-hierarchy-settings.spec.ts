import { expect, test } from '@playwright/test'
import { insertHierarchyChart } from './helpers/chart-hierarchy-settings'
import { applyStatisticalSettings, chartSettingsPanel, chooseChartValue, openStatisticalSettings,
  selectedChartValue, statisticalChartHistory, statisticalChartSnapshot } from './helpers/chart-statistical-settings'

// Shared ordered mapping workbench drafts. No browser/Worker/history check ran.
for (const type of ['treemap', 'sunburst'] as const) {
  test(`${type} ordered add/replace/move/remove mapping stays local until one Apply and survives history/reopen`, async ({ page }) => {
    const title = `${type} ordered hierarchy draft`; await insertHierarchyChart(page, title, type, true)
    const before = await statisticalChartSnapshot(page), editor = await openStatisticalSettings(page, title), data = chartSettingsPanel(page, editor, '数据源与字段')
    await expect(selectedChartValue(page, data, '层级 1 字段')).toHaveText('Region')
    await expect(selectedChartValue(page, data, '层级 2 字段')).toHaveText('Category')
    await expect(selectedChartValue(page, data, '权重字段')).toHaveText('Weight')
    await data.getByRole('combobox', { name: '添加层级字段', exact: true }).click()
    await expect(page.getByRole('option', { name: 'Weight', exact: true })).toHaveCount(0)
    await page.getByRole('option', { name: 'Bucket', exact: true }).click()
    await data.getByRole('button', { name: '添加层级', exact: true }).click()
    await expect(selectedChartValue(page, data, '层级 3 字段')).toHaveText('Bucket')
    await data.getByRole('button', { name: '上移层级 3', exact: true }).click()
    await expect(selectedChartValue(page, data, '层级 2 字段')).toHaveText('Bucket')
    await chooseChartValue(page, data, '层级 1 字段', 'Alternate')
    await data.getByRole('button', { name: '移除层级 3', exact: true }).click()
    await data.getByRole('button', { name: '移除层级 2', exact: true }).click()
    await expect(data.getByRole('button', { name: '移除层级 1', exact: true })).toBeDisabled()
    await chooseChartValue(page, data, '添加层级字段', 'Bucket'); await data.getByRole('button', { name: '添加层级', exact: true }).click()
    await chooseChartValue(page, data, '添加层级字段', 'Category'); await data.getByRole('button', { name: '添加层级', exact: true }).click()
    expect((await statisticalChartSnapshot(page)).contents).toBe(before.contents)
    await applyStatisticalSettings(editor)
    const applied = await statisticalChartSnapshot(page)
    expect(applied.revision).toBe(before.revision + 1); expect(applied.options.mapping.hierarchyIndexes).toEqual([4, 3, 1]); expect(applied.valueFields).toEqual([2])
    expect(applied.raw.hierarchy[0].name).toBe('Alternate Root')
    expect(await statisticalChartSnapshot(page, true)).toMatchObject({ options: applied.options, raw: applied.raw, layout: applied.layout, paint: applied.paint })
    await statisticalChartHistory(page, 'undo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(before.options)
    await statisticalChartHistory(page, 'redo'); await expect.poll(async () => (await statisticalChartSnapshot(page)).options).toEqual(applied.options)
    const closed = await openStatisticalSettings(page, title), closedData = chartSettingsPanel(page, closed, '数据源与字段')
    await closedData.getByRole('button', { name: '下移层级 1', exact: true }).click()
    await closed.getByRole('button', { name: '关闭', exact: true }).last().click()
    expect((await statisticalChartSnapshot(page)).contents).toBe(applied.contents)
    await expect(page.locator('.chart-render-error')).toHaveCount(0)
  })
}
