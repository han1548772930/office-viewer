import { expect, type Page } from '@playwright/test'
import { bootStatisticalWorkbench } from './chart-statistical-settings'

// Shared signed source/authoring for P26/P27 browser drafts. Import runs nothing.
export async function insertHierarchyChart(page: Page, title: string, type: 'treemap' | 'sunburst', extraFields = false) {
  await bootStatisticalWorkbench(page)
  await page.evaluate(async args => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const table: Array<Array<string | number | null>> = [
      ['Region', 'Category', 'Weight'], ['South', 'Leaf A', 8], ['North', 'Leaf B', 12], [null, 'Leaf C', 6],
      ['South', null, 3], ['South', 'Leaf A', -2], ['South', 'Dropped', -4], ['North', null, -5],
    ]
    if (args.extraFields) table.forEach((row, index) => row.push(...(index === 0 ? ['Bucket', 'Alternate'] : ['Bucket One', 'Alternate Root'])))
    const source = args.extraFields ? 'A1:E8' : 'A1:C8'
    sheet.getRange(source).setValues(table); sheet.getRange('C2:C8').setNumberFormat('0.00')
    const builder = sheet.newChart(args.type).setSource(source).setHierarchyFields([0, 1]).setValueField(2)
      .setAbsolutePosition(110, 90).setSize(560, 420).setTitle(args.title).setPalette(['#112233', '#ddeeff'])
    if (args.type === 'sunburst') builder.setSunburstLabel({ visible: false }).resetSunburstLabel().setCenterLabel({ visible: true }).resetCenterLabel()
    await sheet.insertChart(builder.build())
  }, { title, type, extraFields })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  return hitbox
}
