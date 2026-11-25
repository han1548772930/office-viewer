import { expect, type Locator, type Page } from '@playwright/test'
import { bootStatisticalWorkbench, openStatisticalSettings, statisticalChartSnapshot } from './chart-statistical-settings'

// P28 workbench fixtures only. Importing this module starts no work.
export async function insertChordChart(page: Page, title: string, extraField = false, themeToken = false): Promise<Locator> {
  await bootStatisticalWorkbench(page)
  await page.evaluate(async args => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI } = await import(resource.name), sheet = officeAPI.getActiveWorkbook()!.getActiveSheet()!
    const table: Array<Array<string | number>> = [
      ['Source', 'Target', 'Weight'], ['B', 'A', 2], ['A', 'B', 3], ['C', 'D', 4], ['D', 'C', 1], ['B', 'B', 9],
    ]
    if (args.extraField) table.forEach((row, index) => row.push(index === 0 ? 'Alternate' : index))
    const source = args.extraField ? 'A1:D6' : 'A1:C6'
    sheet.getRange(source).setValues(table); sheet.getRange('C2:C6').setNumberFormat('0.00')
    const builder = sheet.newChart('chord').setSource(source).setChordMapping({ sourceIndex: 0, targetIndex: 1, valueIndex: 2 })
      .setAbsolutePosition(110, 90).setSize(560, 400).setTitle(args.title).setPalette(['#112233', '#ddeeff', '#228833', '#884422'])
    if (args.themeToken) builder.setChordNode({ color: 'blue.500' })
    await sheet.insertChart(builder.build())
  }, { title, extraField, themeToken })
  const hitbox = page.getByRole('button', { name: title, exact: true })
  await expect(hitbox).toHaveAttribute('data-rendered', 'true')
  await expect(hitbox).toHaveAttribute('aria-busy', 'false')
  return hitbox
}

export async function openChordSettings(page: Page, title: string): Promise<Locator> {
  // ExcelWorkbench captures a stationary hitbox click and calls
  // useSheetChartPanels.openChartEditor; ChartOverlay itself only selects.
  return openStatisticalSettings(page, title)
}
export function chordTooltip(page: Page): Locator {
  return page.getByRole('tooltip').filter({ hasText: /(?:Weight|计数)：/ })
}
export async function chordSourceVersion(hitbox: Locator): Promise<number> {
  return Number(await hitbox.getAttribute('data-source-version'))
}
export interface ChordWorkbenchGeometry {
  layout: {
    frame: { x: number; y: number; width: number; height: number }
    chord: {
      center: { x: number; y: number }; innerRadius: number; outerRadius: number
      nodes: Array<{ start: number; end: number }>
      links: Array<{ sourceStart: number; sourceEnd: number }>
    }
  }
}
export async function chordScreenPoint(hitbox: Locator, snapshot: ChordWorkbenchGeometry,
  kind: 'node' | 'link' = 'node', index = 0): Promise<{ x: number; y: number }> {
  const box = await hitbox.boundingBox()
  if (!box) throw new Error('Chord hitbox is unavailable')
  const { frame, chord } = snapshot.layout
  const mark = kind === 'node' ? chord.nodes[index]! : chord.links[index]!
  const angle = 'start' in mark ? (mark.start + mark.end) / 2 : (mark.sourceStart + mark.sourceEnd) / 2
  const radius = kind === 'node' ? (chord.innerRadius + chord.outerRadius) / 2 : chord.innerRadius * 0.9
  // Facade layout has worksheet coordinates. Remove the frame origin before
  // applying the real hitbox's zoom, scroll and browser CSS transform.
  return { x: box.x + (chord.center.x - frame.x + Math.cos(angle) * radius) / frame.width * box.width,
    y: box.y + (chord.center.y - frame.y + Math.sin(angle) * radius) / frame.height * box.height }
}
export async function hoverChord(page: Page, hitbox: Locator, kind: 'node' | 'link' = 'node', index = 0) {
  const snapshot = await statisticalChartSnapshot(page), point = await chordScreenPoint(hitbox, snapshot, kind, index)
  await page.mouse.move(0, 0); await page.mouse.move(point.x, point.y)
  await expect(chordTooltip(page)).toBeVisible()
  return { snapshot, point }
}
export async function chordHistorySavepoint(page: Page): Promise<string> {
  return page.evaluate(async () => {
    const resource = performance.getEntriesByType('resource').find(entry => /\/src\/runtime\.ts(?:\?|$)/.test(entry.name))!
    const { officeAPI, studio } = await import(resource.name)
    return JSON.stringify(studio.editor.history.captureSavepoint(officeAPI.getActiveWorkbook()!.getId()))
  })
}
export async function setChordAppearance(page: Page, appearance: '深色' | '浅色'): Promise<void> {
  await page.locator('.workspace-settings-trigger').click()
  await page.locator('.workspace-settings-menu').getByRole('menuitem', { name: /外观/ }).click()
  await page.getByRole('menuitem', { name: appearance, exact: true }).click()
}
