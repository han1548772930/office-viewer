import type { SheetRangeThemeStyle } from '../../domain/model.ts'

// Fixed Univer v1.6.0 range-themes/default.ts and build-in-theme.factory.ts color/style catalog.
const colors = ['blue', 'grey', 'red', 'orange', 'yellow', 'green', 'azure', 'indigo', 'purple', 'magenta'] as const
const rgb = (value: string) => `rgb(${value.split(',').join(', ')})`
const fill = (value: string) => ({ fill: rgb(value) })
const light = [
  ['164,202,254', '225,239,254'], ['205,208,216', '238,239,241'], ['248,180,180', '253,232,232'],
  ['253,186,140', '254,236,220'], ['250,200,21', '255,244,185'], ['132,225,188', '222,247,236'],
  ['126,220,226', '213,245,246'], ['186,198,248', '233,237,255'], ['202,191,253', '237,235,254'], ['248,180,217', '252,232,243'],
] as const
const middle = [
  ['63,131,248', '195,221,253'], ['95,101,116', '227,229,234'], ['240,82,82', '251,213,213'],
  ['255,90,31', '252,217,189'], ['212,157,15', '252,220,106'], ['13,164,113', '188,240,218'],
  ['6,148,162', '175,236,239'], ['70,106,247', '210,218,250'], ['144,97,249', '220,215,254'], ['231,70,148', '250,209,232'],
] as const
const dark = [
  ['30,66,159', '195,221,253', '118,169,250'], ['44,48,64', '227,229,234', '151,157,172'],
  ['155,28,28', '251,213,213', '249,128,128'], ['180,52,3', '252,217,189', '255,138,76'],
  ['154,109,21', '252,220,106', '212,157,15'], ['4,108,78', '188,240,218', '49,196,141'],
  ['3,102,114', '175,236,239', '22,189,202'], ['16,51,191', '210,218,250', '98,128,249'],
  ['74,29,150', '220,215,254', '172,148,250'], ['153,21,75', '250,209,232', '241,126,184'],
] as const
const defaultTheme: SheetRangeThemeStyle = {
  name: 'default', headerRowStyle: { fill: 'rgb(68,114,196)', font: { color: 'rgb(255,255,255)', bold: true }, horizontal: 'center' },
  firstRowStyle: { fill: 'rgb(217,225,242)' },
}
const catalog: SheetRangeThemeStyle[] = [defaultTheme,
  { ...defaultTheme, name: 'default-last-row', lastRowStyle: { font: { bold: true }, horizontal: 'center',
    border: { top: { style: 'thin', color: 'rgb(68,114,196)' } } } },
  ...light.map(([header, color], index) => ({ name: `light-${colors[index]}`, headerRowStyle: fill(header),
    firstColumnStyle: fill('255,255,255'), secondColumnStyle: fill(color), lastRowStyle: fill(header) })),
  ...middle.map(([rowHeader, colHeader], index) => ({ name: `middle-${colors[index]}`, headerRowStyle: fill(rowHeader),
    headerColumnStyle: fill(colHeader), secondRowStyle: fill(colHeader), lastRowStyle: fill(rowHeader), lastColumnStyle: fill(colHeader) })),
  ...dark.map(([rowHeader, firstRow, secondRow], index) => ({ name: `dark-${colors[index]}`,
    headerRowStyle: { ...fill(rowHeader), font: { color: rgb('255,255,255'), bold: true }, horizontal: 'center' as const },
    firstRowStyle: fill(firstRow), secondRowStyle: fill(secondRow), lastRowStyle: fill(rowHeader) })),
]
function freeze<T>(value: T): T {
  if (value && typeof value === 'object') { for (const child of Object.values(value)) freeze(child); Object.freeze(value) }
  return value
}
export const SHEET_BUILTIN_RANGE_THEMES: Readonly<Record<string, SheetRangeThemeStyle>> = freeze(Object.fromEntries(catalog.map(style => [style.name, style])))
export const SHEET_BUILTIN_RANGE_THEME_NAMES: readonly string[] = Object.freeze(catalog.map(style => style.name))
