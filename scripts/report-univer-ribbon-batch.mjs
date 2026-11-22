// Run from the repository root. Read the OSS schema and Pro registration audits; never load a Pro SDK.
import fs from 'node:fs'
import path from 'node:path'
import { execFileSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'

const root = path.resolve(process.argv[2] ?? 'E:/rust_test/univer')
const cache = process.argv[3] ?? '.research.local/univer-1.6.0/packages'
const audit = (...args) => JSON.parse(execFileSync(process.execPath, args, { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }))
const oss = audit('scripts/audit-univer-oss-ribbon.mjs', root)
if (oss.reference.commit !== 'eaf4cfa04fabdbcfc922c692f3450459f8a29546' || oss.reference.tag !== 'v1.6.0') throw Error('Expected fixed Univer v1.6.0 checkout')
const pro = audit('scripts/count-univer-pro-ribbon-schemas.mjs', cache, '--bundle-only', '--inventory')
if (!pro.packages.length || pro.packages.some(p => p.version !== '1.6.0')) throw Error('Expected cached Pro 1.6.0 bundles')
const { placeRibbonGridItems } = await import(pathToFileURL(path.join(root, 'packages/ui/src/views/components/ribbon/ribbon-grid-layout.ts')))
const out = ['# Univer v1.6.0 · 桌面 Ribbon 源码批量清点', '',
  '复算：`node --experimental-strip-types scripts/report-univer-ribbon-batch.mjs`。上游提交 `' + oss.reference.commit + '`。', '',
  '常规表按标签、再按行、组和列排列。格位是将列出的 OSS 注册项全部纳入、未求值 hidden$ 时，用固定版 `ribbon-grid-layout.ts` 得到的布局；不代表某次会话的可见按钮。Classic 列是组内 schema order，不是横向像素坐标。More 只重排已有项，右键、浮动条和子菜单不重复计入。', '']
const groups = { START: ['HISTORY', 'FORMAT', 'LAYOUT', 'NUMBER', 'OTHERS'], INSERT: ['EDIT', 'MEDIA', 'CELL', 'OTHERS'], FORMULAS: ['BASIC', 'OTHERS'], DATA: ['FORMULAS', 'ORGANIZATION', 'RULES', 'ANALYSIS', 'OTHERS'], VIEW: ['DISPLAY', 'VISIBILITY', 'OTHERS'] }
const code = s => '`' + String(s).replaceAll('`', '').replaceAll('|', '\\|') + '`'
for (const p of [oss.sheets, oss.docs]) {
  out.push(`## ${p.product}`, '', `常规根 ${p.rootCount}；上下文根 ${p.contextualCount}。含可选扩展，不等于示例默认挂载数。`, '')
  for (const tab of [...p.tabs.map(t => `RibbonPosition.${t.tab}`), ...new Set(p.contextual.map(e => e.tab))]) {
    const entries = p.entries.filter(e => e.tab === tab)
    const tabName = tab.split('.').at(-1)
    const sequence = groups[tabName] ?? (tab.includes('HEADER_FOOTER') ? ['insertion', 'navigation', 'options', 'position', 'close'] : ['manage'])
    const items = Object.entries(Object.groupBy(entries, e => e.group)).flatMap(([group, records]) => {
      const groupOrder = sequence.indexOf(group.split('.').at(-1))
      return placeRibbonGridItems(records.sort((a,b) => (a.order ?? 0) - (b.order ?? 0)).map(e => ({ key: e.id, e,
        gridLayout: e.gridLayout && Object.fromEntries(Object.entries(e.gridLayout).map(([k,v]) => [k, JSON.parse(v)]))
      }))).map(p => ({ ...p, groupOrder }))
    }).sort((a,b) => a.row - b.row || a.groupOrder - b.groupOrder || a.column - b.column)
    out.push(`### ${tabName} (${items.length})`, '', '| 行 | 组 / 列 (跨行×跨列) | Classic | 根 ID / factory | 来源 |', '|---:|---|---:|---|---|')
    for (const i of items) {
      const e = i.item.e
      out.push(`| ${i.row} | ${e.group.split('.').at(-1)} / ${i.column} (${i.rowSpan}×${i.columnSpan}) | ${e.order ?? 0} | ${code(e.id)} / ${code(e.factory ?? '待追踪')} | ${code(e.source + ':' + e.line)} |`)
    }
    out.push('')
  }
}
out.push('## Pro 发布包', '', `读取 ${pro.packages.length} 个有 ES 入口的 1.6.0 包。下表按包列出 mergeMenu/appendRootMenu 注册调用中识别出的根声明次数，保留不同挂载分支及重复声明；不是唯一功能数，也不是桌面可见数。未解析列仅统计 schema 表达式，factory 的响应式条件需另追踪。`, '',
  '| 包 | 注册调用 | 常规根声明 | 上下文根声明 | 未解析表达式 |', '|---|---:|---:|---:|---:|')
for (const p of pro.packages.filter(p => p.registrations.length)) {
  const entries = p.registrations.flatMap(r => r.entries).filter(e => e.root)
  out.push(`| ${code(p.package)} | ${p.registrations.length} | ${entries.filter(e => e.category === 'ribbon').length} | ${entries.filter(e => e.category === 'contextual-ribbon').length} | ${p.registrations.flatMap(r => r.unresolved).length} |`)
}
out.push('', '商业包的逐声明路径、order、gridLayout、注册语句行列及未解析分支可用 `node scripts/count-univer-pro-ribbon-schemas.mjs .research.local/univer-1.6.0/packages --bundle-only --inventory` 复查。没有执行商业 SDK。', '')
fs.writeFileSync('docs/feature-evidence/RIBBON-REGISTRATION-BATCH-v1.6.0.md', out.join('\n'))
