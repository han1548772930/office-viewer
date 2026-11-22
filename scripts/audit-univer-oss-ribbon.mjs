import fs from 'node:fs'
import path from 'node:path'
import ts from 'typescript'
import { execFileSync } from 'node:child_process'

const root = path.resolve(process.argv.slice(2).find((argument) => !argument.startsWith('--')) ?? 'E:/rust_test/univer')

const products = {
  sheets: [
    'packages/ui/src/menu/schema.ts',
    'packages/sheets-ui/src/menu/schema.ts',
    'packages/sheets-numfmt-ui/src/menu/schema.ts',
    'packages/sheets-formula-ui/src/menu/schema.ts',
    'packages/sheets-drawing-ui/src/menu/schema.ts',
    'packages/sheets-conditional-formatting-ui/src/menu/schema.ts',
    'packages/sheets-filter-ui/src/menu/schema.ts',
    'packages/sheets-hyper-link-ui/src/menu/schema.ts',
    'packages/sheets-data-validation-ui/src/menu/schema.ts',
    'packages/find-replace/src/menu/schema.ts',
    'packages/sheets-note-ui/src/menu/schema.ts',
    'packages/sheets-sort-ui/src/menu/schema.ts',
    'packages/sheets-table-ui/src/menu/schema.ts',
    'packages/sheets-thread-comment-ui/src/menu/schema.ts',
    'packages/sheets-crosshair-highlight/src/menu/schema.ts',
  ],
  docs: [
    'packages/ui/src/menu/schema.ts',
    'packages/docs-ui/src/menu/schema.ts',
    'packages/docs-ui/src/menu/header-footer-ribbon.ts',
    'packages/docs-drawing-ui/src/menu/schema.ts',
    'packages/docs-hyper-link-ui/src/menu/schema.ts',
    'packages/docs-thread-comment-ui/src/menu/schema.ts',
    'packages/docs-find-replace/src/menu/schema.ts',
    'packages/docs-toc-ui/src/menu/schema.ts',
  ],
}

const metadata = new Set(['order', 'title', 'contextual', 'gridLayout', 'menuItemFactory', 'customizeMenuItem', 'headerActionMenuItemFactory', 'replace'])
const ribbonGroupPattern = /^Ribbon(?:Start|Insert|Formulas|Data|View|Others)Group\./
const ribbonTabPattern = /^RibbonPosition\./
const contextualTabIds = {
  docs: new Set(['DOC_TABLE_OF_CONTENTS_RIBBON_TAB', 'DOC_HEADER_FOOTER_RIBBON_TAB']),
  sheets: new Set(),
}

function propertyKey(property) {
  if (!ts.isPropertyAssignment(property)) return undefined
  if (ts.isIdentifier(property.name) || ts.isStringLiteral(property.name) || ts.isNumericLiteral(property.name)) return property.name.text
  return ts.isComputedPropertyName(property.name) ? property.name.expression.getText() : property.name.getText()
}

function literalObject(node) {
  while (node && (ts.isParenthesizedExpression(node) || ts.isAsExpression(node) || ts.isSatisfiesExpression(node))) node = node.expression
  return node && ts.isObjectLiteralExpression(node) ? node : undefined
}

function properties(object) {
  return object?.properties.filter(ts.isPropertyAssignment) ?? []
}

function objectProperty(property) {
  return literalObject(property.initializer)
}

function schemaDeclarations(source) {
  const result = []
  function visit(node) {
    if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name) && /(?:menuSchema|MenuSchema|tableOfContentsRibbonMenuSchema|headerFooterRibbonSchema)$/.test(node.name.text)) {
      const object = literalObject(node.initializer)
      if (object) result.push({ name: node.name.text, object })
    }
    ts.forEachChild(node, visit)
  }
  visit(source)
  return result
}

function metadataFor(object) {
  const result = {}
  for (const property of properties(object)) {
    const key = propertyKey(property)
    if (key === 'order' && ts.isNumericLiteral(property.initializer)) result.order = Number(property.initializer.text)
    if (key === 'gridLayout') {
      const grid = literalObject(property.initializer)
      if (grid) result.gridLayout = Object.fromEntries(properties(grid).map((entry) => [propertyKey(entry), entry.initializer.getText()]))
    }
  }
  return result
}

function childItems(object) {
  return properties(object).filter((property) => !metadata.has(propertyKey(property)))
}

function collectEntries(product, relativePath, source) {
  const entries = []
  const declarations = schemaDeclarations(source)

  function collectGroup(groupKey, groupObject, tabKey, contextual) {
    const tab = tabKey ?? inferTab(groupKey)
    if (!tab) return
    for (const itemProperty of childItems(groupObject)) {
      const itemKey = propertyKey(itemProperty)
      const itemObject = objectProperty(itemProperty)
      const nested = itemObject ? childItems(itemObject) : []
      const children = nested.filter((child) => {
        const key = propertyKey(child)
        return !['type', 'icon', 'tooltip', 'disabled$', 'hidden$', 'label', 'value', 'options', 'shortcut', 'title', 'menuItemFactory', 'customizeMenuItem', 'gridLayout', 'order'].includes(key)
      })
      entries.push({
        product,
        tab,
        group: groupKey,
        id: itemKey,
        children: children.map((child) => propertyKey(child)),
        factory: properties(itemObject).find((entry) => propertyKey(entry) === 'menuItemFactory')?.initializer.getText(),
        ...metadataFor(itemObject),
        contextual,
        source: relativePath,
        line: source.getLineAndCharacterOfPosition(itemProperty.getStart(source)).line + 1,
      })
    }
  }

  function walkObject(object, currentTab, contextual = false) {
    for (const property of properties(object)) {
      const key = propertyKey(property)
      const nested = objectProperty(property)
      if (!nested) continue
      if (ribbonGroupPattern.test(key)) {
        collectGroup(key, nested, currentTab, contextual)
      } else if (ribbonTabPattern.test(key)) {
        walkObject(nested, key, contextual)
      } else if (key === 'MenuManagerPosition.RIBBON') {
        walkObject(nested, currentTab, contextual)
      } else if (contextualTabIds[product].has(key)) {
        walkContextualTab(key, nested)
      }
    }
  }

  function walkContextualTab(tabId, tabObject) {
    for (const groupProperty of properties(tabObject)) {
      const groupObject = objectProperty(groupProperty)
      if (groupObject && !metadata.has(propertyKey(groupProperty))) collectGroup(propertyKey(groupProperty), groupObject, tabId, true)
    }
    for (const spread of tabObject.properties.filter(ts.isSpreadAssignment)) {
      const expression = spread.expression.getText()
      if (!expression.startsWith("Object.fromEntries(['insertion', 'navigation', 'options', 'position'].map")) {
        throw new Error(`Unreviewed contextual generator at ${relativePath}: ${expression}`)
      }
      for (const kind of ['insertion', 'navigation', 'options', 'position']) {
        entries.push({ product, tab: tabId, group: `${tabId}.${kind}`, id: `DOC_HEADER_FOOTER_RIBBON_CONTROL.${kind}`, children: [], order: 0, gridLayout: { row: '1', column: '1', rowSpan: '2' }, contextual: true, factory: 'inline factory with label props.kind', source: relativePath, line: source.getLineAndCharacterOfPosition(spread.getStart(source)).line + 1, generator: expression })
      }
    }
  }

  for (const declaration of declarations) walkObject(declaration.object)

  const used = new Set()
  return entries.filter((entry) => {
    const key = [entry.product, entry.tab, entry.group, entry.id].join('|')
    if (used.has(key)) return false
    used.add(key)
    return true
  })
}

function assembly(product) {
  const mount = `examples/src/${product}/mount.ts`
  const text = fs.readFileSync(path.join(root, mount), 'utf8')
  const source = ts.createSourceFile(mount, text, ts.ScriptTarget.Latest, true)
  const presets = []
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause?.namedBindings || !ts.isNamedImports(statement.importClause.namedBindings)) continue
    for (const element of statement.importClause.namedBindings.elements) {
      if (!element.name.text.endsWith('Preset')) continue
      const name = statement.moduleSpecifier.text.replace('@univerjs/', '')
      const presetPath = `presets/packages/${name}/src/preset.ts`
      const presetText = fs.readFileSync(path.join(root, presetPath), 'utf8')
      const pluginImports = [...presetText.matchAll(/import \{([^}]+)\} from '(@univerjs\/[^']+)';/g)]
        .filter((match) => /Univer\w+Plugin/.test(match[1]))
        .map((match) => ({ package: match[2], plugins: match[1].trim().split(/,\s*/) }))
      presets.push({ name: element.name.text, source: presetPath, pluginImports })
    }
  }
  return { mount, presets }
}

function contextEntries(relativePath) {
  const file = path.join(root, relativePath)
  const source = ts.createSourceFile(file, fs.readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true)
  const entries = []
  const generators = []
  function visit(object, keys) {
    const factory = properties(object).find((property) => propertyKey(property) === 'menuItemFactory')
    if (factory) entries.push({ path: keys, factory: factory.initializer.getText(), source: relativePath, line: source.getLineAndCharacterOfPosition(factory.getStart(source)).line + 1 })
    for (const property of properties(object)) {
      if (metadata.has(propertyKey(property))) continue
      const nested = objectProperty(property)
      if (nested) visit(nested, [...keys, propertyKey(property)])
    }
    for (const spread of object.properties.filter(ts.isSpreadAssignment)) {
      if (/ContextMenuPosition\.|FLOAT_|FloatingObject/.test(keys.join('.'))) generators.push({ path: keys, expression: spread.expression.getText(), source: relativePath, line: source.getLineAndCharacterOfPosition(spread.getStart(source)).line + 1 })
    }
  }
  for (const declaration of schemaDeclarations(source)) visit(declaration.object, [])
  const unique = new Map(entries.map((entry) => [entry.path.join('|'), entry]))
  return { entries: [...unique.values()].filter((entry) => /ContextMenuPosition\.|FLOAT_|FloatingObject/.test(entry.path.join('.'))), generators }
}

function inferTab(groupKey) {
  const value = groupKey.match(/^Ribbon(Start|Insert|Formulas|Data|View|Others)Group\./)?.[1]
  return value ? `RibbonPosition.${value.toUpperCase()}` : undefined
}

function extract(product) {
  const records = []
  for (const relativePath of products[product]) {
    const file = path.join(root, relativePath)
    const text = fs.readFileSync(file, 'utf8')
    const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS)
    records.push(...collectEntries(product, relativePath, source))
  }
  return records
}

function summarize(records, product) {
  const defaultEntries = records.filter((entry) => !entry.contextual)
  const tabs = ['START', 'INSERT', 'FORMULAS', 'DATA', 'VIEW'].map((name) => {
    const tabEntries = defaultEntries.filter((entry) => entry.tab === `RibbonPosition.${name}`)
    const groups = [...new Set(tabEntries.map((entry) => entry.group))]
    return { tab: name, count: tabEntries.length, groups: groups.map((group) => ({ group, count: tabEntries.filter((entry) => entry.group === group).length })) }
  }).filter((entry) => entry.count)
  const contextual = records.filter((entry) => entry.contextual)
  const duplicates = defaultEntries.map((entry) => `${entry.tab}|${entry.group}|${entry.id}`).filter((key, index, all) => all.indexOf(key) !== index)
  const contexts = products[product].map(contextEntries)
  return { product, assembly: assembly(product), rootCount: defaultEntries.length, tabs, contextualCount: contextual.length, contextual, duplicates, entries: records, contextEntries: contexts.flatMap((entry) => entry.entries), contextGenerators: contexts.flatMap((entry) => entry.generators) }
}

const results = Object.fromEntries(Object.keys(products).map((product) => [product, summarize(extract(product), product)]))
results.reference = { root, commit: execFileSync('git', ['-C', root, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(), tag: execFileSync('git', ['-C', root, 'describe', '--tags', '--exact-match'], { encoding: 'utf8' }).trim(), method: 'Static AST registration inventory. Product schemas are projected by ownership. Source expressions are preserved. Factories and their dynamic option lists are not executed.' }
const output = `${JSON.stringify(results, null, 2)}\n`
const outputPath = process.argv.find((argument) => argument.startsWith('--output='))?.slice('--output='.length)
if (outputPath) fs.writeFileSync(outputPath, output, 'utf8')
else process.stdout.write(output)
