import fs from 'node:fs'
import path from 'node:path'
import ts from 'typescript'

const packagesRoot = path.resolve(process.argv.slice(2).find((argument) => !argument.startsWith('--')) ?? '.research.local/univer/packages')
const bundleOnly = process.argv.includes('--bundle-only')
const inventory = process.argv.includes('--inventory')
const packages = fs.readdirSync(packagesRoot, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .filter((entry) => inventory || /^(docs|sheets)-.*-ui$/.test(entry.name))
  .filter((entry) => bundleOnly || fs.existsSync(path.join(packagesRoot, entry.name, 'lib/types/menu/schema.d.ts')))

function importedNames(source) {
  const names = new Map()
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause?.namedBindings ||
        !ts.isNamedImports(statement.importClause.namedBindings)) continue
    for (const element of statement.importClause.namedBindings.elements) {
      const imported = (element.propertyName ?? element.name).text
      const local = element.name.text
      if (/^Ribbon(?:Start|Insert|View|Formulas|Data|Others)Group$/.test(imported) || imported === 'RibbonPosition') {
        names.set(local, imported)
      }
    }
  }
  return names
}

function getKeyName(name, imports) {
  if (ts.isIdentifier(name) || ts.isStringLiteral(name) || ts.isNumericLiteral(name)) return name.text
  if (!ts.isComputedPropertyName(name)) return undefined
  const expression = name.expression
  if (!ts.isPropertyAccessExpression(expression)) return undefined
  const owner = ts.isIdentifier(expression.expression) ? imports.get(expression.expression.text) : undefined
  return owner ? `${owner}.${expression.name.text}` : undefined
}

function objectGroups(object, imports) {
  const groups = []
  for (const property of object.properties) {
    if (!ts.isPropertyAssignment(property) || !ts.isObjectLiteralExpression(property.initializer)) continue
    const name = getKeyName(property.name, imports)
    if (name?.startsWith('Ribbon') && name.endsWith('Group.' + name.split('.').at(-1))) {
      groups.push({ name, count: property.initializer.properties.length })
    }
  }
  return groups
}

const result = []
// Inventory mode walks registrations and schema references without importing a bundle.
// Unknown expressions remain explicit so partial static evidence cannot become a runtime claim.
const moduleCache = new Map()

function readModule(directory) {
  if (moduleCache.has(directory)) return moduleCache.get(directory)
  const indexPath = path.join(directory, 'lib/es/index.js')
  if (!fs.existsSync(indexPath)) return undefined
  const source = ts.createSourceFile(indexPath, fs.readFileSync(indexPath, 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS)
  const context = { source, directory, bindings: new Map(), imports: new Map(), exports: new Map() }
  moduleCache.set(directory, context)
  for (const statement of source.statements) {
    if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        if (ts.isIdentifier(declaration.name) && declaration.initializer) context.bindings.set(declaration.name.text, declaration.initializer)
      }
    } else if (ts.isFunctionDeclaration(statement) && statement.name) {
      context.bindings.set(statement.name.text, statement)
    } else if (ts.isImportDeclaration(statement) && statement.importClause?.namedBindings && ts.isNamedImports(statement.importClause.namedBindings)) {
      for (const element of statement.importClause.namedBindings.elements) {
        context.imports.set(element.name.text, { name: (element.propertyName ?? element.name).text, module: statement.moduleSpecifier.text })
      }
    } else if (ts.isExportDeclaration(statement) && statement.exportClause && ts.isNamedExports(statement.exportClause)) {
      for (const element of statement.exportClause.elements) context.exports.set(element.name.text, (element.propertyName ?? element.name).text)
    }
  }
  return context
}

function reference(node, context) {
  if (ts.isIdentifier(node)) return context.imports.get(node.text)?.name ?? node.text
  if (ts.isPropertyAccessExpression(node)) return `${reference(node.expression, context)}.${node.name.text}`
  return node.getText(context.source)
}

function unknown(node, context) {
  return { expression: node.getText(context.source), reference: reference(node, context) }
}

function isUnknown(value) {
  return value && typeof value === 'object' && 'expression' in value && 'reference' in value
}

function lookup(node, context, seen = new Set()) {
  if (!ts.isIdentifier(node)) return { node, context }
  const identity = `${context.directory}:${node.text}`
  if (seen.has(identity)) return { node, context }
  seen.add(identity)
  const local = context.bindings.get(node.text)
  if (local) return lookup(local, context, seen)
  const imported = context.imports.get(node.text)
  if (imported?.module.startsWith('@univerjs-pro/')) {
    const dependency = readModule(path.join(packagesRoot, imported.module.split('/')[1]))
    const binding = dependency?.bindings.get(dependency.exports.get(imported.name) ?? imported.name)
    if (binding) return lookup(binding, dependency, seen)
  }
  return { node, context }
}

function keyOf(name, context, depth) {
  if (!ts.isComputedPropertyName(name)) return name.text ?? name.getText(context.source)
  const value = valueOf(name.expression, context, depth + 1)
  return isUnknown(value) ? value.reference : String(value)
}

function ownReturns(node) {
  if (ts.isArrowFunction(node) && !ts.isBlock(node.body)) return [node.body]
  const returns = []
  function visit(child) {
    if (ts.isReturnStatement(child) && child.expression) returns.push(child.expression)
    else if (!ts.isFunctionLike(child)) ts.forEachChild(child, visit)
  }
  if (node.body) ts.forEachChild(node.body, visit)
  return returns
}

function valueOf(node, context, depth = 0) {
  if (!node) return undefined
  if (depth > 24) return unknown(node, context)
  if (ts.isStringLiteralLike(node) || ts.isNumericLiteral(node)) return ts.isNumericLiteral(node) ? Number(node.text) : node.text
  if (node.kind === ts.SyntaxKind.TrueKeyword) return true
  if (node.kind === ts.SyntaxKind.FalseKeyword) return false
  if (node.kind === ts.SyntaxKind.NullKeyword) return null
  if (ts.isParenthesizedExpression(node)) return valueOf(node.expression, context, depth + 1)
  if (ts.isPrefixUnaryExpression(node) && node.operator === ts.SyntaxKind.MinusToken) {
    const operand = valueOf(node.operand, context, depth + 1)
    return typeof operand === 'number' ? -operand : unknown(node, context)
  }
  if (ts.isIdentifier(node)) {
    const resolved = lookup(node, context)
    if (resolved.node !== node) return valueOf(resolved.node, resolved.context, depth + 1)
  }
  if (ts.isPropertyAccessExpression(node)) {
    const object = valueOf(node.expression, context, depth + 1)
    if (object && !isUnknown(object) && Object.hasOwn(object, node.name.text)) return object[node.name.text]
  }
  if (ts.isTemplateExpression(node)) {
    let text = node.head.text
    for (const span of node.templateSpans) {
      const value = valueOf(span.expression, context, depth + 1)
      if (typeof value !== 'string' && typeof value !== 'number') return unknown(node, context)
      text += String(value) + span.literal.text
    }
    return text
  }
  if (ts.isObjectLiteralExpression(node)) {
    const object = {}
    for (const property of node.properties) {
      if (ts.isPropertyAssignment(property)) object[keyOf(property.name, context, depth)] = valueOf(property.initializer, context, depth + 1)
      else if (ts.isSpreadAssignment(property)) {
        const spread = valueOf(property.expression, context, depth + 1)
        if (spread && !isUnknown(spread)) Object.assign(object, spread)
        else object[`<spread:${property.pos}>`] = spread
      }
    }
    return object
  }
  if (ts.isCallExpression(node) && node.arguments.length === 0) {
    const callable = lookup(node.expression, context)
    if (ts.isFunctionLike(callable.node)) {
      const returns = ownReturns(callable.node)
      if (returns.length === 1) return valueOf(returns[0], callable.context, depth + 1)
    }
  }
  return unknown(node, context)
}

function location(node, context) {
  const offset = node.getStart(context.source)
  const position = context.source.getLineAndCharacterOfPosition(offset)
  return { file: path.relative(packagesRoot, context.source.fileName).replaceAll('\\', '/'), offset, line: position.line + 1, column: position.character + 1 }
}

function factoryEvidence(expression, context, seen = new Set()) {
  if (!isUnknown(expression)) return { unresolved: expression }
  const callable = context.bindings.get(expression.expression)
  if (!callable || !ts.isFunctionLike(callable) || seen.has(callable)) return { unresolved: expression }
  seen.add(callable)
  const variants = []
  for (const returned of ownReturns(callable)) {
    if (!ts.isObjectLiteralExpression(returned)) continue
    const fields = {}
    for (const property of returned.properties) {
      if (ts.isSpreadAssignment(property) && ts.isCallExpression(property.expression)) {
        const inherited = factoryEvidence(unknown(property.expression.expression, context), context, new Set(seen))
        if (inherited.variants?.length === 1) Object.assign(fields, inherited.variants[0])
        else fields.inheritedFactory = inherited
      } else if (ts.isPropertyAssignment(property)) {
        const key = keyOf(property.name, context, 0)
        fields[key] = valueOf(property.initializer, context)
      }
    }
    variants.push(fields)
  }
  return { reference: expression.reference, source: location(callable, context), variants }
}

const ribbonGroup = /^Ribbon(?:Start|Insert|View|Formulas|Data|Others)Group\./
const menuFields = new Set(['order', 'gridLayout', 'menuItemFactory', 'title', 'contextual', 'replace', 'quickLayout', 'hidden$', 'disabled$'])

function inspectRegistration(node, context) {
  const schema = valueOf(node.arguments[0], context)
  const entries = []
  const unresolved = []
  function walk(object, keys, parents, category) {
    if (!object || typeof object !== 'object' || isUnknown(object)) {
      if (isUnknown(object)) unresolved.push({ path: keys, ...object })
      return
    }
    const contextual = parents.some((parent) => parent.contextual === true) || object.contextual === true
    const currentCategory = contextual ? 'contextual-ribbon' : category
    if (object.menuItemFactory) {
      entries.push({ path: keys, category: currentCategory, root: keys.length === 1 || ribbonGroup.test(keys.at(-2) ?? '') ||
        (keys[0] === 'MenuManagerPosition.RIBBON' && keys.length === 4),
      order: object.order, gridLayout: object.gridLayout, parents: parents.map((parent) => ({ order: parent.order, title: parent.title, contextual: parent.contextual })),
      factory: factoryEvidence(object.menuItemFactory, context) })
    }
    for (const [key, value] of Object.entries(object)) {
      if (menuFields.has(key)) continue
      let nextCategory = currentCategory
      if (ribbonGroup.test(key) || key === 'MenuManagerPosition.RIBBON') nextCategory = contextual ? 'contextual-ribbon' : 'ribbon'
      if (/^ContextMenuPosition\./.test(key)) nextCategory = 'context-menu'
      walk(value, [...keys, key], [...parents, object], nextCategory)
    }
  }
  walk(schema, [], [], 'other-menu')
  return { method: node.expression.name.text, argument: node.arguments[0]?.getText(context.source), source: location(node, context), entries, unresolved }
}

if (inventory) {
  const report = []
  for (const entry of packages) {
    const context = readModule(path.join(packagesRoot, entry.name))
    if (!context) continue
    const registrations = []
    function visit(node) {
      if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) &&
          ['mergeMenu', 'appendRootMenu'].includes(node.expression.name.text)) registrations.push(inspectRegistration(node, context))
      ts.forEachChild(node, visit)
    }
    visit(context.source)
    const manifest = JSON.parse(fs.readFileSync(path.join(context.directory, 'package.json'), 'utf8'))
    report.push({ package: manifest.name, version: manifest.version, registrations })
  }
  console.log(JSON.stringify({ staticOnly: true, packages: report }, null, 2))
  process.exit(0)
}

for (const entry of packages) {
  const indexPath = path.join(packagesRoot, entry.name, 'lib/es/index.js')
  if (!fs.existsSync(indexPath)) continue
  const text = fs.readFileSync(indexPath, 'utf8')
  const source = ts.createSourceFile(indexPath, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS)
  const imports = importedNames(source)
  const objects = []
  function visit(node) {
    if (ts.isVariableDeclaration(node) && node.initializer && ts.isObjectLiteralExpression(node.initializer)) {
      const groups = objectGroups(node.initializer, imports)
      if (groups.length) objects.push({ declaration: node.name.getText(source), groups })
    }
    ts.forEachChild(node, visit)
  }
  visit(source)
  if (objects.length) result.push({ package: entry.name, objects })
}

for (const item of result) {
  console.log(`${item.package}\t${JSON.stringify(item.objects)}`)
}
