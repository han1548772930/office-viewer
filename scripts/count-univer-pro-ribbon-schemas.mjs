import fs from 'node:fs'
import path from 'node:path'
import ts from 'typescript'

const packagesRoot = path.resolve('.research.local/univer/packages')
const packages = fs.readdirSync(packagesRoot, { withFileTypes: true })
  .filter((entry) => entry.isDirectory())
  .filter((entry) => /^(docs|sheets)-.*-ui$/.test(entry.name))
  .filter((entry) => fs.existsSync(path.join(packagesRoot, entry.name, 'lib/types/menu/schema.d.ts')))

function importedNames(source) {
  const names = new Map()
  for (const statement of source.statements) {
    if (!ts.isImportDeclaration(statement) || !statement.importClause?.namedBindings ||
        !ts.isNamedImports(statement.importClause.namedBindings)) continue
    for (const element of statement.importClause.namedBindings.elements) {
      const imported = (element.propertyName ?? element.name).text
      const local = element.name.text
      if (/^Ribbon(?:Start|Insert|View|Formula|Data|Others)Group$/.test(imported) || imported === 'RibbonPosition') {
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
