// Independent, fixed expected scenes. Tuple expansion copies identities and
// coordinates; it performs no accumulation, stacking or role inference.
const fixtures = {
  'style-changes': {
    stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 3, seriesFields: [1], cumulative: 8,
    tuples: [[0, 0, 0, 0, 0, 6, 6], [0, 1, 1, 0, 6, 4, -2], [0, 2, 2, 0, 4, 8, 4], [0, 3, 3, 3, 0, 8, 8]],
    links: [[0, 1, 6, 6], [1, 2, 4, 4]],
  },
  'style-changes-no-subtotal': {
    stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 3, seriesFields: [1], cumulative: 8,
    tuples: [[0, 0, 0, 0, 0, 6, 6], [0, 1, 1, 0, 6, 4, -2], [0, 2, 2, 0, 4, 8, 4]],
    links: [[0, 1, 6, 6], [1, 2, 4, 4]],
  },
  'style-total': {
    stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 3, seriesFields: [1], cumulative: 4,
    tuples: [[0, 0, 0, 0, 0, 6, 6], [0, 1, 1, 0, 6, 4, -2], [0, 2, 2, 1, 0, 4, 4], [0, 3, 3, 3, 0, 4, 4]],
    links: [[0, 1, 6, 6], [1, 2, 4, 4]],
  },
  'style-total-no-subtotal': {
    stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 3, seriesFields: [1], cumulative: 4,
    tuples: [[0, 0, 0, 0, 0, 6, 6], [0, 1, 1, 0, 6, 4, -2], [0, 2, 2, 1, 0, 4, 4]],
    links: [[0, 1, 6, 6], [1, 2, 4, 4]],
  },
  'total-bridge': {
    stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 3, seriesFields: [1], cumulative: 10,
    tuples: [[0, 0, 0, 0, 0, 6, 6], [0, 1, 1, 0, 6, 4, -2], [0, 2, 2, 1, 0, 10, 10]],
    links: [[0, 1, 6, 6], [1, 2, 4, 10]],
  },
  arrangement: {
    stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 2, seriesFields: [1, 2], cumulative: 9,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [0, 1, 1, 0, 10, 7, -3], [0, 2, 2, 3, 0, 9, 9],
      [1, 0, 3, 0, 7, 11, 4], [1, 1, 4, 0, 11, 9, -2], [1, 2, 5, 3, 0, 9, 9]],
    links: [[0, 1, 10, 10], [1, 3, 7, 7], [3, 4, 11, 11]],
  },
  'arrangement-no-subtotal': {
    stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 2, seriesFields: [1, 2], cumulative: 9,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [0, 1, 1, 0, 10, 7, -3], [1, 0, 2, 0, 7, 11, 4], [1, 1, 3, 0, 11, 9, -2]],
    links: [[0, 1, 10, 10], [1, 2, 7, 7], [2, 3, 11, 11]],
  },
  'arrangement-total': {
    stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 2, seriesFields: [1, 2], cumulative: 2,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [0, 1, 1, 0, 10, 7, -3], [0, 2, 2, 3, 0, 2, 2],
      [1, 0, 3, 1, 0, 4, 4], [1, 1, 4, 0, 4, 2, -2], [1, 2, 5, 3, 0, 2, 2]],
    links: [[0, 1, 10, 10], [1, 3, 7, 4], [3, 4, 4, 4]],
  },
  stacked: {
    stackType: 'stacked', useSubtotal: true, sourceCategoryCount: 2, seriesFields: [1, 2], cumulative: 9,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [1, 0, 0, 0, 10, 14, 4], [0, 1, 1, 0, 12, 9, -3], [1, 1, 1, 0, 14, 12, -2],
      [0xffff_ffff, 2, 2, 3, 0, 9, 9]],
    links: [[0, 1, 14, 14]],
  },
  'stacked-no-subtotal': {
    stackType: 'stacked', useSubtotal: false, sourceCategoryCount: 2, seriesFields: [1, 2], cumulative: 9,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [1, 0, 0, 0, 10, 14, 4], [0, 1, 1, 0, 12, 9, -3], [1, 1, 1, 0, 14, 12, -2]],
    links: [[0, 1, 14, 14]],
  },
  'named-changes': {
    stackType: 'arrangement', useSubtotal: false, sourceCategoryCount: 3, seriesFields: [1], cumulative: 5,
    tuples: [[0, 0, 0, 0, 0, 2, 2], [0, 1, 1, 0, 2, 1, -1], [0, 2, 2, 0, 1, 5, 4]],
    links: [[0, 1, 2, 2], [1, 2, 1, 1]],
  },
  'catalog-changes': {
    stackType: 'arrangement', useSubtotal: true, sourceCategoryCount: 3, seriesFields: [1], cumulative: 60,
    tuples: [[0, 0, 0, 0, 0, 10, 10], [0, 1, 1, 0, 10, 30, 20], [0, 2, 2, 0, 30, 60, 30], [0, 3, 3, 3, 0, 60, 60]],
    links: [[0, 1, 10, 10], [1, 2, 30, 30]],
  },
}

export function waterfallLayoutFixture(name) {
  const fixture = fixtures[name]
  if (!fixture) throw new RangeError(`Unknown fixed waterfall fixture: ${name}`)
  const { tuples, links, ...metadata } = structuredClone(fixture)
  return { ...metadata,
    points: tuples.map(([seriesPosition, dataPointIndex, categoryIndex, flags, start, end, value]) => ({
      seriesPosition, dataPointIndex, categoryIndex, synthetic: flags === 3, role: flags === 0 ? 'change' : 'total', start, end, value,
    })),
    connectors: links.map(([fromCategory, toCategory, fromValue, toValue]) => ({
      from: { categoryIndex: fromCategory, value: fromValue }, to: { categoryIndex: toCategory, value: toValue },
    })),
  }
}

export function waterfallComputedFixture(source, name) {
  return { preparedData: source, waterfallLayout: waterfallLayoutFixture(name) }
}
