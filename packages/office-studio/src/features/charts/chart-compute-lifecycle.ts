import { ComputeError } from '../../compute/errors.ts'
import { COMPUTE_LIMITS } from '../../compute/protocol.ts'
import type { ChartStatisticsOutput, ChartStatisticsTask } from '../../compute/chart-protocol'
import type { ChartImageResource } from '../../domain/model'

export interface ChartComputePreparationContext {
  signal?: AbortSignal
  /** Recheck the captured unit, source/config revision and permission. */
  assertCurrent?: () => void
  memoryBudgetBytes?: number
  /** The owning service supplies document/revision/cancellation and transferable
   * scheduling; preparation retains only the string dictionaries on this side. */
  runChartTask?: (task: ChartStatisticsTask) => Promise<ChartStatisticsOutput>
  /** Synchronous host registry lookup; returned data remains owned by the model. */
  resolveImageResource?: (id: string) => ChartImageResource | undefined
}

export const CHART_PREPARATION_CHUNK = 8192

export function chartPreparationBudget(context: ChartComputePreparationContext): number {
  const budget = context.memoryBudgetBytes ?? COMPUTE_LIMITS.defaultMemoryBudget
  if (!Number.isSafeInteger(budget) || budget <= 0 || budget > 0xffff_ffff) {
    throw new ComputeError('COMPUTE_INVALID_INPUT', '图表数据准备的内存预算无效')
  }
  return budget
}

export function assertChartPreparationCurrent(context: ChartComputePreparationContext): void {
  if (context.signal?.aborted) throw new ComputeError('COMPUTE_CANCELLED', '图表数据准备已取消')
  context.assertCurrent?.()
}

export function assertChartPreparationMemory(bytes: number, context: ChartComputePreparationContext): void {
  if (!Number.isSafeInteger(bytes) || bytes < 0 || bytes > chartPreparationBudget(context)) {
    throw new ComputeError('COMPUTE_MEMORY_LIMIT', '图表源数据和计算结果超出内存预算')
  }
}

/** A macro task allows ordinary Worker/DOM cancellation and revision messages
 * to run. Every caller checks again after yielding before consuming more data. */
export async function yieldChartPreparation(context: ChartComputePreparationContext): Promise<void> {
  assertChartPreparationCurrent(context)
  await new Promise<void>(resolve => setTimeout(resolve, 0))
  assertChartPreparationCurrent(context)
}

export async function consumeChartPreparation<T>(steps: Generator<void, T, void>, context: ChartComputePreparationContext = {}): Promise<T> {
  const owner: Iterator<void, unknown, void> = steps
  try {
    chartPreparationBudget(context)
    assertChartPreparationCurrent(context)
    for (;;) {
      const next = steps.next()
      if (next.done) { assertChartPreparationCurrent(context); return next.value }
      await yieldChartPreparation(context)
    }
  } finally {
    // Release a suspended generator's retained source on cancellation/error.
    owner.return?.()
  }
}

/** The synchronous small-source Facade consumes the same projection algorithm. */
export function consumeChartPreparationSync<T>(steps: Generator<void, T, void>, context: ChartComputePreparationContext = {}): T {
  const owner: Iterator<void, unknown, void> = steps
  try {
    chartPreparationBudget(context)
    for (;;) {
      assertChartPreparationCurrent(context)
      const next = steps.next()
      if (next.done) { assertChartPreparationCurrent(context); return next.value }
    }
  } finally { owner.return?.() }
}
