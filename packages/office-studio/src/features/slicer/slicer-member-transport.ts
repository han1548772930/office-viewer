import { OfficeComputeClient, type ComputeClientOptions, type ComputeJob } from '../../compute/client.ts'
import { ComputeError } from '../../compute/errors.ts'
import type { ComputeMetrics } from '../../compute/protocol.ts'
import type { SlicerMembersOutput, SlicerMembersTask } from '../../compute/slicer-member-protocol.ts'
import { readSlicerMembersOutput } from '../../compute/slicer-member-results.ts'
import { isWasmTaskControl } from '../../platform/wasm/task-protocol.ts'

export type SlicerMemberComputeClient = Pick<OfficeComputeClient, 'submit' | 'setDocumentRevision' | 'forgetDocument'>
export interface SlicerMemberTransportContext { unitId: string; baseRevision: number; signal: AbortSignal; timeoutMs: number }
export interface SlicerMemberTransportResult { output: SlicerMembersOutput; metrics: ComputeMetrics }

/** Reuses the protocol-2 bounded compute queue and its real Worker/WASM path.
 * Injection supports Studio's shared client and deterministic transport tests;
 * there is no production TypeScript calculation fallback when WASM is missing.
 */
export class SlicerMemberQueryTransport {
  private readonly client: SlicerMemberComputeClient
  private readonly ownedClient?: OfficeComputeClient
  private readonly running = new Map<ComputeJob, string>()
  private disposed = false

  constructor(client?: SlicerMemberComputeClient, options: ComputeClientOptions = {}) {
    if (client) this.client = client
    else { this.ownedClient = new OfficeComputeClient(options); this.client = this.ownedClient }
  }

  async query(task: SlicerMembersTask, context: SlicerMemberTransportContext): Promise<SlicerMemberTransportResult> {
    if (this.disposed) throw new ComputeError('COMPUTE_DISPOSED', '切片器查询 transport 已释放')
    const job = this.client.submit(task, { documentId: context.unitId, modelRevision: context.baseRevision,
      ownership: 'transfer', priority: 'interactive', signal: context.signal, timeoutMs: context.timeoutMs })
    this.running.set(job, context.unitId)
    try {
      const result = await job.promise
      const control = result.context
      if (this.disposed || context.signal.aborted) throw new ComputeError('COMPUTE_CANCELLED', '切片器查询已取消')
      if (!isWasmTaskControl(control) || control.requestId !== job.requestId || result.requestId !== job.requestId
        || control.kind !== 'slicerMembers' || control.unitId !== context.unitId || control.baseRevision !== context.baseRevision
        || result.documentId !== context.unitId || result.modelRevision !== context.baseRevision) {
        throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '切片器查询返回了不匹配的 unit 或 source revision')
      }
      const output = readSlicerMembersOutput(result.output, task.params)
      if (!output) throw new ComputeError('COMPUTE_PROTOCOL_MISMATCH', '切片器查询返回了无效的窗口或日期边界')
      return { output, metrics: result.metrics }
    } finally { this.running.delete(job) }
  }

  forgetDocument(unitId: string): void {
    for (const [job, owner] of this.running) if (owner === unitId) job.cancel()
    this.client.forgetDocument(unitId)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    for (const job of this.running.keys()) job.cancel()
    this.running.clear()
    this.ownedClient?.dispose()
  }
}
