import { createWasmTaskControl } from '../../packages/office-studio/src/platform/wasm/task-protocol.ts'
import type { WasmTaskControl, WasmTaskKind } from '../../packages/office-studio/src/platform/wasm/task-protocol.ts'

/** Serializable controls for direct production-artifact tests. */
export function wasmTaskFixture(kind: WasmTaskKind, requestId: string, unitId = 'office-wasm-fixture'): WasmTaskControl {
  return createWasmTaskControl(kind, requestId, unitId, 0, {
    maxInputBytes: 32 * 1024 * 1024,
    maxOutputBytes: 32 * 1024 * 1024,
    maxWorkingBytes: 96 * 1024 * 1024,
  })
}
