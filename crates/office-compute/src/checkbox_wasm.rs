use crate::{CheckboxCells, CheckboxCellsParams, CheckboxCellsResult, ComputeError};
use js_sys::{Uint8Array, Uint32Array};
use office_wasm_bridge::{BoundaryCopyStats, BridgeError, TaskControl, TaskKind, copy_bytes_in, copy_indices_in,
    copy_indices_out, error_to_js, from_control, reject_linear_memory_input, to_control};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
fn js_error(error: ComputeError) -> JsValue { error_to_js(BridgeError::new(error.code, error.message)) }
fn consumed() -> JsValue { js_error(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Checkbox task was consumed or failed")) }
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct CheckboxControl { context: TaskControl, params: CheckboxCellsParams }

#[wasm_bindgen]
pub struct TypedCheckboxCells { inner: Option<CheckboxCells>, context: TaskControl, transfer: BoundaryCopyStats }
#[wasm_bindgen]
pub fn create_checkbox_cells(ranges: Uint32Array, merges: Uint32Array, positions: Uint32Array, actions: Uint8Array,
    control: JsValue) -> Result<TypedCheckboxCells, JsValue> {
    for buffer in [ranges.buffer(), merges.buffer(), positions.buffer(), actions.buffer()] { reject_linear_memory_input(&buffer)?; }
    let control: CheckboxControl = from_control(control)?;
    control.context.validate(TaskKind::CheckboxCells).map_err(error_to_js)?;
    control.params.validate_shape(ranges.length() as usize, merges.length() as usize, positions.length() as usize, actions.length() as usize).map_err(js_error)?;
    let memory = control.params.memory_estimate().map_err(js_error)?;
    control.context.budget.check_input(memory.input_bytes).map_err(error_to_js)?;
    control.context.budget.check_output(memory.output_bytes).map_err(error_to_js)?;
    control.context.budget.check_working(memory.estimated_peak_bytes).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let maximum = control.context.budget.max_input_bytes;
    let ranges = copy_indices_in(&ranges, maximum, &mut transfer)?;
    let merges = copy_indices_in(&merges, maximum, &mut transfer)?;
    let positions = copy_indices_in(&positions, maximum, &mut transfer)?;
    let actions = copy_bytes_in(&actions, maximum, &mut transfer)?;
    let inner = CheckboxCells::new(ranges, merges, positions, actions, control.params, control.context.budget.max_working_bytes as usize).map_err(js_error)?;
    Ok(TypedCheckboxCells { inner: Some(inner), context: control.context, transfer })
}
#[wasm_bindgen]
impl TypedCheckboxCells {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        match self.inner.as_mut().ok_or_else(consumed)?.step(budget) {
            Ok(done) => Ok(done), Err(error) => { self.inner = None; Err(js_error(error)) },
        }
    }
    pub fn progress(&self) -> Result<f64, JsValue> { Ok(self.inner.as_ref().ok_or_else(consumed)?.progress()) }
    pub fn finish(&mut self) -> Result<TypedCheckboxCellsResult, JsValue> {
        let result = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        self.inner = None;
        Ok(TypedCheckboxCellsResult { result, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}
#[wasm_bindgen]
pub struct TypedCheckboxCellsResult { result: CheckboxCellsResult, context: TaskControl, transfer: BoundaryCopyStats }
#[wasm_bindgen]
impl TypedCheckboxCellsResult {
    pub fn copy_runs(&mut self) -> Result<Uint32Array, JsValue> {
        self.context.budget.check_output(self.result.runs.len() * 4).map_err(error_to_js)?;
        // Shared bridge allocates the destination before acquiring a synchronous lease,
        // rebuilds the view on current memory and releases it before any await/reentry.
        copy_indices_out(&self.result.runs, &mut self.transfer)
    }
    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> { context: &'a TaskControl, byte_length: usize, run_length: usize, changed_cells: u32, transfer: &'a BoundaryCopyStats }
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata { context: &self.context, byte_length: self.result.runs.len() * 4, run_length: self.result.runs.len(),
            changed_cells: self.result.changed_cells, transfer: &self.transfer })
    }
}
