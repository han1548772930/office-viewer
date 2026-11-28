use crate::{ComputeError, ValidationListQuery, ValidationListQueryParams, ValidationListQueryResult};
use js_sys::{Uint8Array, Uint32Array};
use office_wasm_bridge::{BoundaryCopyStats, BridgeError, TaskControl, TaskKind, copy_bytes_in, copy_indices_in,
    copy_indices_out, error_to_js, from_control, reject_linear_memory_input, to_control};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

fn js_error(error: ComputeError) -> JsValue { error_to_js(BridgeError::new(error.code, error.message)) }
fn consumed() -> JsValue { js_error(ComputeError::new("COMPUTE_RESULT_CONSUMED", "List query was consumed or failed")) }

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ListQueryControl { context: TaskControl, params: ValidationListQueryParams }

#[wasm_bindgen]
pub struct TypedValidationListQuery {
    inner: Option<ValidationListQuery>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
pub fn create_validation_list_query(offsets: Uint32Array, text_bytes: Uint8Array, query_bytes: Uint8Array,
    control: JsValue) -> Result<TypedValidationListQuery, JsValue> {
    for buffer in [offsets.buffer(), text_bytes.buffer(), query_bytes.buffer()] { reject_linear_memory_input(&buffer)?; }
    let control: ListQueryControl = from_control(control)?;
    control.context.validate(TaskKind::ValidationListQuery).map_err(error_to_js)?;
    control.params.validate_shape(offsets.length() as usize, text_bytes.length() as usize, query_bytes.length() as usize).map_err(js_error)?;
    let memory = control.params.memory_estimate().map_err(js_error)?;
    control.context.budget.check_input(memory.input_bytes).map_err(error_to_js)?;
    control.context.budget.check_output(memory.output_bytes).map_err(error_to_js)?;
    control.context.budget.check_working(memory.estimated_peak_bytes).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let maximum = control.context.budget.max_input_bytes;
    let offsets = copy_indices_in(&offsets, maximum, &mut transfer)?;
    let text = copy_bytes_in(&text_bytes, maximum, &mut transfer)?;
    let query = copy_bytes_in(&query_bytes, maximum, &mut transfer)?;
    let inner = ValidationListQuery::new(offsets, text, query, control.params, control.context.budget.max_working_bytes as usize).map_err(js_error)?;
    Ok(TypedValidationListQuery { inner: Some(inner), context: control.context, transfer })
}

#[wasm_bindgen]
impl TypedValidationListQuery {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        match self.inner.as_mut().ok_or_else(consumed)?.step(budget) {
            Ok(done) => Ok(done),
            Err(error) => { self.inner = None; Err(js_error(error)) },
        }
    }
    pub fn progress(&self) -> Result<f64, JsValue> { Ok(self.inner.as_ref().ok_or_else(consumed)?.progress()) }
    pub fn finish(&mut self) -> Result<TypedValidationListQueryResult, JsValue> {
        let result = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        self.inner = None;
        Ok(TypedValidationListQueryResult { result, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}

#[wasm_bindgen]
pub struct TypedValidationListQueryResult {
    result: ValidationListQueryResult,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
impl TypedValidationListQueryResult {
    /// The shared bridge creates a synchronous lease after destination
    /// allocation and copies only the requested indices. No view escapes,
    /// crosses await or survives memory.grow/reentrancy; the JS buffer is independent.
    pub fn copy_indices(&mut self) -> Result<Uint32Array, JsValue> {
        self.context.budget.check_output(self.result.indices.len() * 4).map_err(error_to_js)?;
        copy_indices_out(&self.result.indices, &mut self.transfer)
    }
    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            context: &'a TaskControl,
            index_length: usize,
            byte_length: usize,
            total: u32,
            transfer: &'a BoundaryCopyStats,
        }
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata { context: &self.context, index_length: self.result.indices.len(),
            byte_length: self.result.indices.len() * 4, total: self.result.total, transfer: &self.transfer })
    }
}
