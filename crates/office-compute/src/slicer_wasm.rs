use crate::{ComputeError, SlicerMembers, SlicerMembersParams, SlicerMembersResult};
use js_sys::{Float64Array, Uint8Array, Uint32Array};
use office_wasm_bridge::{BoundaryCopyStats, BridgeError, TaskControl, TaskKind, copy_bytes_in,
    copy_bytes_out, copy_indices_in, copy_indices_out, copy_numbers_in, error_to_js, from_control,
    reject_linear_memory_input, to_control};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

fn js_error(error: ComputeError) -> JsValue { error_to_js(BridgeError::new(error.code, error.message)) }
fn consumed() -> JsValue {
    js_error(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Slicer query was consumed or failed"))
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SlicerControl { context: TaskControl, params: SlicerMembersParams }

#[wasm_bindgen]
pub struct TypedSlicerMembers {
    inner: Option<SlicerMembers>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
#[allow(clippy::too_many_arguments)]
pub fn create_slicer_members(values: Float64Array, states: Uint8Array, text_ids: Uint32Array,
    availability: Uint8Array, text_offsets: Uint32Array, text_bytes: Uint8Array,
    text_ranks: Uint32Array, control: JsValue) -> Result<TypedSlicerMembers, JsValue> {
    for buffer in [values.buffer(), states.buffer(), text_ids.buffer(), availability.buffer(),
        text_offsets.buffer(), text_bytes.buffer(), text_ranks.buffer()] {
        reject_linear_memory_input(&buffer)?;
    }
    let control: SlicerControl = from_control(control)?;
    control.context.validate(TaskKind::SlicerMembers).map_err(error_to_js)?;
    control.params.validate_shape(values.length() as usize, states.length() as usize, text_ids.length() as usize,
        availability.length() as usize, text_offsets.length() as usize, text_bytes.length() as usize,
        text_ranks.length() as usize).map_err(js_error)?;
    let memory = control.params.memory_estimate().map_err(js_error)?;
    control.context.budget.check_input(memory.input_bytes).map_err(error_to_js)?;
    control.context.budget.check_output(memory.output_bytes).map_err(error_to_js)?;
    control.context.budget.check_working(memory.estimated_peak_bytes).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let maximum = control.context.budget.max_input_bytes;
    let values = copy_numbers_in(&values, maximum, &mut transfer)?;
    let states = copy_bytes_in(&states, maximum, &mut transfer)?;
    let text_ids = copy_indices_in(&text_ids, maximum, &mut transfer)?;
    let availability = copy_bytes_in(&availability, maximum, &mut transfer)?;
    let text_offsets = copy_indices_in(&text_offsets, maximum, &mut transfer)?;
    let text_bytes = copy_bytes_in(&text_bytes, maximum, &mut transfer)?;
    let text_ranks = copy_indices_in(&text_ranks, maximum, &mut transfer)?;
    let inner = SlicerMembers::new(values, states, text_ids, availability, text_offsets, text_bytes, text_ranks,
        control.params, control.context.budget.max_working_bytes as usize).map_err(js_error)?;
    Ok(TypedSlicerMembers { inner: Some(inner), context: control.context, transfer })
}

#[wasm_bindgen]
impl TypedSlicerMembers {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        match self.inner.as_mut().ok_or_else(consumed)?.step(budget) {
            Ok(done) => Ok(done),
            Err(error) => { self.inner = None; Err(js_error(error)) },
        }
    }

    pub fn progress(&self) -> Result<f64, JsValue> {
        Ok(self.inner.as_ref().ok_or_else(consumed)?.progress())
    }

    pub fn finish(&mut self) -> Result<TypedSlicerMemberResult, JsValue> {
        let result = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        // Move only the requested output window. Drop retained input/hash/merge
        // buffers before exposing any independent JS output copies.
        self.inner = None;
        Ok(TypedSlicerMemberResult { result, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}

#[wasm_bindgen]
pub struct TypedSlicerMemberResult {
    result: SlicerMembersResult,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

impl TypedSlicerMemberResult {
    fn byte_length(&self) -> Result<usize, JsValue> {
        self.result.indices.len().checked_mul(4).and_then(|n| n.checked_add(self.result.has_data.len()))
            .ok_or_else(|| js_error(ComputeError::memory()))
    }
}

#[wasm_bindgen]
impl TypedSlicerMemberResult {
    /// Shared bridge acquires a synchronous immutable lease only after the
    /// JS destination allocation. Views never escape/await/reenter; each call
    /// rebuilds its view against current memory, including after memory.grow.
    pub fn copy_indices(&mut self) -> Result<Uint32Array, JsValue> {
        self.context.budget.check_output(self.byte_length()?).map_err(error_to_js)?;
        copy_indices_out(&self.result.indices, &mut self.transfer)
    }

    pub fn copy_has_data(&mut self) -> Result<Uint8Array, JsValue> {
        self.context.budget.check_output(self.byte_length()?).map_err(error_to_js)?;
        copy_bytes_out(&self.result.has_data, &mut self.transfer)
    }

    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            context: &'a TaskControl,
            byte_length: usize,
            index_length: usize,
            total: u32,
            date_count: u32,
            min_date: Option<f64>,
            max_date: Option<f64>,
            transfer: &'a BoundaryCopyStats,
        }
        let byte_length = self.byte_length()?;
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata { context: &self.context, byte_length, index_length: self.result.indices.len(),
            total: self.result.total, date_count: self.result.date_count, min_date: self.result.min_date,
            max_date: self.result.max_date, transfer: &self.transfer })
    }
}
