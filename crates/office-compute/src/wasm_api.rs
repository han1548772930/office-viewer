use crate::{ComputeError, NumericAggregate, NumericFilter, NumericStatistics, StableNumericSort};
use js_sys::{Float64Array, Uint8Array, Uint32Array};
use office_wasm_bridge::{
    BoundaryCopyStats, BridgeError, CONTROL_VERSION, TaskControl, TaskKind,
    copy_bytes_in, copy_indices_out, copy_numbers_in, error_to_js, from_control,
    reject_linear_memory_input, to_control, yield_to_browser,
};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

fn js_error(error: ComputeError) -> JsValue {
    error_to_js(BridgeError::new(error.code, error.message))
}

fn consumed() -> JsValue {
    error_to_js(BridgeError::new("COMPUTE_RESULT_CONSUMED", "This calculation was already consumed"))
}

fn validate_inputs(
    values: &Float64Array, states: &Uint8Array, context: &TaskControl, extra_bytes: usize,
) -> Result<(), JsValue> {
    let length = values.length() as usize;
    if length != states.length() as usize {
        return Err(js_error(ComputeError::invalid("Values and states lengths differ")));
    }
    if length > crate::MAX_ELEMENTS {
        return Err(js_error(ComputeError::memory()));
    }
    let bytes = length.checked_mul(9).ok_or_else(|| js_error(ComputeError::memory()))?;
    context.budget.check_input(bytes).map_err(error_to_js)?;
    context.budget.check_working(bytes.saturating_add(extra_bytes)).map_err(error_to_js)
}

fn retain_inputs(
    values: &Float64Array, states: &Uint8Array, context: &TaskControl, transfer: &mut BoundaryCopyStats,
) -> Result<(Vec<f64>, Vec<u8>), JsValue> {
    let values = copy_numbers_in(values, context.budget.max_input_bytes, transfer)?;
    let states = copy_bytes_in(states, context.budget.max_input_bytes, transfer)?;
    Ok((values, states))
}

#[wasm_bindgen]
pub fn compute_protocol_version() -> u32 { CONTROL_VERSION }

#[wasm_bindgen]
pub fn compute_version() -> String { env!("CARGO_PKG_VERSION").into() }

#[wasm_bindgen]
pub async fn compute_yield() -> Result<(), JsValue> { yield_to_browser().await }

#[wasm_bindgen]
pub struct NumericAggregator {
    inner: NumericAggregate,
    values: Vec<f64>,
    states: Vec<u8>,
    position: usize,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
pub fn create_numeric_aggregate(
    values: Float64Array, states: Uint8Array, control: JsValue,
) -> Result<NumericAggregator, JsValue> {
    reject_linear_memory_input(&values.buffer())?;
    reject_linear_memory_input(&states.buffer())?;
    let context: TaskControl = from_control(control)?;
    context.validate(TaskKind::Aggregate).map_err(error_to_js)?;
    validate_inputs(&values, &states, &context, 0)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let (values, states) = retain_inputs(&values, &states, &context, &mut transfer)?;
    // Validate before the first step, so a late invalid state cannot mutate a prefix.
    crate::error::validate_states(&states).map_err(js_error)?;
    Ok(NumericAggregator { inner: NumericAggregate::default(), values, states,
        position: 0, context, transfer })
}

#[wasm_bindgen]
impl NumericAggregator {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        let end = self.position.saturating_add(budget.clamp(1, 1_000_000) as usize)
            .min(self.values.len());
        self.inner.push(&self.values[self.position..end], &self.states[self.position..end])
            .map_err(js_error)?;
        self.position = end;
        Ok(end == self.values.len())
    }

    pub fn progress(&self) -> f64 {
        if self.values.is_empty() { 1.0 } else { self.position as f64 / self.values.len() as f64 }
    }

    pub fn result(&mut self) -> Result<JsValue, JsValue> {
        if self.position != self.values.len() {
            return Err(js_error(ComputeError::new("COMPUTE_NOT_READY", "Aggregation has not finished")));
        }
        #[derive(Serialize)]
        struct ResultDto<'a> {
            context: &'a TaskControl,
            statistics: NumericStatistics,
            transfer: &'a BoundaryCopyStats,
        }
        self.transfer.control_dto_conversions += 1;
        to_control(&ResultDto { context: &self.context, statistics: self.inner.statistics(),
            transfer: &self.transfer })
    }
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct SortControl {
    context: TaskControl,
    row_count: u32,
    directions: Vec<u8>,
    nulls_last: bool,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct FilterControl {
    context: TaskControl,
    row_count: u32,
    operators: Vec<u8>,
    lower: Vec<f64>,
    upper: Vec<f64>,
    match_all: bool,
}

#[wasm_bindgen]
pub struct NumericIndexResult {
    indices: Vec<u32>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
impl NumericIndexResult {
    /// No result Vec clone or glue copy. Only the effective index region is
    /// copied once into a JS-owned buffer; callers may transfer that buffer.
    pub fn copy_indices(&mut self) -> Result<Uint32Array, JsValue> {
        self.context.budget.check_output(self.indices.len() * 4).map_err(error_to_js)?;
        copy_indices_out(&self.indices, &mut self.transfer)
    }

    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            context: &'a TaskControl, byte_length: usize, transfer: &'a BoundaryCopyStats,
        }
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata { context: &self.context, byte_length: self.indices.len() * 4,
            transfer: &self.transfer })
    }
}

#[wasm_bindgen]
pub struct NumericIndexSort {
    inner: Option<StableNumericSort>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
pub fn create_numeric_sort(
    values: Float64Array, states: Uint8Array, control: JsValue,
) -> Result<NumericIndexSort, JsValue> {
    reject_linear_memory_input(&values.buffer())?;
    reject_linear_memory_input(&states.buffer())?;
    let params: SortControl = from_control(control)?;
    params.context.validate(TaskKind::Sort).map_err(error_to_js)?;
    if params.row_count as usize > crate::MAX_ROWS || params.directions.len() > crate::MAX_KEYS {
        return Err(js_error(ComputeError::memory()));
    }
    if params.directions.is_empty() || params.row_count as usize * params.directions.len() != values.length() as usize {
        return Err(js_error(ComputeError::invalid("Invalid column-major sort dimensions")));
    }
    let indices_bytes = params.row_count as usize * 8;
    validate_inputs(&values, &states, &params.context, indices_bytes + params.directions.len())?;
    params.context.budget.check_output(params.row_count as usize * 4).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let (values, states) = retain_inputs(&values, &states, &params.context, &mut transfer)?;
    let inner = StableNumericSort::new(values, states, params.row_count as usize,
        params.directions, params.nulls_last).map_err(js_error)?;
    Ok(NumericIndexSort { inner: Some(inner), context: params.context, transfer })
}

#[wasm_bindgen]
impl NumericIndexSort {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        Ok(self.inner.as_mut().ok_or_else(consumed)?.step(budget))
    }
    pub fn progress(&self) -> Result<f64, JsValue> {
        Ok(self.inner.as_ref().ok_or_else(consumed)?.progress())
    }
    pub fn finish(&mut self) -> Result<NumericIndexResult, JsValue> {
        let indices = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        // Drop retained inputs/scratch while moving the result Vec into its owner.
        self.inner = None;
        Ok(NumericIndexResult { indices, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}

#[wasm_bindgen]
pub struct NumericIndexFilter {
    inner: Option<NumericFilter>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
pub fn create_numeric_filter(
    values: Float64Array, states: Uint8Array, control: JsValue,
) -> Result<NumericIndexFilter, JsValue> {
    reject_linear_memory_input(&values.buffer())?;
    reject_linear_memory_input(&states.buffer())?;
    let params: FilterControl = from_control(control)?;
    params.context.validate(TaskKind::Filter).map_err(error_to_js)?;
    if params.row_count as usize > crate::MAX_ROWS || params.operators.len() > crate::MAX_KEYS {
        return Err(js_error(ComputeError::memory()));
    }
    if params.operators.is_empty() || params.row_count as usize * params.operators.len() != values.length() as usize
        || params.lower.len() != params.operators.len() || params.upper.len() != params.operators.len() {
        return Err(js_error(ComputeError::invalid("Invalid column-major filter dimensions or operands")));
    }
    let extra_bytes = params.row_count as usize * 4
        + params.operators.len() + (params.lower.len() + params.upper.len()) * 8;
    validate_inputs(&values, &states, &params.context, extra_bytes)?;
    params.context.budget.check_output(params.row_count as usize * 4).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let (values, states) = retain_inputs(&values, &states, &params.context, &mut transfer)?;
    let inner = NumericFilter::new(values, states, params.row_count as usize,
        params.operators, params.lower, params.upper, params.match_all).map_err(js_error)?;
    Ok(NumericIndexFilter { inner: Some(inner), context: params.context, transfer })
}

#[wasm_bindgen]
impl NumericIndexFilter {
    pub fn step(&mut self, budget: u32) -> Result<bool, JsValue> {
        Ok(self.inner.as_mut().ok_or_else(consumed)?.step(budget))
    }
    pub fn progress(&self) -> Result<f64, JsValue> {
        Ok(self.inner.as_ref().ok_or_else(consumed)?.progress())
    }
    pub fn finish(&mut self) -> Result<NumericIndexResult, JsValue> {
        let indices = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        self.inner = None;
        Ok(NumericIndexResult { indices, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}
