use crate::{ChartStatistics, ChartStatisticsParams, ChartStatisticsResult, ComputeError};
use js_sys::{Float64Array, Uint8Array, Uint32Array};
use office_wasm_bridge::{
    BoundaryCopyStats, BridgeError, TaskControl, TaskKind, copy_bytes_in, copy_indices_in,
    copy_indices_out, copy_numbers_in, copy_numbers_out, error_to_js, from_control,
    reject_linear_memory_input, to_control,
};
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;

fn js_error(error: ComputeError) -> JsValue {
    error_to_js(BridgeError::new(error.code, error.message))
}

fn consumed() -> JsValue {
    js_error(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Chart calculation was already consumed"))
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ChartControl { context: TaskControl, params: ChartStatisticsParams }

#[wasm_bindgen]
pub struct NumericChartStatistics {
    inner: Option<ChartStatistics>,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

#[wasm_bindgen]
pub fn create_chart_statistics(values: Float64Array, states: Uint8Array, category_ids: Uint32Array,
    control: JsValue) -> Result<NumericChartStatistics, JsValue> {
    reject_linear_memory_input(&values.buffer())?;
    reject_linear_memory_input(&states.buffer())?;
    reject_linear_memory_input(&category_ids.buffer())?;
    let control: ChartControl = from_control(control)?;
    control.context.validate(TaskKind::ChartStatistics).map_err(error_to_js)?;
    control.params.validate_shape(values.length() as usize, states.length() as usize,
        category_ids.length() as usize).map_err(js_error)?;
    let memory = control.params.memory_estimate().map_err(js_error)?;
    control.context.budget.check_input(memory.input_bytes).map_err(error_to_js)?;
    control.context.budget.check_output(memory.output_bytes).map_err(error_to_js)?;
    control.context.budget.check_working(memory.estimated_peak_bytes).map_err(error_to_js)?;
    let mut transfer = BoundaryCopyStats { control_dto_conversions: 1, ..Default::default() };
    let values = copy_numbers_in(&values, control.context.budget.max_input_bytes, &mut transfer)?;
    let states = copy_bytes_in(&states, control.context.budget.max_input_bytes, &mut transfer)?;
    let categories = copy_indices_in(&category_ids, control.context.budget.max_input_bytes, &mut transfer)?;
    let inner = ChartStatistics::new(values, states, categories, control.params,
        control.context.budget.max_working_bytes as usize).map_err(js_error)?;
    Ok(NumericChartStatistics { inner: Some(inner), context: control.context, transfer })
}

#[wasm_bindgen]
impl NumericChartStatistics {
    pub fn step(&mut self, jobs: u32) -> Result<bool, JsValue> {
        self.inner.as_mut().ok_or_else(consumed)?.step(jobs).map_err(js_error)
    }

    pub fn progress(&self) -> Result<f64, JsValue> {
        Ok(self.inner.as_ref().ok_or_else(consumed)?.progress())
    }

    pub fn finish(&mut self) -> Result<NumericChartResult, JsValue> {
        let result = self.inner.as_mut().ok_or_else(consumed)?.take_result().map_err(js_error)?;
        // Retained inputs and job scratch are released before JS output copies.
        self.inner = None;
        Ok(NumericChartResult { result, context: self.context.clone(), transfer: self.transfer.clone() })
    }
}

#[wasm_bindgen]
pub struct NumericChartResult {
    result: ChartStatisticsResult,
    context: TaskControl,
    transfer: BoundaryCopyStats,
}

impl NumericChartResult {
    fn byte_length(&self) -> Result<usize, JsValue> {
        self.result.values.len().checked_mul(8)
            .and_then(|bytes| bytes.checked_add(self.result.indices.len().checked_mul(4)?))
            .ok_or_else(|| js_error(ComputeError::memory()))
    }
}

#[wasm_bindgen]
impl NumericChartResult {
    pub fn copy_values(&mut self) -> Result<Float64Array, JsValue> {
        self.context.budget.check_output(self.byte_length()?).map_err(error_to_js)?;
        copy_numbers_out(&self.result.values, &mut self.transfer)
    }

    pub fn copy_indices(&mut self) -> Result<Uint32Array, JsValue> {
        self.context.budget.check_output(self.byte_length()?).map_err(error_to_js)?;
        copy_indices_out(&self.result.indices, &mut self.transfer)
    }

    pub fn metadata(&mut self) -> Result<JsValue, JsValue> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct Metadata<'a> {
            context: &'a TaskControl,
            value_length: usize,
            index_length: usize,
            byte_length: usize,
            segments: &'a [crate::ChartSegment],
            transfer: &'a BoundaryCopyStats,
        }
        let byte_length = self.byte_length()?;
        self.transfer.control_dto_conversions += 1;
        to_control(&Metadata { context: &self.context, value_length: self.result.values.len(),
            index_length: self.result.indices.len(), byte_length, segments: &self.result.segments,
            transfer: &self.transfer })
    }
}
