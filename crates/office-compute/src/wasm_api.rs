use crate::{ComputeError, NumericAggregate, NumericFilter, StableNumericSort};
use wasm_bindgen::prelude::*;

fn js_error(error: ComputeError) -> JsValue {
    JsValue::from_str(
        &serde_json::json!({ "code": error.code, "message": error.message }).to_string(),
    )
}

#[wasm_bindgen]
pub fn compute_version() -> String {
    env!("CARGO_PKG_VERSION").into()
}

#[wasm_bindgen]
pub struct NumericAggregator {
    inner: NumericAggregate,
}

#[wasm_bindgen]
impl NumericAggregator {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        Self {
            inner: NumericAggregate::default(),
        }
    }
    pub fn push(&mut self, values: &[f64], states: &[u8]) -> Result<(), JsValue> {
        self.inner.push(values, states).map_err(js_error)
    }
    pub fn result(&self) -> Vec<f64> {
        self.inner.packed_result()
    }
}

impl Default for NumericAggregator {
    fn default() -> Self {
        Self::new()
    }
}

#[wasm_bindgen]
pub struct NumericIndexSort {
    inner: StableNumericSort,
}

#[wasm_bindgen]
impl NumericIndexSort {
    #[wasm_bindgen(constructor)]
    pub fn new(
        values: Vec<f64>,
        states: Vec<u8>,
        rows: u32,
        directions: Vec<u8>,
        nulls_last: bool,
    ) -> Result<Self, JsValue> {
        Ok(Self {
            inner: StableNumericSort::new(values, states, rows as usize, directions, nulls_last)
                .map_err(js_error)?,
        })
    }
    pub fn step(&mut self, budget: u32) -> bool {
        self.inner.step(budget)
    }
    pub fn progress(&self) -> f64 {
        self.inner.progress()
    }
    pub fn finish(&mut self) -> Result<Vec<u32>, JsValue> {
        self.inner.take_result().map_err(js_error)
    }
}

#[wasm_bindgen]
pub struct NumericIndexFilter {
    inner: NumericFilter,
}

#[wasm_bindgen]
impl NumericIndexFilter {
    #[wasm_bindgen(constructor)]
    pub fn new(
        values: Vec<f64>,
        states: Vec<u8>,
        rows: u32,
        operators: Vec<u8>,
        lower: Vec<f64>,
        upper: Vec<f64>,
        all: bool,
    ) -> Result<Self, JsValue> {
        Ok(Self {
            inner: NumericFilter::new(values, states, rows as usize, operators, lower, upper, all)
                .map_err(js_error)?,
        })
    }
    pub fn step(&mut self, budget: u32) -> bool {
        self.inner.step(budget)
    }
    pub fn progress(&self) -> f64 {
        self.inner.progress()
    }
    pub fn finish(&mut self) -> Result<Vec<u32>, JsValue> {
        self.inner.take_result().map_err(js_error)
    }
}
