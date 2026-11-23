#![cfg(target_arch = "wasm32")]

use js_sys::{Array, Float64Array, Object, Reflect, Uint8Array, Uint32Array, WebAssembly};
use office_wasm_bridge::{
    BoundaryCopyStats, EditorInputHost, copy_bytes_in, copy_bytes_out,
    copy_indices_in, copy_indices_out, copy_numbers_in, copy_numbers_out,
    office_wasm_bridge_capabilities, yield_to_browser,
};
use std::{cell::Cell, rc::Rc};
use wasm_bindgen::{JsCast, prelude::*};
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_dedicated_worker);

#[wasm_bindgen]
extern "C" {
    #[wasm_bindgen(catch, js_name = structuredClone)]
    fn clone_with_transfer(value: &JsValue, options: &JsValue) -> Result<JsValue, JsValue>;
}

fn error_code(error: JsValue) -> String {
    let detail = Reflect::get(&error, &JsValue::from_str("error")).unwrap();
    Reflect::get(&detail, &JsValue::from_str("code")).unwrap().as_string().unwrap()
}

fn detach(value: &JsValue, buffer: &JsValue) {
    let transfers = Array::new();
    transfers.push(buffer);
    let options = Object::new();
    Reflect::set(&options, &JsValue::from_str("transfer"), &transfers).unwrap();
    clone_with_transfer(value, &options).unwrap();
}

#[wasm_bindgen_test(async)]
async fn worker_uses_worker_timer_and_never_requires_window_or_document() {
    assert!(web_sys::window().is_none());
    let worker: web_sys::WorkerGlobalScope = js_sys::global().dyn_into().unwrap();
    let delivered = Rc::new(Cell::new(false));
    let state = delivered.clone();
    let callback = Closure::new(move || state.set(true));
    let timer = worker.set_timeout_with_callback_and_timeout_and_arguments_0(
        callback.as_ref().unchecked_ref(), 0,
    ).unwrap();
    yield_to_browser().await.unwrap();
    assert!(delivered.get());
    worker.clear_timeout_with_handle(timer);
    drop(callback);
    let capabilities = office_wasm_bridge_capabilities().unwrap();
    assert_eq!(Reflect::get(&capabilities, &JsValue::from_str("protocolVersion")).unwrap().as_f64(), Some(2.0));
    assert_eq!(Reflect::get(&capabilities, &JsValue::from_str("domHosts")).unwrap().as_bool(), Some(false));
    let error = match EditorInputHost::new(JsValue::UNDEFINED, JsValue::UNDEFINED) {
        Ok(_) => panic!("Worker must reject DOM input hosts"),
        Err(error) => error,
    };
    let nested = Reflect::get(&error, &JsValue::from_str("error")).unwrap();
    assert_eq!(Reflect::get(&nested, &JsValue::from_str("code")).unwrap().as_string().as_deref(), Some("WASM_DOM_UNAVAILABLE"));
}

#[wasm_bindgen_test]
fn worker_bulk_buffers_remain_independent_across_async_boundaries() {
    let source = Uint8Array::new_with_length(3);
    source.set_index(0, 42);
    let mut stats = BoundaryCopyStats::default();
    let retained = copy_bytes_in(&source, 3, &mut stats).unwrap();
    let output = copy_bytes_out(&retained, &mut stats).unwrap();
    assert_eq!(output.get_index(0), 42);
    source.set_index(0, 99);
    assert_eq!(output.get_index(0), 42);
    assert_eq!((stats.input_copies, stats.input_bytes, stats.output_copies, stats.output_bytes), (1, 3, 1, 3));
}

#[wasm_bindgen_test]
fn numeric_and_index_subviews_copy_only_the_effective_region_and_survive_growth() {
    let source = Float64Array::new_with_length(5);
    for (index, value) in [-99.0, 1.25, -2.5, f64::MAX, 99.0].into_iter().enumerate() {
        source.set_index(index as u32, value);
    }
    let index_source = Uint32Array::new_with_length(4);
    for (index, value) in [99, 7, u32::MAX, 99].into_iter().enumerate() {
        index_source.set_index(index as u32, value);
    }
    let mut stats = BoundaryCopyStats::default();
    let numbers = copy_numbers_in(&source.subarray(1, 4), 24, &mut stats).unwrap();
    let indices = copy_indices_in(&index_source.subarray(1, 3), 8, &mut stats).unwrap();
    source.set_index(1, 0.0);
    index_source.set_index(1, 0);
    assert_eq!(numbers, vec![1.25, -2.5, f64::MAX]);
    assert_eq!(indices, vec![7, u32::MAX]);
    let output_values = copy_numbers_out(&numbers, &mut stats).unwrap();
    let output_indices = copy_indices_out(&indices, &mut stats).unwrap();
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&output_values.buffer(), &memory.buffer()));
    assert!(!Object::is(&output_indices.buffer(), &memory.buffer()));
    memory.grow(0);
    assert_eq!(output_values.to_vec(), numbers);
    assert_eq!(output_indices.to_vec(), indices);
    assert_eq!((stats.input_copies, stats.input_bytes, stats.output_copies, stats.output_bytes), (2, 32, 2, 32));
    assert_eq!(stats.rust_buffer_clones, 0);
}

#[wasm_bindgen_test]
fn numeric_budgets_reject_before_copy_and_valid_empty_buffers_are_not_detached() {
    let mut stats = BoundaryCopyStats::default();
    assert_eq!(error_code(copy_numbers_in(&Float64Array::new_with_length(2), 15, &mut stats).unwrap_err()), "WASM_MEMORY_LIMIT");
    assert_eq!(error_code(copy_indices_in(&Uint32Array::new_with_length(2), 7, &mut stats).unwrap_err()), "WASM_MEMORY_LIMIT");
    assert!(copy_numbers_in(&Float64Array::new_with_length(0), 0, &mut stats).unwrap().is_empty());
    assert!(copy_indices_in(&Uint32Array::new_with_length(0), 0, &mut stats).unwrap().is_empty());
    assert_eq!((stats.input_copies, stats.input_bytes), (0, 0));
}

#[wasm_bindgen_test]
fn transferred_inputs_are_rejected_instead_of_becoming_successful_empty_copies() {
    let values = Float64Array::new_with_length(2);
    let indices = Uint32Array::new_with_length(2);
    let bytes = Uint8Array::new_with_length(2);
    detach(values.as_ref(), &values.buffer());
    detach(indices.as_ref(), &indices.buffer());
    detach(bytes.as_ref(), &bytes.buffer());
    let mut stats = BoundaryCopyStats::default();
    assert_eq!(error_code(copy_numbers_in(&values, 16, &mut stats).unwrap_err()), "WASM_BUFFER_DETACHED");
    assert_eq!(error_code(copy_indices_in(&indices, 8, &mut stats).unwrap_err()), "WASM_BUFFER_DETACHED");
    assert_eq!(error_code(copy_bytes_in(&bytes, 2, &mut stats).unwrap_err()), "WASM_BUFFER_DETACHED");
    assert_eq!((stats.input_copies, stats.input_bytes), (0, 0));
}

#[wasm_bindgen_test]
fn every_bulk_input_rejects_the_current_linear_memory_buffer() {
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    let mut stats = BoundaryCopyStats::default();
    let numbers = Float64Array::new_with_byte_offset_and_length(&memory.buffer(), 0, 1);
    assert_eq!(error_code(copy_numbers_in(&numbers, 8, &mut stats).unwrap_err()), "WASM_INPUT_ALIAS");
    let indices = Uint32Array::new_with_byte_offset_and_length(&memory.buffer(), 0, 1);
    assert_eq!(error_code(copy_indices_in(&indices, 4, &mut stats).unwrap_err()), "WASM_INPUT_ALIAS");
    let bytes = Uint8Array::new_with_byte_offset_and_length(&memory.buffer(), 0, 1);
    assert_eq!(error_code(copy_bytes_in(&bytes, 1, &mut stats).unwrap_err()), "WASM_INPUT_ALIAS");
    assert_eq!((stats.input_copies, stats.input_bytes), (0, 0));
}
