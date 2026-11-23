#![cfg(target_arch = "wasm32")]

use js_sys::{Reflect, Uint8Array};
use office_wasm_bridge::{
    BoundaryCopyStats, EditorInputHost, copy_bytes_in, copy_bytes_out,
    office_wasm_bridge_capabilities, yield_to_browser,
};
use std::{cell::Cell, rc::Rc};
use wasm_bindgen::{JsCast, prelude::*};
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_dedicated_worker);

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
