#![cfg(target_arch = "wasm32")]

use js_sys::{Object, Promise, Reflect, Uint8Array, WebAssembly};
use office_wasm_bridge::{
    BoundaryCopyStats, BridgeError, ByteLeaseBuffer, CanvasDomHost, EditorInputHost,
    JS_SAFE_INTEGER, copy_bytes_in, copy_bytes_out, error_to_js, from_control, to_control,
    yield_to_browser,
};
use serde::{Deserialize, Serialize};
use std::{cell::Cell, collections::BTreeMap, rc::Rc};
use wasm_bindgen::{JsCast, prelude::*};
use wasm_bindgen_futures::JsFuture;
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_browser);

fn error_code(error: JsValue) -> String {
    let detail = Reflect::get(&error, &JsValue::from_str("error")).unwrap();
    Reflect::get(&detail, &JsValue::from_str("code")).unwrap().as_string().unwrap()
}

#[wasm_bindgen_test]
fn structured_dto_preserves_objects_nulls_and_rejects_unsafe_integers() {
    #[derive(Debug, Serialize, Deserialize, PartialEq)]
    struct Payload { index: BTreeMap<String, u64>, absent: Option<String> }
    let payload = Payload { index: BTreeMap::from([("row".into(), 7)]), absent: None };
    let js = to_control(&payload).unwrap();
    assert!(Reflect::get(&js, &JsValue::from_str("absent")).unwrap().is_null());
    let index = Reflect::get(&js, &JsValue::from_str("index")).unwrap();
    assert_eq!(Reflect::get(&index, &JsValue::from_str("row")).unwrap().as_f64(), Some(7.0));
    assert_eq!(from_control::<Payload>(js).unwrap(), payload);
    assert_eq!(error_code(to_control(&(JS_SAFE_INTEGER + 1)).unwrap_err()), "WASM_CONTROL_SERIALIZATION");
}

#[wasm_bindgen_test]
fn typed_array_offsets_have_one_inbound_and_one_effective_outbound_copy() {
    let source = Uint8Array::new_with_length(8);
    for index in 0..8 { source.set_index(index, index as u8); }
    let mut stats = BoundaryCopyStats::default();
    let retained = copy_bytes_in(&source.subarray(2, 5), 3, &mut stats).unwrap();
    assert_eq!(retained, vec![2, 3, 4]);
    assert_eq!(stats.input_copies, 1);
    assert_eq!(stats.input_bytes, 3);
    let output = copy_bytes_out(&retained, &mut stats).unwrap();
    assert_eq!(output.to_vec(), retained);
    assert_eq!(output.byte_offset(), 0);
    assert_eq!(stats.output_copies, 1);
    assert_eq!(stats.output_bytes, 3);
    assert_eq!(stats.rust_buffer_clones, 0);
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&output.buffer(), &memory.buffer()));
    memory.grow(0);
    assert_eq!(output.to_vec(), vec![2, 3, 4]);
    assert_eq!(error_code(copy_bytes_in(&source, 2, &mut stats).unwrap_err()), "WASM_MEMORY_LIMIT");
}

#[wasm_bindgen_test]
fn leases_reject_reentry_release_stale_generation_and_memory_changes() {
    let mut owner = ByteLeaseBuffer::from_vec(vec![10, 20, 30], BoundaryCopyStats::default());
    let token = owner.begin_lease().unwrap();
    let view = owner.lease_view(token).unwrap();
    assert_eq!(view.get_index(1), 20);
    assert_eq!(error_code(owner.dispose().unwrap_err()), "WASM_LEASE_ACTIVE");
    assert_eq!(error_code(owner.copy_bytes().unwrap_err()), "WASM_LEASE_ACTIVE");
    assert_eq!(error_code(owner.metadata().unwrap_err()), "WASM_LEASE_ACTIVE");
    assert_eq!(error_code(owner.end_lease(token + 1).unwrap_err()), "WASM_LEASE_STALE");
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    memory.grow(0);
    assert_eq!(view.length(), 0);
    assert_eq!(error_code(owner.lease_view(token).unwrap_err()), "WASM_MEMORY_CHANGED");
    owner.end_lease(token).unwrap();
    assert_eq!(error_code(owner.lease_view(token).unwrap_err()), "WASM_LEASE_STALE");
    let next = owner.begin_lease().unwrap();
    assert_ne!(next, token);
    assert_eq!(owner.lease_view(next).unwrap().get_index(1), 20);
    owner.end_lease(next).unwrap();
    owner.dispose().unwrap();
    assert_eq!(error_code(owner.dispose().unwrap_err()), "WASM_BUFFER_DISPOSED");
}

#[wasm_bindgen_test(async)]
async fn future_yields_to_a_real_browser_timer_and_preserves_rejection_dto() {
    let window = web_sys::window().unwrap();
    let delivered = Rc::new(Cell::new(false));
    let state = delivered.clone();
    let callback = Closure::new(move || state.set(true));
    let timer = window.set_timeout_with_callback_and_timeout_and_arguments_0(
        callback.as_ref().unchecked_ref(), 0,
    ).unwrap();
    yield_to_browser().await.unwrap();
    assert!(delivered.get());
    window.clear_timeout_with_handle(timer);
    drop(callback);
    let rejected = JsFuture::from(Promise::reject(&error_to_js(
        BridgeError::new("TEST_REJECTION", "Preserved structured failure"),
    ))).await.unwrap_err();
    assert_eq!(error_code(rejected), "TEST_REJECTION");
}

#[wasm_bindgen_test]
fn input_host_preserves_ime_focus_context_utf16_and_listener_lifetime() {
    let document = web_sys::window().unwrap().document().unwrap();
    let input: web_sys::HtmlInputElement = document.create_element("input").unwrap().dyn_into().unwrap();
    input.set_type("text");
    input.set_value("A🧪B");
    document.body().unwrap().append_child(&input).unwrap();
    let context = serde_json::json!({ "unitId": "doc", "generation": 4, "modelRevision": 9 });
    let mut host = EditorInputHost::new(input.clone().into(), to_control(&context).unwrap()).unwrap();
    let range = serde_json::json!({ "context": context, "start": 1, "end": 3, "direction": "forward" });
    input.dispatch_event(&web_sys::Event::new("compositionstart").unwrap()).unwrap();
    assert!(!host.focus_and_select(to_control(&range).unwrap()).unwrap());
    input.dispatch_event(&web_sys::Event::new("compositionend").unwrap()).unwrap();
    assert!(host.focus_and_select(to_control(&range).unwrap()).unwrap());
    assert_eq!(input.selection_start().unwrap(), Some(1));
    assert_eq!(input.selection_end().unwrap(), Some(3));
    assert_eq!(document.active_element().unwrap(), input.clone().unchecked_into::<web_sys::Element>());
    let next = serde_json::json!({ "unitId": "doc", "generation": 5, "modelRevision": 10 });
    host.update_context(to_control(&next).unwrap()).unwrap();
    assert!(!host.focus_and_select(to_control(&range).unwrap()).unwrap());
    host.dispose().unwrap();
    input.dispatch_event(&web_sys::Event::new("compositionstart").unwrap()).unwrap();
    let status: serde_json::Value = from_control(host.status().unwrap()).unwrap();
    assert_eq!(status["disposed"], true);
    assert_eq!(status["composing"], false);
    input.remove();
}

#[wasm_bindgen_test]
fn canvas_host_sets_css_geometry_and_rejects_disposed_handle() {
    let document = web_sys::window().unwrap().document().unwrap();
    let canvas: web_sys::HtmlCanvasElement = document.create_element("canvas").unwrap().dyn_into().unwrap();
    let mut host = CanvasDomHost::new(canvas.clone()).unwrap();
    host.set_css_size(320.0, 240.0).unwrap();
    assert_eq!(canvas.style().get_property_value("width").unwrap(), "320px");
    assert_eq!(canvas.style().get_property_value("height").unwrap(), "240px");
    assert!(host.set_css_size(f64::NAN, 240.0).is_err());
    host.dispose();
    assert_eq!(error_code(host.set_css_size(1.0, 1.0).unwrap_err()), "WASM_DOM_DISPOSED");
}
