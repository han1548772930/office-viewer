#![cfg(target_arch = "wasm32")]
// Real browser/Worker binding drafts; no runner has been executed.
use js_sys::{Object, Uint8Array, Uint32Array, WebAssembly};
use office_compute::{ValidationListQueryParams, compute_yield, create_validation_list_query};
use office_wasm_bridge::{CONTROL_VERSION, TaskBudget, TaskControl, TaskKind, from_control, to_control};
use wasm_bindgen::{JsCast, JsValue};
use wasm_bindgen_test::*;
wasm_bindgen_test_configure!(run_in_dedicated_worker);

fn control() -> JsValue {
    to_control(&serde_json::json!({ "context": TaskControl { protocol_version: CONTROL_VERSION, request_id: "validation-list-binding".into(),
        unit_id: "book".into(), base_revision: 4, kind: TaskKind::ValidationListQuery, budget: TaskBudget::default() },
        "params": ValidationListQueryParams { row_count: 2, text_byte_length: 8, query_byte_length: 2, start: 0, count: 2 } })).unwrap()
}
fn input() -> (Uint32Array, Uint8Array, Uint8Array) {
    (Uint32Array::from(&[0, 4, 8][..]), Uint8Array::from(&[97, 0, 98, 0, 97, 0, 99, 0][..]), Uint8Array::from(&[97, 0][..]))
}

#[wasm_bindgen_test]
fn owned_window_survives_source_mutation_growth_release_and_repeated_copies() {
    let (offsets, text, query) = input();
    let mut task = create_validation_list_query(offsets.clone(), text.clone(), query.clone(), control()).unwrap();
    offsets.set_index(1, 0); text.set_index(0, 122); query.set_index(0, 122);
    assert!(task.finish().is_err());
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    assert!(task.step(1).is_err()); assert!(task.finish().is_err());
    let first = result.copy_indices().unwrap();
    let meta: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(first.to_vec(), vec![0, 1]); assert_eq!(meta["total"], 2);
    assert_eq!(meta["transfer"]["inputCopies"], 3); assert_eq!(meta["transfer"]["inputBytes"], 22);
    assert_eq!(meta["transfer"]["outputCopies"], 1); assert_eq!(meta["transfer"]["outputBytes"], 8);
    assert_eq!(meta["transfer"]["jsonConversions"], 0); assert_eq!(meta["transfer"]["rustBufferClones"], 0);
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&first.buffer(), &memory.buffer()));
    memory.grow(0);
    let second = result.copy_indices().unwrap(); drop(result);
    assert_eq!(first.to_vec(), vec![0, 1]); assert_eq!(first.to_vec(), second.to_vec());
    assert!(!Object::is(&first.buffer(), &second.buffer()));
}

#[wasm_bindgen_test]
fn controls_budget_and_linear_memory_aliases_fail_before_output() {
    let (offsets, text, query) = input();
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    let alias = Uint8Array::new_with_byte_offset_and_length(&memory.buffer(), 0, 8);
    assert!(create_validation_list_query(offsets.clone(), alias, query.clone(), control()).is_err());
    for (field, value) in [("kind", serde_json::json!("sort")), ("protocolVersion", serde_json::json!(1))] {
        let mut options: serde_json::Value = from_control(control()).unwrap(); options["context"][field] = value;
        assert!(create_validation_list_query(offsets.clone(), text.clone(), query.clone(), to_control(&options).unwrap()).is_err());
    }
    let mut options: serde_json::Value = from_control(control()).unwrap(); options["params"]["unknown"] = serde_json::json!(1);
    assert!(create_validation_list_query(offsets.clone(), text.clone(), query.clone(), to_control(&options).unwrap()).is_err());
    options["params"].as_object_mut().unwrap().remove("unknown"); options["context"]["budget"]["maxWorkingBytes"] = serde_json::json!(1);
    assert!(create_validation_list_query(offsets, text, query, to_control(&options).unwrap()).is_err());
}

#[wasm_bindgen_test(async)]
async fn partial_prefix_and_dictionary_work_can_yield_release_and_restart() {
    let (offsets, text, query) = input();
    let mut task = create_validation_list_query(offsets, text, query, control()).unwrap();
    assert!(!task.step(1).unwrap()); compute_yield().await.unwrap(); drop(task);
    let (offsets, text, query) = input();
    let mut task = create_validation_list_query(offsets, text, query, control()).unwrap();
    while !task.step(1).unwrap() { compute_yield().await.unwrap(); }
    assert_eq!(task.finish().unwrap().copy_indices().unwrap().to_vec(), vec![0, 1]);
}
