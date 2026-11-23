#![cfg(target_arch = "wasm32")]

use js_sys::{Float64Array, Object, Reflect, Uint8Array, WebAssembly};
use office_compute::{create_numeric_aggregate, create_numeric_filter, create_numeric_sort, compute_protocol_version};
use office_wasm_bridge::{CONTROL_VERSION, TaskBudget, TaskControl, TaskKind, from_control, to_control};
use wasm_bindgen::{JsCast, JsValue};
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_dedicated_worker);

fn context(kind: TaskKind) -> TaskControl {
    TaskControl { protocol_version: CONTROL_VERSION, request_id: "binding-test".into(),
        unit_id: "workbook-test".into(), base_revision: 8, kind, budget: TaskBudget::default() }
}
fn values() -> Float64Array {
    let values = Float64Array::new_with_length(3);
    values.set_index(0, 4.0); values.set_index(1, 7.0); values.set_index(2, 9.0);
    values
}
fn states() -> Uint8Array {
    let states = Uint8Array::new_with_length(3);
    states.fill(1, 0, 3);
    states
}
fn code(error: JsValue) -> String {
    let detail = Reflect::get(&error, &JsValue::from_str("error")).unwrap();
    Reflect::get(&detail, &JsValue::from_str("code")).unwrap().as_string().unwrap()
}

#[wasm_bindgen_test]
fn aggregate_owns_one_input_copy_and_returns_structured_statistics() {
    assert_eq!(compute_protocol_version(), 2);
    let input = values();
    let control = context(TaskKind::Aggregate);
    let mut aggregate = create_numeric_aggregate(input.clone(), states(), to_control(&control).unwrap()).unwrap();
    input.set_index(0, 999.0);
    assert_eq!(code(aggregate.result().unwrap_err()), "COMPUTE_NOT_READY");
    assert!(!aggregate.step(1).unwrap());
    assert!(aggregate.step(2).unwrap());
    let result: serde_json::Value = from_control(aggregate.result().unwrap()).unwrap();
    assert_eq!(result["context"]["baseRevision"], 8);
    assert_eq!(result["statistics"]["numericCount"], 3);
    assert_eq!(result["statistics"]["sum"], 20.0);
    assert_eq!(result["transfer"]["inputCopies"], 2);
    assert_eq!(result["transfer"]["inputBytes"], 27);
    assert_eq!(result["transfer"]["jsonConversions"], 0);
    assert_eq!(result["transfer"]["rustBufferClones"], 0);
}

#[wasm_bindgen_test]
fn sort_moves_indices_to_an_independent_result_and_rejects_double_finish() {
    let params = serde_json::json!({ "context": context(TaskKind::Sort), "rowCount": 3, "directions": [1], "nullsLast": true });
    let mut task = create_numeric_sort(values(), states(), to_control(&params).unwrap()).unwrap();
    assert_eq!(code(match task.finish() { Err(error) => error, Ok(_) => panic!("unfinished") }), "COMPUTE_NOT_READY");
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    assert_eq!(code(task.step(1).unwrap_err()), "COMPUTE_RESULT_CONSUMED");
    assert_eq!(code(match task.finish() { Err(error) => error, Ok(_) => panic!("already consumed") }), "COMPUTE_RESULT_CONSUMED");
    let copied = result.copy_indices().unwrap();
    assert_eq!(copied.to_vec(), vec![2, 1, 0]);
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&copied.buffer(), &memory.buffer()));
    memory.grow(0);
    assert_eq!(copied.to_vec(), vec![2, 1, 0]);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["byteLength"], 12);
    assert_eq!(metadata["transfer"]["outputCopies"], 1);
    assert_eq!(metadata["transfer"]["outputBytes"], 12);
    assert_eq!(metadata["transfer"]["rustBufferClones"], 0);
}

#[wasm_bindgen_test]
fn filter_returns_only_its_effective_index_region_and_validates_control_kind() {
    let params = serde_json::json!({ "context": context(TaskKind::Filter), "rowCount": 3,
        "operators": [4], "lower": [5], "upper": [0], "matchAll": true });
    let mut task = create_numeric_filter(values(), states(), to_control(&params).unwrap()).unwrap();
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    assert_eq!(result.copy_indices().unwrap().to_vec(), vec![1, 2]);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["byteLength"], 8);
    assert_eq!(metadata["transfer"]["outputBytes"], 8);
    let mut wrong = context(TaskKind::Sort);
    wrong.protocol_version = 1;
    let error = match create_numeric_aggregate(values(), states(), to_control(&wrong).unwrap()) {
        Ok(_) => panic!("protocol 1 must be rejected"), Err(error) => error,
    };
    assert_eq!(code(error), "WASM_PROTOCOL_MISMATCH");
}
