#![cfg(target_arch = "wasm32")]

use js_sys::{Float64Array, Object, Reflect, Uint8Array, Uint32Array, WebAssembly};
use office_compute::{create_chart_statistics, compute_protocol_version};
use office_wasm_bridge::{CONTROL_VERSION, TaskBudget, TaskControl, TaskKind, from_control, to_control};
use wasm_bindgen::{JsCast, JsValue};
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_dedicated_worker);

// Binding drafts for the unified validation phase. Native chart arithmetic is
// covered in src/chart/tests.rs; this file exercises real JS buffers and DTOs.
fn context() -> TaskControl {
    TaskControl { protocol_version: CONTROL_VERSION, request_id: "chart-binding".into(),
        unit_id: "chart-workbook".into(), base_revision: 11, kind: TaskKind::ChartStatistics,
        budget: TaskBudget::default() }
}

fn inputs() -> (Float64Array, Uint8Array, Uint32Array) {
    let values = Float64Array::new_with_length(3);
    for (index, value) in [4.0, 7.0, 9.0].into_iter().enumerate() { values.set_index(index as u32, value); }
    let states = Uint8Array::new_with_length(3);
    states.fill(1, 0, 3);
    let categories = Uint32Array::new_with_length(3);
    for index in 0..3 { categories.set_index(index, index); }
    (values, states, categories)
}

fn params() -> serde_json::Value {
    serde_json::json!({ "rowCount": 3, "columnCount": 1, "categoryCount": 3,
        "series": [{ "field": 7, "seriesPosition": 0, "column": 0 }], "jobs": [
            { "kind": "histogram", "seriesIndex": 0, "slot": 2, "binning": { "mode": "count", "count": 2 } },
            { "kind": "boxplot", "seriesIndex": 0, "slot": 5, "quartileMethod": "inclusive", "whiskerMultiplier": 1.5 },
        ] })
}

fn code(error: JsValue) -> String {
    let nested = Reflect::get(&error, &JsValue::from_str("error")).unwrap();
    Reflect::get(&nested, &JsValue::from_str("code")).unwrap().as_string().unwrap()
}

#[wasm_bindgen_test]
fn chart_controls_preserve_slots_and_one_owned_copy_per_input() {
    assert_eq!(compute_protocol_version(), 2);
    let (values, states, categories) = inputs();
    let control = serde_json::json!({ "context": context(), "params": params() });
    let mut task = create_chart_statistics(values.clone(), states.clone(), categories.clone(), to_control(&control).unwrap()).unwrap();
    values.set_index(0, 999.0);
    states.set_index(0, 0);
    categories.set_index(0, 2);
    assert_eq!(code(match task.finish() { Ok(_) => panic!("unfinished"), Err(error) => error }), "COMPUTE_NOT_READY");
    assert!(!task.step(1).unwrap());
    assert_eq!(task.progress().unwrap(), 0.5);
    assert!(task.step(1).unwrap());
    let mut result = task.finish().unwrap();
    assert_eq!(code(task.progress().unwrap_err()), "COMPUTE_RESULT_CONSUMED");
    assert_eq!(code(task.step(1).unwrap_err()), "COMPUTE_RESULT_CONSUMED");
    assert_eq!(code(match task.finish() { Ok(_) => panic!("consumed"), Err(error) => error }), "COMPUTE_RESULT_CONSUMED");
    let copied_values = result.copy_values().unwrap();
    let copied_indices = result.copy_indices().unwrap();
    assert_eq!(copied_values.to_vec(), vec![4.0, 6.5, 6.5, 9.0, 4.0, 5.5, 7.0, 8.0, 9.0]);
    assert_eq!(copied_indices.to_vec(), vec![1, 1, 1, 2]);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["context"]["kind"], "chartStatistics");
    assert_eq!(metadata["context"]["baseRevision"], 11);
    assert_eq!(metadata["segments"][0]["slot"], 2);
    assert_eq!(metadata["segments"][1]["slot"], 5);
    assert_eq!(metadata["segments"][0]["field"], 7);
    assert_eq!(metadata["segments"][0]["seriesPosition"], 0);
    assert_eq!(metadata["byteLength"], 88);
    assert_eq!(metadata["transfer"]["inputCopies"], 3);
    assert_eq!(metadata["transfer"]["inputBytes"], 39);
    assert_eq!(metadata["transfer"]["outputCopies"], 2);
    assert_eq!(metadata["transfer"]["outputBytes"], 88);
    assert_eq!(metadata["transfer"]["controlDtoConversions"], 2);
    assert_eq!(metadata["transfer"]["rustBufferClones"], 0);
    assert_eq!(metadata["transfer"]["jsonConversions"], 0);
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&copied_values.buffer(), &memory.buffer()));
    assert!(!Object::is(&copied_indices.buffer(), &memory.buffer()));
    drop(result);
    memory.grow(0);
    assert_eq!(copied_indices.to_vec(), vec![1, 1, 1, 2]);
    assert_eq!(copied_values.get_index(0), 4.0);
}

#[wasm_bindgen_test]
fn chart_protocol_kind_dimensions_and_budgets_reject_before_admission() {
    for (change, expected) in [
        (serde_json::json!({ "protocolVersion": 1 }), "WASM_PROTOCOL_MISMATCH"),
        (serde_json::json!({ "kind": "aggregate" }), "WASM_CONTROL_INVALID"),
        (serde_json::json!({ "budget": { "maxInputBytes": 1, "maxOutputBytes": 1024, "maxWorkingBytes": 16_777_216 } }), "WASM_MEMORY_LIMIT"),
        (serde_json::json!({ "budget": { "maxInputBytes": 1024, "maxOutputBytes": 1, "maxWorkingBytes": 16_777_216 } }), "WASM_MEMORY_LIMIT"),
        (serde_json::json!({ "budget": { "maxInputBytes": 1024, "maxOutputBytes": 1_048_576, "maxWorkingBytes": 1 } }), "WASM_MEMORY_LIMIT"),
    ] {
        let mut control = serde_json::json!({ "context": context(), "params": params() });
        for (key, value) in change.as_object().unwrap() { control["context"][key] = value.clone(); }
        let (values, states, categories) = inputs();
        let error = match create_chart_statistics(values.clone(), states, categories, to_control(&control).unwrap()) {
            Ok(_) => panic!("invalid control accepted"), Err(error) => error,
        };
        assert_eq!(code(error), expected);
        assert_eq!(values.to_vec(), vec![4.0, 7.0, 9.0]);
    }
    let (values, states, _) = inputs();
    let control = serde_json::json!({ "context": context(), "params": params() });
    let error = match create_chart_statistics(values, states, Uint32Array::new_with_length(2), to_control(&control).unwrap()) {
        Ok(_) => panic!("invalid shape accepted"), Err(error) => error,
    };
    assert_eq!(code(error), "COMPUTE_INVALID_INPUT");
}

#[wasm_bindgen_test]
fn empty_chart_jobs_have_empty_independent_results_and_no_bulk_copy_count() {
    let empty = serde_json::json!({ "rowCount": 0, "columnCount": 1, "categoryCount": 0,
        "series": [{ "field": 7, "seriesPosition": 0, "column": 0 }],
        "jobs": [{ "kind": "boxplot", "seriesIndex": 0, "slot": 0, "quartileMethod": "inclusive", "whiskerMultiplier": 1.5 }] });
    let control = serde_json::json!({ "context": context(), "params": empty });
    let mut task = create_chart_statistics(Float64Array::new_with_length(0), Uint8Array::new_with_length(0),
        Uint32Array::new_with_length(0), to_control(&control).unwrap()).unwrap();
    assert!(task.step(1).unwrap());
    let mut result = task.finish().unwrap();
    assert_eq!(result.copy_values().unwrap().length(), 0);
    assert_eq!(result.copy_indices().unwrap().length(), 0);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["segments"][0]["available"], false);
    assert_eq!(metadata["segments"][0]["validSampleCount"], 0);
    assert_eq!(metadata["transfer"]["inputCopies"], 0);
    assert_eq!(metadata["transfer"]["outputCopies"], 0);
}
