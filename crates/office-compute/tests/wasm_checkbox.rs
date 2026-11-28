#![cfg(target_arch = "wasm32")]
// Real Worker binding drafts; intentionally unexecuted until unified validation.
use js_sys::{Object, Uint8Array, Uint32Array, WebAssembly};
use office_compute::{CheckboxCellsParams, create_checkbox_cells, compute_yield};
use office_wasm_bridge::{TaskControl, TaskBudget, TaskKind, CONTROL_VERSION, from_control, to_control};
use wasm_bindgen::{JsCast, JsValue};
use wasm_bindgen_test::*;
wasm_bindgen_test_configure!(run_in_dedicated_worker);
fn control() -> JsValue {
    to_control(&serde_json::json!({ "context": TaskControl { protocol_version: CONTROL_VERSION, request_id: "checkbox-binding".into(),
        unit_id: "book".into(), base_revision: 4, kind: TaskKind::CheckboxCells, budget: TaskBudget::default() },
        "params": CheckboxCellsParams { row_count:64,column_count:64,range_count:1,merge_count:0,existing_count:2,covered_cells:4096,max_output_runs:64 } })).unwrap()
}
fn input() -> (Uint32Array,Uint32Array,Uint32Array,Uint8Array) {
    (Uint32Array::from(&[0,63,0,63][..]),Uint32Array::new_with_length(0),Uint32Array::from(&[0,0,1,1][..]),Uint8Array::from(&[2,0][..]))
}
#[wasm_bindgen_test]
fn typed_bulk_input_is_copied_once_and_owned_output_survives_growth_release_and_repeated_copies() {
    let (r,m,p,a)=input(); let mut task=create_checkbox_cells(r.clone(),m,p,a.clone(),control()).unwrap();
    r.set_index(1,0); a.set_index(0,0); assert!(task.finish().is_err());
    while !task.step(17).unwrap() {}
    let mut result=task.finish().unwrap(); assert!(task.step(1).is_err()); assert!(task.finish().is_err());
    let first=result.copy_runs().unwrap(); let meta:serde_json::Value=from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(meta["changedCells"],4095); assert_eq!(meta["transfer"]["inputCopies"],3); assert_eq!(meta["transfer"]["inputBytes"],34);
    assert_eq!(meta["transfer"]["outputCopies"],1); assert_eq!(meta["transfer"]["outputBytes"],first.byte_length());
    assert_eq!(meta["transfer"]["controlDtoConversions"],2); assert_eq!(meta["transfer"]["jsonConversions"],0); assert_eq!(meta["transfer"]["rustBufferClones"],0);
    let memory:WebAssembly::Memory=wasm_bindgen::memory().unchecked_into(); assert!(!Object::is(&first.buffer(),&memory.buffer()));
    memory.grow(0); let second=result.copy_runs().unwrap(); assert_eq!(first.to_vec(),second.to_vec());
    drop(result); assert_eq!(first.to_vec(),second.to_vec()); assert!(!Object::is(&first.buffer(),&second.buffer()));
}
#[wasm_bindgen_test]
fn alias_wrong_control_unknown_fields_and_budget_fail_before_output_is_exposed() {
    let (r,m,p,a)=input(); let memory:WebAssembly::Memory=wasm_bindgen::memory().unchecked_into();
    let alias=Uint32Array::new_with_byte_offset_and_length(&memory.buffer(),0,4);
    assert!(create_checkbox_cells(alias,m.clone(),p.clone(),a.clone(),control()).is_err());
    let mut options:serde_json::Value=from_control(control()).unwrap(); options["params"]["unknown"]=serde_json::json!(1);
    assert!(create_checkbox_cells(r.clone(),m.clone(),p.clone(),a.clone(),to_control(&options).unwrap()).is_err());
    options["params"].as_object_mut().unwrap().remove("unknown"); options["context"]["kind"]=serde_json::json!("sort");
    assert!(create_checkbox_cells(r.clone(),m.clone(),p.clone(),a.clone(),to_control(&options).unwrap()).is_err());
    options["context"]["kind"]=serde_json::json!("checkboxCells"); options["context"]["budget"]["maxWorkingBytes"]=serde_json::json!(1);
    assert!(create_checkbox_cells(r,m,p,a,to_control(&options).unwrap()).is_err());
}
#[wasm_bindgen_test(async)]
async fn macro_task_yield_allows_partial_tasks_to_be_released_and_next_tasks_to_finish() {
    let (r,m,p,a)=input(); let mut task=create_checkbox_cells(r,m,p,a,control()).unwrap(); assert!(!task.step(1).unwrap());
    compute_yield().await.unwrap(); drop(task);
    let (r,m,p,a)=input(); let mut task=create_checkbox_cells(r,m,p,a,control()).unwrap();
    while !task.step(128).unwrap() { compute_yield().await.unwrap(); }
    assert!(!task.finish().unwrap().copy_runs().unwrap().to_vec().is_empty());
}
