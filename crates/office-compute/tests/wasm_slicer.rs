#![cfg(target_arch = "wasm32")]

use js_sys::{Array, Float64Array, Function, Object, Reflect, Uint8Array, Uint32Array, WebAssembly};
use office_compute::{create_slicer_members, compute_yield, TypedSlicerMembers};
use office_wasm_bridge::{CONTROL_VERSION, TaskBudget, TaskControl, TaskKind, from_control, to_control};
use wasm_bindgen::{JsCast, JsValue};
use wasm_bindgen_test::*;

wasm_bindgen_test_configure!(run_in_dedicated_worker);

// Real wasm32 binding drafts. All remain unexecuted until unified validation.
fn context() -> TaskControl {
    TaskControl { protocol_version: CONTROL_VERSION, request_id: "slicer-binding".into(),
        unit_id: "slicer-book".into(), base_revision: 8, kind: TaskKind::SlicerMembers, budget: TaskBudget::default() }
}
fn params(mode: &str) -> serde_json::Value {
    serde_json::json!({ "mode": mode,"rowCount": 8,"dictionaryCount": 3,"textByteLength": 10,
        "start": 0,"count": 3,"locale": "en-US","sortOrder": "ascending","customListSort": false,
        "crossFilter": "showItemsWithNoData","hideItemsWithNoData": false })
}
type Inputs = (Float64Array,Uint8Array,Uint32Array,Uint8Array,Uint32Array,Uint8Array,Uint32Array);
fn inputs() -> Inputs {
    let values = Float64Array::new_with_length(8);
    for (index,value) in [0.0,1.0,1.0,0.0,0.0,1.0,0.0,10.5].into_iter().enumerate() { values.set_index(index as u32,value); }
    let states = Uint8Array::new_with_length(8);
    for (index,value) in [0,1,2,3,5,1,3,2].into_iter().enumerate() { states.set_index(index as u32,value); }
    let text_ids = Uint32Array::new_with_length(8);
    for (index,value) in [u32::MAX,u32::MAX,u32::MAX,0,1,u32::MAX,2,u32::MAX].into_iter().enumerate() { text_ids.set_index(index as u32,value); }
    let availability = Uint8Array::new_with_length(8);
    for (index,value) in [1,0,0,1,1,1,1,1].into_iter().enumerate() { availability.set_index(index as u32,value); }
    let offsets = Uint32Array::new_with_length(4);
    for (index,value) in [0,2,10,10].into_iter().enumerate() { offsets.set_index(index as u32,value); }
    let bytes = Uint8Array::new_with_length(10);
    for (index,value) in [49,0,35,0,78,0,47,0,65,0].into_iter().enumerate() { bytes.set_index(index as u32,value); }
    let ranks = Uint32Array::new_with_length(3);
    for (index,value) in [2,1,0].into_iter().enumerate() { ranks.set_index(index as u32,value); }
    (values,states,text_ids,availability,offsets,bytes,ranks)
}
fn query(input: Inputs, control: JsValue) -> Result<TypedSlicerMembers,JsValue> {
    create_slicer_members(input.0,input.1,input.2,input.3,input.4,input.5,input.6,control)
}
fn code(error: JsValue) -> String {
    let nested = Reflect::get(&error,&JsValue::from_str("error")).unwrap();
    Reflect::get(&nested,&JsValue::from_str("code")).unwrap().as_string().unwrap()
}
fn control() -> JsValue { to_control(&serde_json::json!({ "context": context(),"params": params("items") })).unwrap() }

#[wasm_bindgen_test]
fn owned_inputs_exact_types_or_availability_and_bounded_outputs_survive_growth_and_drop() {
    let input = inputs(); let mut task = query(input.clone(),control()).unwrap();
    input.0.set_index(1,999.0); input.1.set_index(2,1); input.3.set_index(5,0);
    assert_eq!(code(match task.finish() { Ok(_) => panic!("unfinished"),Err(error) => error }),"COMPUTE_NOT_READY");
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    assert_eq!(code(task.step(1).unwrap_err()),"COMPUTE_RESULT_CONSUMED");
    assert_eq!(code(match task.finish() { Ok(_) => panic!("double finish"),Err(error) => error }),"COMPUTE_RESULT_CONSUMED");
    let indices = result.copy_indices().unwrap(); let data = result.copy_has_data().unwrap();
    assert_eq!(indices.to_vec(),vec![0,1,2]); assert_eq!(data.to_vec(),vec![1,1,0]);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["total"],7); assert_eq!(metadata["dateCount"],2);
    assert_eq!(metadata["minDate"],1.0); assert_eq!(metadata["maxDate"],10.5);
    assert_eq!(metadata["context"]["baseRevision"],8); assert_eq!(metadata["byteLength"],15);
    assert_eq!(metadata["transfer"]["inputCopies"],7); assert_eq!(metadata["transfer"]["inputBytes"],8*14+3*8+4+10);
    assert_eq!(metadata["transfer"]["outputCopies"],2); assert_eq!(metadata["transfer"]["outputBytes"],15);
    assert_eq!(metadata["transfer"]["controlDtoConversions"],2); assert_eq!(metadata["transfer"]["rustBufferClones"],0);
    assert_eq!(metadata["transfer"]["jsonConversions"],0);
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    assert!(!Object::is(&indices.buffer(),&memory.buffer())); assert!(!Object::is(&data.buffer(),&memory.buffer()));
    drop(result); memory.grow(0);
    assert_eq!(indices.to_vec(),vec![0,1,2]); assert_eq!(data.to_vec(),vec![1,1,0]);
}

#[wasm_bindgen_test]
fn real_independent_window_buffers_transfer_and_detach_without_transferring_linear_memory() {
    let mut task = query(inputs(),control()).unwrap(); while !task.step(16).unwrap() {}
    let mut result = task.finish().unwrap(); let indices = result.copy_indices().unwrap(); let data = result.copy_has_data().unwrap();
    let packet = Object::new(); Reflect::set(&packet,&JsValue::from_str("indices"),&indices).unwrap();
    Reflect::set(&packet,&JsValue::from_str("hasData"),&data).unwrap();
    let transfer = Array::new(); transfer.push(&indices.buffer()); transfer.push(&data.buffer());
    let options = Object::new(); Reflect::set(&options,&JsValue::from_str("transfer"),&transfer).unwrap();
    let global = js_sys::global(); let clone: Function = Reflect::get(&global,&JsValue::from_str("structuredClone")).unwrap().dyn_into().unwrap();
    let received = clone.call2(&global,&packet,&options).unwrap();
    assert_eq!(indices.buffer().byte_length(),0); assert_eq!(data.buffer().byte_length(),0);
    let received_indices: Uint32Array = Reflect::get(&received,&JsValue::from_str("indices")).unwrap().dyn_into().unwrap();
    assert_eq!(received_indices.to_vec(),vec![0,1,2]);
}

#[wasm_bindgen_test]
fn protocol_kind_unknown_fields_dimensions_and_boundary_budgets_reject_before_admission() {
    for (change,expected) in [
        (serde_json::json!({"protocolVersion":1}),"WASM_PROTOCOL_MISMATCH"),
        (serde_json::json!({"kind":"aggregate"}),"WASM_CONTROL_INVALID"),
        (serde_json::json!({"budget":{"maxInputBytes":1,"maxOutputBytes":1024,"maxWorkingBytes":16_777_216}}),"WASM_MEMORY_LIMIT"),
        (serde_json::json!({"budget":{"maxInputBytes":1024,"maxOutputBytes":1,"maxWorkingBytes":16_777_216}}),"WASM_MEMORY_LIMIT"),
        (serde_json::json!({"budget":{"maxInputBytes":1024,"maxOutputBytes":1024,"maxWorkingBytes":1}}),"WASM_MEMORY_LIMIT"),
    ] {
        let mut value = serde_json::json!({"context":context(),"params":params("items")});
        for (key,entry) in change.as_object().unwrap() { value["context"][key] = entry.clone(); }
        assert_eq!(code(match query(inputs(),to_control(&value).unwrap()) { Ok(_) => panic!("invalid control"),Err(error) => error }),expected);
    }
    let mut value = serde_json::json!({"context":context(),"params":params("items")}); value["params"]["extra"] = serde_json::json!(true);
    assert!(query(inputs(),to_control(&value).unwrap()).is_err());
    let input = inputs(); input.1.set_index(7,6);
    let mut task = query(input,control()).unwrap();
    let mut failed = false; for _ in 0..30 { if task.step(1).is_err() { failed = true; break; } }
    assert!(failed); assert_eq!(code(match task.finish() { Ok(_) => panic!("partial result"),Err(error) => error }),"COMPUTE_RESULT_CONSUMED");
}

#[wasm_bindgen_test]
fn aliases_are_rejected_and_fresh_output_copies_rebuild_views_after_growth() {
    let input = inputs(); let options = control();
    let memory: WebAssembly::Memory = wasm_bindgen::memory().unchecked_into();
    let values = Float64Array::new_with_byte_offset_and_length(&memory.buffer(),0,8);
    let error = create_slicer_members(values,input.1,input.2,input.3,input.4,input.5,input.6,options).err().unwrap();
    assert_eq!(code(error),"WASM_INPUT_ALIAS");
    let mut task = query(inputs(),control()).unwrap(); while !task.step(16).unwrap() {}
    let mut result = task.finish().unwrap(); let first = result.copy_indices().unwrap();
    memory.grow(0); let second = result.copy_indices().unwrap();
    assert_eq!(first.to_vec(),second.to_vec()); assert!(!Object::is(&first.buffer(),&second.buffer()));
}

#[wasm_bindgen_test(async)]
async fn real_bridge_macro_task_yield_allows_partial_query_owners_to_be_released_and_reused() {
    let mut task = query(inputs(),control()).unwrap(); assert!(!task.step(1).unwrap());
    compute_yield().await.unwrap(); drop(task);
    let mut replacement = query(inputs(),control()).unwrap();
    while !replacement.step(2).unwrap() { compute_yield().await.unwrap(); }
    let mut result = replacement.finish().unwrap(); assert_eq!(result.copy_indices().unwrap().to_vec(),vec![0,1,2]);
}

#[wasm_bindgen_test]
fn typed_date_metadata_has_no_member_arrays_and_empty_copies_are_not_counted() {
    let mut options = params("metadata");
    options["dictionaryCount"] = serde_json::json!(0);
    options["textByteLength"] = serde_json::json!(0);
    let values = Float64Array::new_with_length(8);
    values.set_index(2,-2.25); values.set_index(7,10.5);
    let states = Uint8Array::new_with_length(8); states.set_index(2,2); states.set_index(7,2);
    let text_ids = Uint32Array::new_with_length(8);
    for row in 0..8 { text_ids.set_index(row,u32::MAX); }
    let available = Uint8Array::new_with_length(8);
    let mut task = query((values,states,text_ids,available,Uint32Array::new_with_length(1),
        Uint8Array::new_with_length(0),Uint32Array::new_with_length(0)),
        to_control(&serde_json::json!({"context":context(),"params":options})).unwrap()).unwrap();
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    let indices = result.copy_indices().unwrap(); let has_data = result.copy_has_data().unwrap();
    assert_eq!(indices.length(),0); assert_eq!(has_data.length(),0);
    assert!(!Object::is(&indices.buffer(),&has_data.buffer()));
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["total"],0); assert_eq!(metadata["dateCount"],2);
    assert_eq!(metadata["minDate"],-2.25); assert_eq!(metadata["maxDate"],10.5);
    assert_eq!(metadata["transfer"]["inputCopies"],5); assert_eq!(metadata["transfer"]["inputBytes"],8*14+4);
    assert_eq!(metadata["transfer"]["outputCopies"],0); assert_eq!(metadata["transfer"]["outputBytes"],0);
}

#[wasm_bindgen_test]
fn empty_sources_serialize_absent_date_bounds_as_null_and_empty_outputs_transfer() {
    let mut options = params("items");
    options["rowCount"] = serde_json::json!(0); options["dictionaryCount"] = serde_json::json!(0);
    options["textByteLength"] = serde_json::json!(0);
    let mut task = query((Float64Array::new_with_length(0),Uint8Array::new_with_length(0),
        Uint32Array::new_with_length(0),Uint8Array::new_with_length(0),Uint32Array::new_with_length(1),
        Uint8Array::new_with_length(0),Uint32Array::new_with_length(0)),
        to_control(&serde_json::json!({"context":context(),"params":options})).unwrap()).unwrap();
    assert!(task.step(1).unwrap());
    let mut result = task.finish().unwrap();
    let indices = result.copy_indices().unwrap(); let data = result.copy_has_data().unwrap();
    let metadata = result.metadata().unwrap();
    assert!(Reflect::get(&metadata,&JsValue::from_str("minDate")).unwrap().is_null());
    assert!(Reflect::get(&metadata,&JsValue::from_str("maxDate")).unwrap().is_null());
    let metadata: serde_json::Value = from_control(metadata).unwrap();
    assert_eq!(metadata["transfer"]["inputCopies"],1); assert_eq!(metadata["transfer"]["inputBytes"],4);
    assert_eq!(metadata["transfer"]["outputCopies"],0); assert_eq!(metadata["dateCount"],0);
    let packet = Object::new(); Reflect::set(&packet,&JsValue::from_str("indices"),&indices).unwrap();
    Reflect::set(&packet,&JsValue::from_str("hasData"),&data).unwrap();
    let transfer = Array::new(); transfer.push(&indices.buffer()); transfer.push(&data.buffer());
    let options = Object::new(); Reflect::set(&options,&JsValue::from_str("transfer"),&transfer).unwrap();
    let global = js_sys::global(); let clone: Function = Reflect::get(&global,&JsValue::from_str("structuredClone")).unwrap().dyn_into().unwrap();
    let received = clone.call2(&global,&packet,&options).unwrap();
    let received_indices: Uint32Array = Reflect::get(&received,&JsValue::from_str("indices")).unwrap().dyn_into().unwrap();
    assert_eq!(received_indices.length(),0);
    // Accessing the old detached zero-length buffer still throws, proving it
    // is distinct from a live empty output (byte_length alone cannot tell).
    let constructor: Function = Reflect::get(&global,&JsValue::from_str("Uint8Array")).unwrap().dyn_into().unwrap();
    let arguments = Array::new(); arguments.push(&indices.buffer());
    assert!(Reflect::construct(&constructor,&arguments).is_err());
}

#[wasm_bindgen_test]
fn canvas_binding_preserves_date_raw_captions_and_packed_partial_states_without_widening_items() {
    let mut options = params("canvas");
    options["rowCount"] = serde_json::json!(4); options["dictionaryCount"] = serde_json::json!(2);
    options["textByteLength"] = serde_json::json!(26); options["count"] = serde_json::json!(2);
    let values = Float64Array::new_with_length(4);
    for (index,value) in [100.0,100.0,100.0,0.0].into_iter().enumerate() { values.set_index(index as u32,value); }
    let states = Uint8Array::new_with_length(4);
    for (index,value) in [2,2,1,3].into_iter().enumerate() { states.set_index(index as u32,value); }
    let text_ids = Uint32Array::new_with_length(8);
    for (index,value) in [u32::MAX,u32::MAX,u32::MAX,0,1,0,0,0].into_iter().enumerate() { text_ids.set_index(index as u32,value); }
    let availability = Uint8Array::new_with_length(4);
    for (index,value) in [2,1,2,0].into_iter().enumerate() { availability.set_index(index as u32,value); }
    let offsets = Uint32Array::new_with_length(3);
    for (index,value) in [0,6,26].into_iter().enumerate() { offsets.set_index(index as u32,value); }
    let bytes = Uint8Array::new_with_length(26);
    let mut position = 0;
    for text in ["100","1900-04-09"] { for unit in text.encode_utf16() {
        let encoded = unit.to_le_bytes(); bytes.set_index(position,encoded[0]); bytes.set_index(position+1,encoded[1]); position += 2;
    } }
    let ranks = Uint32Array::new_with_length(2); ranks.set_index(1,1);
    let mut task = query((values,states,text_ids,availability,offsets,bytes,ranks),
        to_control(&serde_json::json!({"context":context(),"params":options})).unwrap()).unwrap();
    while !task.step(1).unwrap() {}
    let mut result = task.finish().unwrap();
    assert_eq!(result.copy_indices().unwrap().to_vec(),vec![2,0]);
    assert_eq!(result.copy_has_data().unwrap().to_vec(),vec![4,3]);
    let metadata: serde_json::Value = from_control(result.metadata().unwrap()).unwrap();
    assert_eq!(metadata["total"],2); assert_eq!(metadata["dateCount"],2);
    assert_eq!(metadata["transfer"]["inputCopies"],7); assert_eq!(metadata["transfer"]["inputBytes"],4*18+2*8+4+26);
    assert_eq!(metadata["transfer"]["outputCopies"],2); assert_eq!(metadata["transfer"]["outputBytes"],10);
}
