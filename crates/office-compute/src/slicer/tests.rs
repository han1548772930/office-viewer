use super::*;

// Native contract drafts. Fixed expected rows exercise the Rust authority;
// these tests have not been run in the current implementation-only phase.
#[derive(Clone)]
struct Fixture {
    values: Vec<f64>, states: Vec<u8>, text_ids: Vec<u32>, availability: Vec<u8>,
    offsets: Vec<u32>, bytes: Vec<u8>, ranks: Vec<u32>, params: SlicerMembersParams,
}

fn dictionary(texts: &[&str]) -> (Vec<u32>, Vec<u8>) {
    let mut offsets = vec![0];
    let mut bytes = Vec::new();
    for text in texts {
        for unit in text.encode_utf16() { bytes.extend_from_slice(&unit.to_le_bytes()); }
        offsets.push(bytes.len() as u32);
    }
    (offsets, bytes)
}

fn fixture() -> Fixture {
    let (offsets, bytes) = dictionary(&["1", "true", "#N/A", ""]);
    Fixture {
        values: vec![0.0,1.0,1.0,0.0,1.0,0.0,0.0,0.0,0.0,1.0,-0.0,0.0,5.5,-2.25],
        states: vec![SLICER_BLANK,SLICER_NUMBER,SLICER_DATE,SLICER_STRING,SLICER_BOOLEAN,SLICER_STRING,
            SLICER_ERROR,SLICER_STRING,SLICER_STRING,SLICER_NUMBER,SLICER_NUMBER,SLICER_NUMBER,SLICER_DATE,SLICER_DATE],
        text_ids: vec![SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,0,SLICER_NO_TEXT,1,2,2,3,
            SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT],
        availability: vec![0,0,0,1,1,0,0,1,0,1,0,1,1,1],
        offsets, ranks: vec![2,3,1,0],
        params: SlicerMembersParams { mode: SlicerQueryMode::Items, row_count: 14, dictionary_count: 4,
            text_byte_length: bytes.len() as u32, start: 0, count: 100, locale: "en-US".into(),
            sort_order: SlicerSortOrder::Ascending, custom_list_sort: false,
            cross_filter: SlicerCrossFilter::ShowItemsWithNoData, hide_items_with_no_data: false },
        bytes,
    }
}

fn task(f: Fixture) -> Result<SlicerMembers, ComputeError> {
    let maximum = f.params.memory_estimate()?.estimated_peak_bytes;
    SlicerMembers::new(f.values,f.states,f.text_ids,f.availability,f.offsets,f.bytes,f.ranks,f.params,maximum)
}
fn finish(f: Fixture, budget: u32) -> SlicerMembersResult {
    let mut query = task(f).unwrap();
    while !query.step(budget).unwrap() {}
    query.take_result().unwrap()
}
const ASCENDING: &[u32] = &[0,10,1,13,2,12,8,7,3,5,4,6];
const AVAILABLE: &[u32] = &[10,1,13,12,7,3,4];

#[test]
fn raw_types_empty_strings_error_text_and_signed_zero_remain_distinct() {
    let result = finish(fixture(), 1);
    assert_eq!(result.total, 12);
    assert_eq!(result.indices, ASCENDING);
    assert_eq!(result.has_data, vec![0,1,1,1,0,1,0,1,1,0,1,0]);
    assert_eq!(result.indices.iter().filter(|row| **row == 10 || **row == 11).count(), 1);
}

#[test]
fn duplicate_availability_is_or_and_keeps_the_first_source_row() {
    let result = finish(fixture(), 17);
    let position = result.indices.iter().position(|row| *row == 1).unwrap();
    assert_eq!(result.has_data[position], 1);
    assert!(!result.indices.contains(&9));
    let zero = result.indices.iter().position(|row| *row == 10).unwrap();
    assert_eq!(result.has_data[zero], 1);
    assert!(!result.indices.contains(&11));
}

#[test]
fn ascending_and_descending_order_use_numeric_date_and_supplied_locale_ranks() {
    let mut f = fixture();
    f.params.sort_order = SlicerSortOrder::Descending;
    assert_eq!(finish(f, 3).indices, vec![6,4,5,3,7,8,12,2,13,1,10,0]);
}

#[test]
fn equal_locale_or_calendar_ranks_keep_raw_identities_and_source_tie_order() {
    let (offsets, bytes) = dictionary(&["January", "Jan", "ALPHA", "alpha"]);
    let mut f = fixture();
    f.values = vec![0.0;4]; f.states = vec![SLICER_STRING;4]; f.text_ids = vec![1,0,3,2];
    f.availability = vec![1;4]; f.offsets = offsets; f.ranks = vec![0,0,1,1];
    f.params.row_count = 4; f.params.text_byte_length = bytes.len() as u32; f.bytes = bytes;
    f.params.custom_list_sort = true;
    assert_eq!(finish(f.clone(), 1).indices, vec![0,1,2,3]);
    f.params.sort_order = SlicerSortOrder::Descending;
    assert_eq!(finish(f, 1).indices, vec![2,3,0,1]);
}

#[test]
fn none_ignores_hide_while_hide_precedes_stable_data_at_top() {
    for hide in [false,true] {
        let mut f = fixture(); f.params.cross_filter = SlicerCrossFilter::None;
        f.params.hide_items_with_no_data = hide;
        let result = finish(f, 2);
        assert_eq!(result.indices, ASCENDING); assert_eq!(result.has_data, vec![1;12]);
    }
    for cross in [SlicerCrossFilter::ShowItemsWithDataAtTop,SlicerCrossFilter::ShowItemsWithNoData] {
        let mut f = fixture(); f.params.cross_filter = cross; f.params.hide_items_with_no_data = true;
        let result = finish(f, 2);
        assert_eq!(result.total, 7); assert_eq!(result.indices, AVAILABLE); assert_eq!(result.has_data, vec![1;7]);
    }
}

#[test]
fn data_at_top_is_a_stable_partition_before_bounded_window_slicing() {
    let mut f = fixture(); f.params.cross_filter = SlicerCrossFilter::ShowItemsWithDataAtTop;
    assert_eq!(finish(f.clone(), 1).indices, vec![10,1,13,12,7,3,4,0,2,8,5,6]);
    f.params.start = 5; f.params.count = 4;
    let result = finish(f, 1);
    assert_eq!(result.total, 12); assert_eq!(result.indices, vec![3,4,0,2]);
    assert_eq!(result.has_data, vec![1,1,0,0]);
}

#[test]
fn real_totals_tail_pages_and_safe_out_of_range_start_do_not_depend_on_window() {
    let mut f = fixture(); f.params.start = 10; f.params.count = 200;
    let result = finish(f.clone(), 3);
    assert_eq!(result.total, 12); assert_eq!(result.indices, vec![4,6]);
    f.params.start = JS_SAFE_INTEGER;
    let result = finish(f, 3);
    assert_eq!(result.total, 12); assert!(result.indices.is_empty()); assert!(result.has_data.is_empty());
}

#[test]
fn bounds_use_only_typed_dates_preserve_fractional_serials_and_ignore_availability() {
    let mut f = fixture(); f.values[1] = -999.0; f.values[9] = 9999.0;
    f.params.hide_items_with_no_data = true;
    let result = finish(f, 1);
    assert_eq!(result.date_count, 3); assert_eq!(result.min_date, Some(-2.25)); assert_eq!(result.max_date, Some(5.5));
}

#[test]
fn metadata_skips_member_index_and_returns_empty_outputs_with_real_date_bounds() {
    let mut f = fixture(); f.params.mode = SlicerQueryMode::Metadata;
    for row in 0..f.values.len() { if f.states[row] != SLICER_DATE { f.states[row] = SLICER_BLANK; f.values[row] = 0.0; } }
    f.text_ids.fill(SLICER_NO_TEXT); f.offsets = vec![0]; f.bytes.clear(); f.ranks.clear();
    f.params.dictionary_count = 0; f.params.text_byte_length = 0;
    let result = finish(f, 2);
    assert_eq!(result.total, 0); assert!(result.indices.is_empty()); assert!(result.has_data.is_empty());
    assert_eq!(result.date_count, 3); assert_eq!(result.min_date, Some(-2.25)); assert_eq!(result.max_date, Some(5.5));
}

#[test]
fn empty_sources_and_no_date_sources_have_null_bounds() {
    let mut f = fixture(); f.values.clear(); f.states.clear(); f.text_ids.clear(); f.availability.clear();
    f.offsets = vec![0]; f.bytes.clear(); f.ranks.clear();
    f.params.row_count = 0; f.params.dictionary_count = 0; f.params.text_byte_length = 0;
    let result = finish(f, 1);
    assert_eq!(result.total, 0); assert_eq!(result.date_count, 0);
    assert_eq!(result.min_date, None); assert_eq!(result.max_date, None);
}

#[test]
fn exact_utf16_dictionary_preserves_lone_surrogate_identities() {
    let mut f = fixture(); f.values = vec![0.0;2]; f.states = vec![SLICER_STRING;2];
    f.text_ids = vec![0,1]; f.availability = vec![1;2]; f.offsets = vec![0,2,4];
    f.bytes = vec![0x00,0xd8,0x01,0xd8]; f.ranks = vec![0,0];
    f.params.row_count = 2; f.params.dictionary_count = 2; f.params.text_byte_length = 4;
    let result = finish(f, 1);
    assert_eq!(result.total, 2); assert_eq!(result.indices, vec![0,1]);
}

#[test]
fn malformed_dictionary_offsets_duplicates_and_rank_bounds_reject_before_steps() {
    for change in 0..4 {
        let mut f = fixture();
        match change {
            0 => f.offsets[0] = 2,
            1 => f.offsets[1] = 1,
            2 => f.ranks[0] = 4,
            _ => { let (offsets,bytes) = dictionary(&["1","1","#N/A",""]); f.offsets = offsets;
                f.params.text_byte_length = bytes.len() as u32; f.bytes = bytes; },
        }
        assert_eq!(task(f).err().unwrap().code, "COMPUTE_INVALID_INPUT");
    }
}

#[test]
fn late_invalid_rows_never_publish_partial_members_or_date_metadata() {
    for change in 0..6 {
        let mut f = fixture();
        match change {
            0 => f.values[13] = f64::NAN,
            1 => f.values[13] = f64::INFINITY,
            2 => f.states[13] = 6,
            3 => f.availability[13] = 2,
            4 => { f.states[13] = SLICER_BOOLEAN; f.values[13] = 2.0; },
            _ => f.text_ids[13] = 0,
        }
        let mut query = task(f).unwrap();
        let error = (0..100).find_map(|_| query.step(1).err()).expect("invalid row must fail within bounded steps");
        assert_eq!(error.code, "COMPUTE_INVALID_INPUT");
        assert_eq!(query.take_result().unwrap_err().code, "COMPUTE_RESULT_CONSUMED");
        assert_eq!(query.step(1).unwrap_err().code, "COMPUTE_RESULT_CONSUMED");
    }
}

#[test]
fn step_budgets_preserve_results_and_progress_and_consumption_are_explicit() {
    for budget in [0,1,2,17,1_000_000,u32::MAX] {
        let mut query = task(fixture()).unwrap();
        assert_eq!(query.take_result().unwrap_err().code, "COMPUTE_NOT_READY");
        let mut previous = 0.0;
        loop {
            let done = query.step(budget).unwrap(); let progress = query.progress();
            assert!((0.0..=1.0).contains(&progress)); assert!(progress >= previous); previous = progress;
            if done { break; }
        }
        assert_eq!(query.progress(), 1.0); assert_eq!(query.take_result().unwrap().indices, ASCENDING);
        assert_eq!(query.take_result().unwrap_err().code, "COMPUTE_RESULT_CONSUMED");
    }
}

#[test]
fn shape_control_and_conservative_admission_budget_have_exact_boundaries() {
    let f = fixture(); let memory = f.params.memory_estimate().unwrap();
    assert_eq!(memory.input_bytes, 14*14 + 4*8 + 4 + f.bytes.len());
    assert_eq!(memory.output_bytes, 14*5);
    let error = SlicerMembers::new(f.values,f.states,f.text_ids,f.availability,f.offsets,f.bytes,f.ranks,
        f.params.clone(),memory.estimated_peak_bytes-1).err().unwrap();
    assert_eq!(error.code, "COMPUTE_MEMORY_LIMIT");
    let mut params = f.params; params.start = JS_SAFE_INTEGER + 1;
    assert_eq!(params.validate().unwrap_err().code, "COMPUTE_INVALID_INPUT");
    params.start = 0; params.count = 0;
    assert!(params.validate().is_err());
    let mut value = serde_json::to_value(fixture().params).unwrap(); value["unknown"] = serde_json::json!(true);
    assert!(serde_json::from_value::<SlicerMembersParams>(value).is_err());
}

#[test]
fn boolean_order_and_identity_do_not_collide_with_numeric_or_text_payloads() {
    let (offsets,bytes) = dictionary(&["true","false"]);
    let mut f = fixture();
    f.values = vec![1.0,0.0,1.0,0.0,0.0,0.0,1.0];
    f.states = vec![SLICER_BOOLEAN,SLICER_BOOLEAN,SLICER_NUMBER,SLICER_NUMBER,SLICER_STRING,SLICER_STRING,SLICER_BOOLEAN];
    f.text_ids = vec![SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,0,1,SLICER_NO_TEXT];
    f.availability = vec![0,1,1,1,1,1,1]; f.offsets = offsets; f.ranks = vec![1,0];
    f.params.row_count = 7; f.params.dictionary_count = 2; f.params.text_byte_length = bytes.len() as u32; f.bytes = bytes;
    let ascending = finish(f.clone(),1);
    assert_eq!(ascending.total,6); assert_eq!(ascending.indices,vec![3,2,5,4,1,0]); assert_eq!(ascending.has_data,vec![1;6]);
    f.params.sort_order = SlicerSortOrder::Descending;
    assert_eq!(finish(f,1).indices,vec![0,1,4,5,2,3]);
}

fn canvas_fixture() -> Fixture {
    let mut f = fixture();
    let (offsets,bytes) = dictionary(&["1","true","#N/A","","0","5.5","-2.25"]);
    f.params.mode = SlicerQueryMode::Canvas; f.params.dictionary_count = 7;
    f.params.text_byte_length = bytes.len() as u32; f.offsets = offsets; f.bytes = bytes;
    f.ranks = vec![4,6,1,0,3,5,2];
    f.text_ids.extend([3,0,0,0,1,1,2,2,3,0,4,4,5,6]);
    for row in [1,4,9,10,12,13] { f.availability[row] |= 2; }
    f
}

#[test]
fn canvas_globally_folds_captions_after_typed_sort_and_returns_caption_total_and_tri_state() {
    let result = finish(canvas_fixture(),1);
    assert_eq!(result.total,7); assert_eq!(result.indices,vec![0,10,1,13,12,7,5]);
    assert_eq!(result.has_data,vec![0,5,5,3,3,5,5]);
    assert_eq!(result.date_count,3); assert_eq!(result.min_date,Some(-2.25)); assert_eq!(result.max_date,Some(5.5));
    let mut f = canvas_fixture(); f.params.start = 1; f.params.count = 3;
    let window = finish(f,1);
    assert_eq!(window.total,7); assert_eq!(window.indices,vec![10,1,13]); assert_eq!(window.has_data,vec![5,5,3]);
    let mut descending = canvas_fixture(); descending.params.sort_order = SlicerSortOrder::Descending;
    assert_eq!(finish(descending,1).indices,vec![6,4,3,8,12,13,10]);
}

#[test]
fn canvas_hide_and_data_at_top_fold_globally_but_hidden_rows_still_contribute_partial_selection() {
    let mut f = canvas_fixture(); f.params.cross_filter = SlicerCrossFilter::ShowItemsWithDataAtTop;
    let result = finish(f.clone(),1);
    assert_eq!(result.indices,vec![10,1,13,12,7,4,0]); assert_eq!(result.has_data,vec![5,5,3,3,5,5,0]);
    f.params.hide_items_with_no_data = true;
    let result = finish(f,1);
    assert_eq!(result.total,6); assert_eq!(result.indices,vec![10,1,13,12,7,4]);
    assert_eq!(result.has_data,vec![5,5,3,3,5,5]);
    let mut f = canvas_fixture(); f.params.cross_filter = SlicerCrossFilter::None; f.params.hide_items_with_no_data = true;
    assert_eq!(finish(f,1).has_data,vec![1,5,5,3,3,5,5]);
}

#[test]
fn canvas_distinct_date_captions_do_not_corrupt_raw_identity_or_all_source_caption_selection() {
    let mut f = fixture(); let (offsets,bytes) = dictionary(&["100","1900-04-09"]);
    f.values = vec![100.0,100.0,100.0,0.0]; f.states = vec![SLICER_DATE,SLICER_DATE,SLICER_NUMBER,SLICER_STRING];
    f.text_ids = vec![SLICER_NO_TEXT,SLICER_NO_TEXT,SLICER_NO_TEXT,0,1,0,0,0];
    f.availability = vec![2,1,2,0]; f.offsets = offsets; f.ranks = vec![0,1];
    f.params.mode = SlicerQueryMode::Canvas; f.params.row_count = 4; f.params.dictionary_count = 2;
    f.params.text_byte_length = bytes.len() as u32; f.bytes = bytes;
    let result = finish(f.clone(),1);
    assert_eq!(result.total,2); assert_eq!(result.indices,vec![2,0]); assert_eq!(result.has_data,vec![4,3]);
    f.params.hide_items_with_no_data = true;
    let result = finish(f,1);
    assert_eq!(result.total,1); assert_eq!(result.indices,vec![0]); assert_eq!(result.has_data,vec![3]);
}

#[test]
fn canvas_dictionary_may_include_raw_error_and_different_source_caption_without_identity_loss() {
    let mut f = fixture(); let (offsets,bytes) = dictionary(&["#N/A","0"]);
    f.values = vec![0.0]; f.states = vec![SLICER_ERROR]; f.text_ids = vec![0,1]; f.availability = vec![2];
    f.offsets = offsets; f.ranks = vec![0,1]; f.params.mode = SlicerQueryMode::Canvas;
    f.params.row_count = 1; f.params.dictionary_count = 2; f.params.text_byte_length = bytes.len() as u32; f.bytes = bytes;
    let result = finish(f.clone(),1); assert_eq!(result.total,1); assert_eq!(result.indices,vec![0]); assert_eq!(result.has_data,vec![2]);
    f.text_ids[1] = 2;
    let mut query = task(f).unwrap(); assert_eq!(query.step(1).unwrap_err().code,"COMPUTE_INVALID_INPUT");
    assert_eq!(query.take_result().unwrap_err().code,"COMPUTE_RESULT_CONSUMED");
}

#[test]
fn canvas_memory_step_progress_and_consumption_keep_the_same_bounded_ownership_contract() {
    let f = canvas_fixture(); let memory = f.params.memory_estimate().unwrap();
    assert_eq!(memory.input_bytes,14*18+7*8+4+f.bytes.len());
    for budget in [0,1,3,17,1_000_000] {
        let mut query = task(f.clone()).unwrap(); let mut previous = 0.0;
        loop { let done = query.step(budget).unwrap(); let progress = query.progress();
            assert!(progress >= previous && progress <= 1.0); previous = progress; if done { break; } }
        assert_eq!(query.take_result().unwrap().total,7);
        assert_eq!(query.step(1).unwrap_err().code,"COMPUTE_RESULT_CONSUMED");
    }
}
