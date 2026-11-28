// Native algorithm drafts: do not run before unified validation.
use super::*;
fn params(rows: u32, columns: u32, ranges: &[u32], merges: &[u32], positions: &[u32], cap: u32) -> CheckboxCellsParams {
    CheckboxCellsParams { row_count: rows, column_count: columns, range_count: ranges.len() as u32 / 4,
        merge_count: merges.len() as u32 / 4, existing_count: positions.len() as u32 / 2,
        covered_cells: ranges.chunks_exact(4).map(|r| (r[1]-r[0]+1)*(r[3]-r[2]+1)).sum(), max_output_runs: cap }
}
fn finish(task: &mut CheckboxCells, budget: u32) -> CheckboxCellsResult {
    let mut previous = 0.0;
    loop { let done = task.step(budget).unwrap(); let progress = task.progress(); assert!(progress >= previous && progress <= 1.0);
        previous = progress; if done { break; } }
    assert_eq!(task.progress(),1.0); task.take_result().unwrap()
}
fn expanded(result: &CheckboxCellsResult) -> BTreeMap<(u32,u32),u8> {
    let mut cells = BTreeMap::new();
    for r in result.runs.chunks_exact(5) { for row in r[0]..=r[1] { for column in r[2]..=r[3] {
        assert!(cells.insert((row,column),r[4] as u8).is_none(),"overlapping output");
    } } }
    assert_eq!(cells.len(),result.changed_cells as usize); cells
}
#[test]
fn empty_band_compresses_to_one_rectangle_and_no_partial_plan_is_observable() {
    let r = vec![0,63,0,63]; let p = params(64,64,&r,&[],&[],64);
    let mut task = CheckboxCells::new(r,vec![],vec![],vec![],p,64*1024*1024).unwrap();
    assert_eq!(task.take_result().err().unwrap().code,"COMPUTE_NOT_READY"); assert!(!task.step(1).unwrap());
    let result = finish(&mut task,17); assert_eq!(result.runs,vec![0,63,0,63,1]); assert_eq!(result.changed_cells,4096);
    assert_eq!(task.take_result().err().unwrap().code,"COMPUTE_RESULT_CONSUMED");
}
#[test]
fn sparse_actions_preserve_invalid_contents_and_merge_members_with_only_one_master_write() {
    let r = vec![0,2,0,4]; let m = vec![0,1,1,2]; let positions = vec![0,1,0,3,2,4];
    let p = params(3,5,&r,&m,&positions,15); let mut task = CheckboxCells::new(r,m,positions,vec![2,0,0],p,64*1024*1024).unwrap();
    let result = finish(&mut task,1); let cells = expanded(&result);
    assert_eq!(cells.get(&(0,1)),Some(&2)); assert_eq!(cells.get(&(0,0)),Some(&1)); assert_eq!(cells.len(),10);
    for position in [(0,2),(1,1),(1,2),(0,3),(2,4)] { assert!(!cells.contains_key(&position)); }
}
#[test]
fn selecting_only_merge_members_writes_no_master_outside_the_rule() {
    let r = vec![1,1,1,2]; let m = vec![0,1,1,2]; let p = params(3,3,&r,&m,&[],2);
    let mut task = CheckboxCells::new(r,m,vec![],vec![],p,64*1024*1024).unwrap();
    let result = finish(&mut task,1); assert!(result.runs.is_empty()); assert_eq!(result.changed_cells,0);
}
#[test]
fn rectangular_union_and_merge_clipping_match_an_independent_small_cell_oracle_for_every_action_pattern() {
    let r = vec![0,1,0,2,2,3,1,3]; let m = vec![1,2,1,2];
    for pattern in 0..81u32 {
        let positions = vec![0,0,1,1,2,3,3,2]; let mut value = pattern;
        let actions: Vec<u8> = (0..4).map(|_| { let action = (value%3) as u8; value/=3; action }).collect();
        let p = params(4,4,&r,&m,&positions,12);
        let mut expected = BTreeMap::new();
        for row in 0..4 { for column in 0..4 {
            if !(row<=1 && column<=2 || row>=2 && column>=1) || [(1,2),(2,1),(2,2)].contains(&(row,column)) { continue; }
            let index = positions.chunks_exact(2).position(|p| p[0]==row && p[1]==column);
            let action = index.map(|i|actions[i]).unwrap_or(1); if action!=0 { expected.insert((row,column),action); }
        } }
        let mut task = CheckboxCells::new(r.clone(),m.clone(),positions,actions,p,64*1024*1024).unwrap();
        assert_eq!(expanded(&finish(&mut task,pattern%11+1)),expected);
    }
}
#[test]
fn overlap_bad_coordinates_area_and_unknown_actions_are_rejected_before_success() {
    let r = vec![0,1,0,1,1,2,1,2]; let p = params(3,3,&r,&[],&[],8);
    assert_eq!(CheckboxCells::new(r,vec![],vec![],vec![],p,64*1024*1024).err().unwrap().code,"COMPUTE_INVALID_INPUT");
    for (positions,actions) in [(vec![0,0,0,0],vec![1,2]),(vec![9,0],vec![0]),(vec![0,0],vec![3])] {
        let r = vec![0,1,0,1]; let p = params(2,2,&r,&[],&positions,4);
        let mut task = CheckboxCells::new(r,vec![],positions,actions,p,64*1024*1024).unwrap();
        assert_eq!(task.step(8).err().unwrap().code,"COMPUTE_INVALID_INPUT");
        assert_eq!(task.take_result().err().unwrap().code,"COMPUTE_NOT_READY");
    }
}
#[test]
fn input_and_fragmented_output_budgets_reject_instead_of_returning_partial_cells() {
    let r = vec![0,1,0,1]; let p = params(2,2,&r,&[],&[],4);
    assert_eq!(CheckboxCells::new(r.clone(),vec![],vec![],vec![],p,1).err().unwrap().code,"COMPUTE_MEMORY_LIMIT");
    let positions = vec![0,0,1,1]; let p = params(2,2,&r,&[],&positions,1);
    let mut task = CheckboxCells::new(r,vec![],positions,vec![2,2],p,64*1024*1024).unwrap();
    assert_eq!(task.step(128).err().unwrap().code,"COMPUTE_MEMORY_LIMIT");
    assert_eq!(task.take_result().err().unwrap().code,"COMPUTE_NOT_READY");
}
#[test]
fn cancelled_partial_owners_can_be_dropped_and_replaced_without_retained_output() {
    let r = vec![0,63,0,63]; let p = params(64,64,&r,&[],&[],64);
    let mut partial = CheckboxCells::new(r.clone(),vec![],vec![],vec![],p.clone(),64*1024*1024).unwrap();
    assert!(!partial.step(16).unwrap()); drop(partial);
    let mut replacement = CheckboxCells::new(r,vec![],vec![],vec![],p,64*1024*1024).unwrap();
    assert_eq!(finish(&mut replacement,128).changed_cells,4096);
}
#[test]
fn typed_shape_rejects_unknown_fields_and_empty_inputs_have_independent_complete_results() {
    let p = params(2,2,&[],&[],&[],0); let mut value = serde_json::to_value(&p).unwrap(); value["unknown"] = serde_json::json!(1);
    assert!(serde_json::from_value::<CheckboxCellsParams>(value).is_err());
    let mut task = CheckboxCells::new(vec![],vec![],vec![],vec![],p,64*1024*1024).unwrap();
    assert_eq!(task.step(0).err().unwrap().code,"COMPUTE_INVALID_INPUT");
    assert!(finish(&mut task,1).runs.is_empty());
}
