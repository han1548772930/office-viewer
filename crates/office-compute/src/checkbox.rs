use crate::{ComputeError, MAX_ELEMENTS};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};

type Rect = [u32; 4];

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CheckboxCellsParams {
    pub row_count: u32,
    pub column_count: u32,
    pub range_count: u32,
    pub merge_count: u32,
    pub existing_count: u32,
    pub covered_cells: u32,
    pub max_output_runs: u32,
}
pub struct CheckboxCellsMemory { pub input_bytes: usize, pub output_bytes: usize, pub estimated_peak_bytes: usize }
pub struct CheckboxCellsResult { pub runs: Vec<u32>, pub changed_cells: u32 }

impl CheckboxCellsParams {
    pub fn validate_shape(&self, ranges: usize, merges: usize, positions: usize, actions: usize) -> Result<(), ComputeError> {
        if self.row_count == 0 || self.row_count > 1_048_576 || self.column_count == 0 || self.column_count > 16_384
            || self.range_count > 4_096 || self.merge_count as usize > MAX_ELEMENTS || self.existing_count as usize > MAX_ELEMENTS
            || self.covered_cells as usize > MAX_ELEMENTS || self.existing_count > self.covered_cells
            || self.max_output_runs > self.covered_cells || self.covered_cells > 0 && self.max_output_runs == 0
            || ranges != self.range_count as usize * 4 || merges != self.merge_count as usize * 4
            || positions != self.existing_count as usize * 2 || actions != self.existing_count as usize {
            return Err(ComputeError::invalid("Invalid checkbox dimensions, sparse positions or control shape"));
        }
        Ok(())
    }
    pub fn memory_estimate(&self) -> Result<CheckboxCellsMemory, ComputeError> {
        let input = (u64::from(self.range_count) + u64::from(self.merge_count)) * 16 + u64::from(self.existing_count) * 9;
        let output = u64::from(self.max_output_runs) * 20;
        let peak = input * 2 + u64::from(self.existing_count) * 96 + u64::from(self.max_output_runs) * 160
            + (u64::from(self.merge_count) + u64::from(self.range_count)) * 32 + 8 * 1024 * 1024;
        Ok(CheckboxCellsMemory { input_bytes: usize::try_from(input).map_err(|_| ComputeError::memory())?,
            output_bytes: usize::try_from(output).map_err(|_| ComputeError::memory())?,
            estimated_peak_bytes: usize::try_from(peak).map_err(|_| ComputeError::memory())? })
    }
}
fn area(r: Rect) -> u64 { u64::from(r[1] - r[0] + 1) * u64::from(r[3] - r[2] + 1) }
fn valid_rect(r: Rect, p: &CheckboxCellsParams) -> bool { r[0] <= r[1] && r[2] <= r[3] && r[1] < p.row_count && r[3] < p.column_count }
fn intersection(a: Rect, b: Rect) -> Option<Rect> {
    let r = [a[0].max(b[0]), a[1].min(b[1]), a[2].max(b[2]), a[3].min(b[3])];
    (r[0] <= r[1] && r[2] <= r[3]).then_some(r)
}

/** Pure compute kernel. Scalar/permission semantics stay in the shared typed host classifier.
 * Geometry preparation and scanning consume a bounded step budget; no output can be taken before all phases complete.
 */
pub struct CheckboxCells {
    params: CheckboxCellsParams,
    fragments: Vec<Rect>, scratch: Vec<Rect>, merges: Vec<u32>, positions: Vec<u32>, actions: Vec<u8>,
    sparse: BTreeMap<u64, u8>, sparse_cursor: usize, cut_cursor: usize, fragment_cursor: usize,
    scanning: bool, range_cursor: usize, row: u32, column: u32,
    run: Option<(u32, u32, u32, u8)>, runs: Vec<u32>, last_run: HashMap<(u32, u32, u8), usize>,
    processed: u64, changed: u32, done: bool, consumed: bool,
}
impl CheckboxCells {
    pub fn new(ranges: Vec<u32>, merges: Vec<u32>, positions: Vec<u32>, actions: Vec<u8>, params: CheckboxCellsParams,
        memory_budget: usize) -> Result<Self, ComputeError> {
        params.validate_shape(ranges.len(), merges.len(), positions.len(), actions.len())?;
        if params.memory_estimate()?.estimated_peak_bytes > memory_budget { return Err(ComputeError::memory()); }
        let mut fragments = Vec::new();
        fragments.try_reserve_exact(params.range_count as usize).map_err(|_| ComputeError::memory())?;
        let mut covered = 0;
        for chunk in ranges.chunks_exact(4) {
            let rect = [chunk[0], chunk[1], chunk[2], chunk[3]];
            if !valid_rect(rect, &params) || fragments.iter().any(|old| intersection(*old, rect).is_some()) {
                return Err(ComputeError::invalid("Checkbox ranges must be clipped, valid and disjoint"));
            }
            covered += area(rect); fragments.push(rect);
        }
        if covered != u64::from(params.covered_cells) { return Err(ComputeError::invalid("Checkbox covered area does not match its ranges")); }
        for chunk in merges.chunks_exact(4) {
            if !valid_rect([chunk[0], chunk[1], chunk[2], chunk[3]], &params) { return Err(ComputeError::invalid("Invalid checkbox merge")); }
        }
        let mut runs = Vec::new();
        runs.try_reserve_exact(params.max_output_runs as usize * 5).map_err(|_| ComputeError::memory())?;
        let mut last_run = HashMap::new();
        last_run.try_reserve(params.max_output_runs as usize).map_err(|_| ComputeError::memory())?;
        Ok(Self { params, fragments, scratch: Vec::new(), merges, positions, actions, sparse: BTreeMap::new(), sparse_cursor: 0,
            cut_cursor: 0, fragment_cursor: 0, scanning: false, range_cursor: 0, row: 0, column: 0, run: None, runs,
            last_run, processed: 0, changed: 0, done: false, consumed: false })
    }
    fn consumed_error() -> ComputeError { ComputeError::new("COMPUTE_RESULT_CONSUMED", "Checkbox plan was consumed") }
    fn push_fragment(&mut self, r: Rect) -> Result<(), ComputeError> {
        let cap = self.params.max_output_runs.max(self.params.range_count) as usize;
        if self.scratch.len() >= cap { return Err(ComputeError::memory()); }
        if self.scratch.len() == self.scratch.capacity() {
            self.scratch.try_reserve_exact((cap - self.scratch.len()).min(512)).map_err(|_| ComputeError::memory())?;
        }
        self.scratch.push(r); Ok(())
    }
    fn cut(&self) -> Option<Rect> {
        let i = self.cut_cursor / 2 * 4; let r = [self.merges[i], self.merges[i+1], self.merges[i+2], self.merges[i+3]];
        if self.cut_cursor % 2 == 0 { (r[2] < r[3]).then(|| [r[0], r[0], r[2] + 1, r[3]]) }
        else { (r[0] < r[1]).then(|| [r[0] + 1, r[1], r[2], r[3]]) }
    }
    fn cut_fragment(&mut self, source: Rect, cut: Rect) -> Result<(), ComputeError> {
        let Some(hit) = intersection(source, cut) else { return self.push_fragment(source); };
        self.processed += area(hit);
        if source[0] < hit[0] { self.push_fragment([source[0], hit[0]-1, source[2], source[3]])?; }
        if hit[1] < source[1] { self.push_fragment([hit[1]+1, source[1], source[2], source[3]])?; }
        if source[2] < hit[2] { self.push_fragment([hit[0], hit[1], source[2], hit[2]-1])?; }
        if hit[3] < source[3] { self.push_fragment([hit[0], hit[1], hit[3]+1, source[3]])?; }
        Ok(())
    }
    fn flush_run(&mut self) -> Result<(), ComputeError> {
        let Some((row, start, end, action)) = self.run.take() else { return Ok(()); };
        let key = (start, end, action);
        if let Some(&index) = self.last_run.get(&key) {
            if self.runs[index+1] + 1 == row { self.runs[index+1] = row; return Ok(()); }
        }
        if self.runs.len() / 5 >= self.params.max_output_runs as usize { return Err(ComputeError::memory()); }
        let index = self.runs.len(); self.runs.extend_from_slice(&[row, row, start, end, u32::from(action)]);
        self.last_run.insert(key, index); Ok(())
    }
    pub fn step(&mut self, budget: u32) -> Result<bool, ComputeError> {
        if self.consumed { return Err(Self::consumed_error()); }
        if budget == 0 { return Err(ComputeError::invalid("Checkbox step budget must be positive")); }
        for _ in 0..budget {
            if self.done { return Ok(true); }
            if self.sparse_cursor < self.actions.len() {
                let i = self.sparse_cursor; let row = self.positions[i*2]; let col = self.positions[i*2+1]; let action = self.actions[i];
                if row >= self.params.row_count || col >= self.params.column_count || action > 2
                    || !self.fragments.iter().any(|r| r[0] <= row && row <= r[1] && r[2] <= col && col <= r[3]) {
                    return Err(ComputeError::invalid("Sparse checkbox position or action is outside the rule"));
                }
                if self.sparse.insert(u64::from(row) * u64::from(self.params.column_count) + u64::from(col), action).is_some() {
                    return Err(ComputeError::invalid("Duplicate sparse checkbox position"));
                }
                self.sparse_cursor += 1; continue;
            }
            if self.cut_cursor < self.merges.len() / 2 {
                let Some(cut) = self.cut() else { self.cut_cursor += 1; continue; };
                if self.fragment_cursor < self.fragments.len() {
                    let source = self.fragments[self.fragment_cursor]; self.cut_fragment(source, cut)?; self.fragment_cursor += 1; continue;
                }
                std::mem::swap(&mut self.fragments, &mut self.scratch); self.scratch.clear(); self.fragment_cursor = 0; self.cut_cursor += 1; continue;
            }
            if !self.scanning {
                self.scanning = true;
                if let Some(r) = self.fragments.first() { self.row = r[0]; self.column = r[2]; }
            }
            if self.range_cursor >= self.fragments.len() { self.done = true; return Ok(true); }
            let r = self.fragments[self.range_cursor];
            let action = *self.sparse.get(&(u64::from(self.row) * u64::from(self.params.column_count) + u64::from(self.column))).unwrap_or(&1);
            if action == 0 { self.flush_run()?; }
            else {
                if let Some((row, _, end, previous)) = self.run {
                    if row == self.row && end + 1 == self.column && previous == action { self.run.as_mut().unwrap().2 = self.column; }
                    else { self.flush_run()?; self.run = Some((self.row, self.column, self.column, action)); }
                } else { self.run = Some((self.row, self.column, self.column, action)); }
                self.changed += 1;
            }
            self.processed += 1;
            if self.column < r[3] { self.column += 1; }
            else {
                self.flush_run()?;
                if self.row < r[1] { self.row += 1; self.column = r[2]; }
                else {
                    self.range_cursor += 1;
                    if let Some(next) = self.fragments.get(self.range_cursor) { self.row = next[0]; self.column = next[2]; }
                    else { self.done = true; }
                }
            }
        }
        Ok(self.done)
    }
    pub fn progress(&self) -> f64 { if self.params.covered_cells == 0 { if self.done { 1.0 } else { 0.0 } }
        else { self.processed as f64 / f64::from(self.params.covered_cells) } }
    pub fn take_result(&mut self) -> Result<CheckboxCellsResult, ComputeError> {
        if self.consumed { return Err(Self::consumed_error()); }
        if !self.done { return Err(ComputeError::new("COMPUTE_NOT_READY", "Checkbox plan is incomplete")); }
        self.consumed = true; Ok(CheckboxCellsResult { runs: std::mem::take(&mut self.runs), changed_cells: self.changed })
    }
}

#[cfg(test)]
mod tests;
