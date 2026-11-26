use crate::{ComputeError, MAX_ROWS};
use office_wasm_bridge::JS_SAFE_INTEGER;
use serde::{Deserialize, Serialize};
use std::{cmp::Ordering, collections::HashMap};

pub const SLICER_NO_TEXT: u32 = u32::MAX;
pub const SLICER_BLANK: u8 = 0;
pub const SLICER_NUMBER: u8 = 1;
pub const SLICER_DATE: u8 = 2;
pub const SLICER_STRING: u8 = 3;
pub const SLICER_BOOLEAN: u8 = 4;
pub const SLICER_ERROR: u8 = 5;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SlicerQueryMode { Items, Canvas, Metadata }
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SlicerSortOrder { Ascending, Descending }
#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum SlicerCrossFilter { None, ShowItemsWithDataAtTop, ShowItemsWithNoData }

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SlicerMembersParams {
    pub mode: SlicerQueryMode,
    pub row_count: u32,
    pub dictionary_count: u32,
    pub text_byte_length: u32,
    pub start: u64,
    pub count: u32,
    /// Intl locale/custom calendar ranks are prepared in the Worker. Rust
    /// never substitutes an ordinal string comparison for those ranks.
    pub locale: String,
    pub sort_order: SlicerSortOrder,
    pub custom_list_sort: bool,
    pub cross_filter: SlicerCrossFilter,
    pub hide_items_with_no_data: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlicerMembersMemory {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub estimated_peak_bytes: usize,
}

impl SlicerMembersParams {
    pub fn validate(&self) -> Result<(), ComputeError> {
        let dictionary_maximum = self.row_count.saturating_mul(if self.mode == SlicerQueryMode::Canvas { 2 } else { 1 });
        if self.row_count as usize > MAX_ROWS || self.dictionary_count > dictionary_maximum
            || self.count == 0 || self.count as usize > MAX_ROWS || self.start > JS_SAFE_INTEGER
            || self.text_byte_length % 2 != 0 || self.locale.trim().is_empty() || self.locale.len() > 128 {
            return Err(ComputeError::invalid("Invalid typed Slicer query parameters"));
        }
        if self.mode == SlicerQueryMode::Metadata && (self.dictionary_count != 0 || self.text_byte_length != 0) {
            return Err(ComputeError::invalid("Date metadata must not carry a text dictionary"));
        }
        Ok(())
    }

    /// Admission estimate shared with TypeScript. Model/source heap memory is
    /// separate; this estimate includes retained input, dictionary/Intl prep,
    /// hash entries, merge scratch and independent window outputs.
    pub fn memory_estimate(&self) -> Result<SlicerMembersMemory, ComputeError> {
        self.validate()?;
        let rows = self.row_count as usize;
        let dictionary = self.dictionary_count as usize;
        let text = self.text_byte_length as usize;
        let input_bytes = rows.checked_mul(if self.mode == SlicerQueryMode::Canvas { 18 } else { 14 })
            .and_then(|n| n.checked_add(dictionary.checked_mul(8)?.checked_add(4)?))
            .and_then(|n| n.checked_add(text)).ok_or_else(ComputeError::memory)?;
        let output_bytes = if self.mode != SlicerQueryMode::Metadata {
            rows.min(self.count as usize).checked_mul(5).ok_or_else(ComputeError::memory)?
        } else { 0 };
        let estimated_peak_bytes = input_bytes.checked_mul(2)
            .and_then(|n| n.checked_add(rows.checked_mul(128)?))
            .and_then(|n| n.checked_add(dictionary.checked_mul(160)?))
            .and_then(|n| n.checked_add(text.checked_mul(2)?))
            .and_then(|n| n.checked_add(output_bytes.checked_mul(2)?))
            .and_then(|n| n.checked_add(8 * 1024 * 1024)).ok_or_else(ComputeError::memory)?;
        Ok(SlicerMembersMemory { input_bytes, output_bytes, estimated_peak_bytes })
    }

    pub fn validate_shape(&self, values: usize, states: usize, text_ids: usize, availability: usize,
        offsets: usize, text: usize, ranks: usize) -> Result<(), ComputeError> {
        self.validate()?;
        let rows = self.row_count as usize;
        let dictionary = self.dictionary_count as usize;
        let text_rows = rows * (if self.mode == SlicerQueryMode::Canvas { 2 } else { 1 });
        if values != rows || states != rows || text_ids != text_rows || availability != rows
            || offsets != dictionary + 1 || text != self.text_byte_length as usize || ranks != dictionary {
            return Err(ComputeError::invalid("Typed Slicer columns/dictionary dimensions differ"));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct MemberKey { kind: u8, payload: u64 }
#[derive(Debug, Clone, Copy)]
struct Member { first_row: u32, has_data: bool }
#[derive(Default)]
struct CanvasCaption { first_row: Option<u32>, has_data: bool, selected: bool, unselected: bool }
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase { Scan, Sort, Project, CanvasOutput, Done }

#[derive(Debug)]
pub struct SlicerMembersResult {
    pub indices: Vec<u32>,
    pub has_data: Vec<u8>,
    pub total: u32,
    pub date_count: u32,
    pub min_date: Option<f64>,
    pub max_date: Option<f64>,
}

/// One authority for typed member identity, duplicate availability OR,
/// stable ordering, hide/data-at-top projection and exact Date bounds.
/// Input Vecs move into this owner; no retained input or result Vec clone.
pub struct SlicerMembers {
    values: Vec<f64>,
    states: Vec<u8>,
    text_ids: Vec<u32>,
    availability: Vec<u8>,
    text_ranks: Vec<u32>,
    params: SlicerMembersParams,
    members: Vec<Member>,
    lookup: HashMap<MemberKey, u32>,
    order: Vec<u32>,
    scratch: Vec<u32>,
    captions: Vec<CanvasCaption>,
    caption_order: Vec<u32>,
    canvas_cursor: usize,
    row: usize,
    phase: Phase,
    width: usize,
    left: usize,
    middle: usize,
    right: usize,
    i: usize,
    j: usize,
    out: usize,
    sort_work: u64,
    sort_total: u64,
    project_cursor: usize,
    project_pass: u8,
    result: Option<SlicerMembersResult>,
    failed: bool,
}

fn reserve<T>(length: usize) -> Result<Vec<T>, ComputeError> {
    let mut result = Vec::new();
    result.try_reserve_exact(length).map_err(|_| ComputeError::memory())?;
    Ok(result)
}

impl SlicerMembers {
    #[allow(clippy::too_many_arguments)]
    pub fn new(values: Vec<f64>, states: Vec<u8>, text_ids: Vec<u32>, availability: Vec<u8>,
        text_offsets: Vec<u32>, text_bytes: Vec<u8>, text_ranks: Vec<u32>, params: SlicerMembersParams,
        maximum_working_bytes: usize) -> Result<Self, ComputeError> {
        params.validate_shape(values.len(), states.len(), text_ids.len(), availability.len(),
            text_offsets.len(), text_bytes.len(), text_ranks.len())?;
        if params.memory_estimate()?.estimated_peak_bytes > maximum_working_bytes {
            return Err(ComputeError::memory());
        }
        // Dictionary interning is an explicit wire invariant. Verify exact
        // UTF-16 bytes, not locale ranks, before accepting any row identities.
        // Borrowed keys live only during this validation; text is never cloned.
        if text_offsets.first() != Some(&0) || text_offsets.last().copied() != Some(params.text_byte_length) {
            return Err(ComputeError::invalid("Invalid UTF-16 dictionary bounds"));
        }
        let mut dictionary = HashMap::new();
        dictionary.try_reserve(params.dictionary_count as usize).map_err(|_| ComputeError::memory())?;
        for (index, pair) in text_offsets.windows(2).enumerate() {
            if pair[0] > pair[1] || pair[1] as usize > text_bytes.len() || pair[0] % 2 != 0 || pair[1] % 2 != 0
                || text_ranks[index] >= params.dictionary_count {
                return Err(ComputeError::invalid("Invalid UTF-16 dictionary offset/rank"));
            }
            let text = &text_bytes[pair[0] as usize..pair[1] as usize];
            if dictionary.insert(text, ()).is_some() {
                return Err(ComputeError::invalid("Text dictionary entries must be exactly interned"));
            }
        }
        drop(dictionary);
        // Offsets/text have served identity validation and no longer need to
        // remain retained during source deduplication and numeric/rank sorting.
        drop(text_offsets);
        drop(text_bytes);
        let rows = if params.mode != SlicerQueryMode::Metadata { params.row_count as usize } else { 0 };
        let mut lookup = HashMap::new();
        lookup.try_reserve(rows).map_err(|_| ComputeError::memory())?;
        let members = reserve(rows)?;
        let order = reserve(rows)?;
        let mut scratch = reserve(rows)?;
        scratch.resize(rows, 0);
        let mut captions = reserve(if params.mode == SlicerQueryMode::Canvas { params.dictionary_count as usize } else { 0 })?;
        if params.mode == SlicerQueryMode::Canvas { captions.resize_with(params.dictionary_count as usize, CanvasCaption::default); }
        let caption_order = reserve(if params.mode == SlicerQueryMode::Canvas { rows } else { 0 })?;
        let output_count = rows.min(params.count as usize);
        let result = SlicerMembersResult { indices: reserve(output_count)?, has_data: reserve(output_count)?,
            total: 0, date_count: 0, min_date: None, max_date: None };
        Ok(Self { values, states, text_ids, availability, text_ranks, params, members, lookup, order, scratch,
            captions, caption_order, canvas_cursor: 0,
            row: 0, phase: Phase::Scan, width: 1, left: 0, middle: 0, right: 0, i: 0, j: 0, out: 0,
            sort_work: 0, sort_total: 0, project_cursor: 0, project_pass: 0, result: Some(result), failed: false })
    }

    fn row_key(&self, row: usize) -> Result<MemberKey, ComputeError> {
        let kind = self.states[row];
        let number = self.values[row];
        let text = self.text_ids[row];
        if kind > SLICER_ERROR || self.availability[row] > (if self.params.mode == SlicerQueryMode::Canvas { 3 } else { 1 })
            || self.params.mode == SlicerQueryMode::Metadata && kind != SLICER_BLANK && kind != SLICER_DATE {
            return Err(ComputeError::invalid("Invalid typed Slicer row type/availability"));
        }
        if self.params.mode == SlicerQueryMode::Canvas && self.text_ids[self.params.row_count as usize + row] >= self.params.dictionary_count {
            return Err(ComputeError::invalid("Canvas source caption ID is outside the exact dictionary"));
        }
        let payload = match kind {
            SLICER_STRING | SLICER_ERROR => {
                if text >= self.params.dictionary_count || number != 0.0 {
                    return Err(ComputeError::invalid("Invalid typed Slicer text identity"));
                }
                text as u64
            }
            SLICER_NUMBER | SLICER_DATE | SLICER_BOOLEAN => {
                if text != SLICER_NO_TEXT || !number.is_finite()
                    || kind == SLICER_BOOLEAN && number != 0.0 && number != 1.0 {
                    return Err(ComputeError::invalid("Invalid typed Slicer numeric/boolean identity"));
                }
                if number == 0.0 { 0 } else { number.to_bits() }
            }
            _ => {
                if text != SLICER_NO_TEXT || number != 0.0 {
                    return Err(ComputeError::invalid("Blank Slicer rows must not carry a value"));
                }
                0
            }
        };
        Ok(MemberKey { kind, payload })
    }

    fn scan_row(&mut self) -> Result<(), ComputeError> {
        let row = self.row;
        let key = self.row_key(row)?;
        if self.params.mode == SlicerQueryMode::Canvas {
            // Selection is caption-based over all source rows, independent of
            // typed representative choice and hide/data-at-top projection.
            let caption = &mut self.captions[self.text_ids[self.params.row_count as usize + row] as usize];
            if self.availability[row] & 2 != 0 { caption.selected = true; } else { caption.unselected = true; }
        }
        if key.kind == SLICER_DATE {
            let result = self.result.as_mut().unwrap();
            let date = self.values[row];
            result.date_count += 1;
            result.min_date = Some(result.min_date.map_or(date, |bound| bound.min(date)));
            result.max_date = Some(result.max_date.map_or(date, |bound| bound.max(date)));
        }
        if self.params.mode != SlicerQueryMode::Metadata {
            let has_data = self.params.cross_filter == SlicerCrossFilter::None || self.availability[row] & 1 != 0;
            if let Some(index) = self.lookup.get(&key).copied() {
                self.members[index as usize].has_data |= has_data;
            } else {
                let index = self.members.len() as u32;
                self.lookup.insert(key, index);
                self.members.push(Member { first_row: row as u32, has_data });
                self.order.push(index);
            }
        }
        self.row += 1;
        Ok(())
    }

    fn start_sort(&mut self) {
        let count = self.members.len();
        self.scratch.truncate(count);
        let passes = if count < 2 { 0 } else { usize::BITS - (count - 1).leading_zeros() };
        self.sort_total = count as u64 * passes as u64;
        self.phase = if count < 2 { Phase::Project } else { Phase::Sort };
        self.start_merge();
    }

    fn start_merge(&mut self) {
        self.middle = (self.left + self.width).min(self.members.len());
        self.right = (self.left + self.width * 2).min(self.members.len());
        self.i = self.left;
        self.j = self.middle;
        self.out = self.left;
    }

    fn compare(&self, first: u32, second: u32) -> Ordering {
        let left = self.members[first as usize].first_row as usize;
        let right = self.members[second as usize].first_row as usize;
        let kind = self.states[left];
        let mut order = kind.cmp(&self.states[right]);
        if order == Ordering::Equal {
            order = match kind {
                SLICER_NUMBER | SLICER_DATE | SLICER_BOOLEAN => self.values[left].partial_cmp(&self.values[right]).unwrap(),
                SLICER_STRING | SLICER_ERROR => self.text_ranks[self.text_ids[left] as usize]
                    .cmp(&self.text_ranks[self.text_ids[right] as usize]),
                _ => Ordering::Equal,
            };
        }
        if self.params.sort_order == SlicerSortOrder::Descending { order.reverse() } else { order }
    }

    fn sort_write(&mut self) {
        let take_left = self.i < self.middle && (self.j >= self.right
            || self.compare(self.order[self.i], self.order[self.j]) != Ordering::Greater);
        self.scratch[self.out] = if take_left {
            let row = self.order[self.i]; self.i += 1; row
        } else {
            let row = self.order[self.j]; self.j += 1; row
        };
        self.out += 1;
        self.sort_work += 1;
        if self.out == self.right {
            self.left = self.right;
            if self.left >= self.members.len() {
                std::mem::swap(&mut self.order, &mut self.scratch);
                self.width *= 2;
                self.left = 0;
                if self.width >= self.members.len() { self.phase = Phase::Project; }
            }
            self.start_merge();
        }
    }

    fn two_pass_projection(&self) -> bool {
        self.params.cross_filter == SlicerCrossFilter::ShowItemsWithDataAtTop && !self.params.hide_items_with_no_data
    }

    fn project_member(&mut self) {
        let member = self.members[self.order[self.project_cursor] as usize];
        let cross_none = self.params.cross_filter == SlicerCrossFilter::None;
        let eligible = if cross_none { true }
            else if self.params.hide_items_with_no_data { member.has_data }
            else if self.two_pass_projection() { member.has_data == (self.project_pass == 0) }
            else { true };
        if eligible {
            if self.params.mode == SlicerQueryMode::Canvas {
                let id = self.text_ids[self.params.row_count as usize + member.first_row as usize];
                let caption = &mut self.captions[id as usize];
                if caption.first_row.is_none() { caption.first_row = Some(member.first_row); self.caption_order.push(id); }
                caption.has_data |= member.has_data;
            } else {
                let result = self.result.as_mut().unwrap();
                if result.total as u64 >= self.params.start && result.indices.len() < self.params.count as usize {
                    result.indices.push(member.first_row);
                    result.has_data.push(u8::from(member.has_data));
                }
                result.total += 1;
            }
        }
        self.project_cursor += 1;
    }

    fn project_caption(&mut self) {
        let caption = &self.captions[self.caption_order[self.canvas_cursor] as usize];
        let result = self.result.as_mut().unwrap();
        if self.canvas_cursor as u64 >= self.params.start && result.indices.len() < self.params.count as usize {
            result.indices.push(caption.first_row.unwrap());
            let checked = caption.selected && !caption.unselected;
            let partial = caption.selected && caption.unselected;
            result.has_data.push(u8::from(caption.has_data) | u8::from(checked) << 1 | u8::from(partial) << 2);
        }
        self.canvas_cursor += 1;
    }

    pub fn progress(&self) -> f64 {
        match self.phase {
            Phase::Scan => if self.values.is_empty() { 0.0 } else {
                let scan_weight = if self.params.mode == SlicerQueryMode::Metadata { 1.0 } else { 0.5 };
                scan_weight * self.row as f64 / self.values.len() as f64
            },
            Phase::Sort => 0.5 + (if self.params.mode == SlicerQueryMode::Canvas { 0.3 } else { 0.4 }) * self.sort_work as f64 / self.sort_total as f64,
            Phase::Project => {
                let passes = if self.two_pass_projection() { 2 } else { 1 };
                let total = self.members.len() * passes;
                let start = if self.params.mode == SlicerQueryMode::Canvas { 0.8 } else { 0.9 };
                if total == 0 { start } else { start + 0.1 * (self.project_pass as usize * self.members.len()
                    + self.project_cursor) as f64 / total as f64 }
            },
            Phase::CanvasOutput => if self.caption_order.is_empty() { 0.9 } else { 0.9 + 0.1 * self.canvas_cursor as f64 / self.caption_order.len() as f64 },
            Phase::Done => 1.0,
        }
    }

    pub fn step(&mut self, budget: u32) -> Result<bool, ComputeError> {
        if self.failed || self.result.is_none() {
            return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Slicer query was consumed or failed"));
        }
        let mut remaining = budget.clamp(1, 1_000_000);
        while self.phase != Phase::Done && remaining > 0 {
            match self.phase {
                Phase::Scan => {
                    if self.row < self.values.len() {
                        if let Err(error) = self.scan_row() { self.failed = true; return Err(error); }
                        remaining -= 1;
                    } else if self.params.mode == SlicerQueryMode::Metadata { self.phase = Phase::Done; }
                    else { self.start_sort(); }
                },
                Phase::Sort => { self.sort_write(); remaining -= 1; },
                Phase::Project => {
                    if self.project_cursor < self.members.len() { self.project_member(); remaining -= 1; }
                    else if self.two_pass_projection() && self.project_pass == 0 {
                        self.project_pass = 1; self.project_cursor = 0;
                    } else if self.params.mode == SlicerQueryMode::Canvas {
                        self.result.as_mut().unwrap().total = self.caption_order.len() as u32;
                        self.phase = Phase::CanvasOutput;
                    } else { self.phase = Phase::Done; }
                },
                Phase::CanvasOutput => {
                    if self.canvas_cursor < self.caption_order.len() { self.project_caption(); remaining -= 1; }
                    else { self.phase = Phase::Done; }
                },
                Phase::Done => {},
            }
        }
        Ok(self.phase == Phase::Done)
    }

    pub fn take_result(&mut self) -> Result<SlicerMembersResult, ComputeError> {
        if self.failed || self.result.is_none() {
            return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "Slicer query was consumed or failed"));
        }
        if self.phase != Phase::Done {
            return Err(ComputeError::new("COMPUTE_NOT_READY", "Slicer query has not completed"));
        }
        Ok(self.result.take().unwrap())
    }
}

#[cfg(test)]
#[path = "slicer/tests.rs"]
mod tests;
