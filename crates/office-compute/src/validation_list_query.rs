use crate::error::{MAX_ELEMENTS, MAX_ROWS};
use crate::ComputeError;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ValidationListQueryParams {
    pub row_count: u32,
    pub text_byte_length: u32,
    pub query_byte_length: u32,
    pub start: u32,
    pub count: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ValidationListQueryMemory {
    pub input_bytes: usize,
    pub output_bytes: usize,
    pub estimated_peak_bytes: usize,
}

impl ValidationListQueryParams {
    pub fn validate_shape(&self, offsets: usize, text: usize, query: usize) -> Result<(), ComputeError> {
        if self.row_count as usize > MAX_ROWS || self.count == 0 || self.count as usize > MAX_ROWS
            || self.text_byte_length % 2 != 0 || self.query_byte_length % 2 != 0
            || self.text_byte_length as usize / 2 > MAX_ELEMENTS || self.query_byte_length as usize / 2 > MAX_ELEMENTS
            || offsets != self.row_count as usize + 1
            || text != self.text_byte_length as usize || query != self.query_byte_length as usize
        {
            return Err(ComputeError::invalid("List query dimensions or UTF-16 byte lengths are invalid"));
        }
        Ok(())
    }

    pub fn memory_estimate(&self) -> Result<ValidationListQueryMemory, ComputeError> {
        let input_bytes = (self.row_count as usize).checked_add(1).and_then(|n| n.checked_mul(4))
            .and_then(|n| n.checked_add(self.text_byte_length as usize))
            .and_then(|n| n.checked_add(self.query_byte_length as usize)).ok_or_else(ComputeError::memory)?;
        let output_bytes = (self.count.min(self.row_count.saturating_sub(self.start)) as usize).checked_mul(4).ok_or_else(ComputeError::memory)?;
        let prefix_bytes = (self.query_byte_length as usize).checked_mul(2).ok_or_else(ComputeError::memory)?;
        let independent_output_bytes = output_bytes.checked_mul(2).ok_or_else(ComputeError::memory)?;
        let estimated_peak_bytes = input_bytes.checked_mul(2)
            .and_then(|n| n.checked_add(prefix_bytes))
            .and_then(|n| n.checked_add(independent_output_bytes))
            .and_then(|n| n.checked_add(8 * 1024 * 1024)).ok_or_else(ComputeError::memory)?;
        Ok(ValidationListQueryMemory { input_bytes, output_bytes, estimated_peak_bytes })
    }
}

#[derive(Debug, PartialEq, Eq)]
pub struct ValidationListQueryResult {
    pub indices: Vec<u32>,
    pub total: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Phase { Validate, Prefix, Search, Ready, Consumed }

/// Exact UTF-16LE substring matching. JS owns locale folding, so Rust never
/// replaces lone surrogates or invents a different Unicode/case contract.
/// KMP prefix fallbacks, dictionary checks and text work each consume a step
/// budget unit. Inputs and the bounded output window have one retained owner.
pub struct ValidationListQuery {
    offsets: Vec<u32>,
    text: Vec<u8>,
    query: Vec<u8>,
    prefix: Vec<u32>,
    params: ValidationListQueryParams,
    phase: Phase,
    validated: usize,
    prefix_index: usize,
    prefix_match: usize,
    row: u32,
    cursor: usize,
    matched: usize,
    total: u32,
    indices: Vec<u32>,
}

fn unit(bytes: &[u8], index: usize) -> u16 {
    u16::from_le_bytes([bytes[index * 2], bytes[index * 2 + 1]])
}

impl ValidationListQuery {
    pub fn new(offsets: Vec<u32>, text: Vec<u8>, query: Vec<u8>, params: ValidationListQueryParams,
        memory_budget: usize) -> Result<Self, ComputeError> {
        params.validate_shape(offsets.len(), text.len(), query.len())?;
        if params.memory_estimate()?.estimated_peak_bytes > memory_budget { return Err(ComputeError::memory()); }
        if offsets.first() != Some(&0) || offsets.last().copied() != Some(params.text_byte_length) {
            return Err(ComputeError::invalid("List query dictionary endpoints are invalid"));
        }
        let mut prefix = Vec::new();
        prefix.try_reserve_exact(query.len() / 2).map_err(|_| ComputeError::memory())?;
        prefix.resize(query.len() / 2, 0);
        let mut indices = Vec::new();
        indices.try_reserve_exact(params.count.min(params.row_count.saturating_sub(params.start)) as usize).map_err(|_| ComputeError::memory())?;
        Ok(Self { offsets, text, query, prefix, params, phase: Phase::Validate, validated: 0,
            prefix_index: 1, prefix_match: 0, row: 0, cursor: 0, matched: 0, total: 0, indices })
    }

    pub fn step(&mut self, budget: u32) -> Result<bool, ComputeError> {
        if self.phase == Phase::Consumed { return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "List query was consumed")); }
        if budget == 0 { return Err(ComputeError::invalid("List query step budget must be positive")); }
        let mut remaining = budget;
        while remaining > 0 && self.phase != Phase::Ready {
            match self.phase {
                Phase::Validate => {
                    let offset = self.offsets[self.validated];
                    if offset % 2 != 0 || offset > self.params.text_byte_length
                        || self.validated > 0 && offset < self.offsets[self.validated - 1]
                    {
                        return Err(ComputeError::invalid("List query dictionary offsets are invalid"));
                    }
                    self.validated += 1;
                    if self.validated == self.offsets.len() { self.phase = Phase::Prefix; }
                    remaining -= 1;
                }
                Phase::Prefix => {
                    if self.prefix_index >= self.prefix.len() { self.phase = Phase::Search; continue; }
                    if unit(&self.query, self.prefix_index) == unit(&self.query, self.prefix_match) {
                        self.prefix_match += 1;
                        self.prefix[self.prefix_index] = self.prefix_match as u32;
                        self.prefix_index += 1;
                    } else if self.prefix_match > 0 {
                        self.prefix_match = self.prefix[self.prefix_match - 1] as usize;
                    } else {
                        self.prefix[self.prefix_index] = 0;
                        self.prefix_index += 1;
                    }
                    remaining -= 1;
                }
                Phase::Search => {
                    if self.row >= self.params.row_count { self.phase = Phase::Ready; continue; }
                    if self.query.is_empty() {
                        self.finish_row(true);
                    } else if self.cursor >= self.offsets[self.row as usize + 1] as usize {
                        self.finish_row(false);
                    } else if unit(&self.text, self.cursor / 2) == unit(&self.query, self.matched) {
                        self.cursor += 2;
                        self.matched += 1;
                        if self.matched == self.prefix.len() { self.finish_row(true); }
                    } else if self.matched > 0 {
                        self.matched = self.prefix[self.matched - 1] as usize;
                    } else {
                        self.cursor += 2;
                    }
                    remaining -= 1;
                }
                Phase::Ready | Phase::Consumed => break,
            }
        }
        Ok(self.phase == Phase::Ready)
    }

    fn finish_row(&mut self, matches: bool) {
        if matches {
            if self.total >= self.params.start && self.total < self.params.start.saturating_add(self.params.count) {
                self.indices.push(self.row);
            }
            self.total += 1;
        }
        self.row += 1;
        self.matched = 0;
        if self.row < self.params.row_count { self.cursor = self.offsets[self.row as usize] as usize; }
        else { self.phase = Phase::Ready; }
    }

    /// Progress counts fully classified source rows, not a timer or an estimate
    /// of comparisons. Prefix and offset validation remain at zero until rows finish.
    pub fn progress(&self) -> f64 {
        if matches!(self.phase, Phase::Ready | Phase::Consumed) { return 1.0; }
        if self.params.row_count == 0 { return 0.0; }
        self.row as f64 / self.params.row_count as f64
    }

    pub fn take_result(&mut self) -> Result<ValidationListQueryResult, ComputeError> {
        if self.phase == Phase::Consumed { return Err(ComputeError::new("COMPUTE_RESULT_CONSUMED", "List query was consumed")); }
        if self.phase != Phase::Ready { return Err(ComputeError::new("COMPUTE_NOT_READY", "List query is unfinished")); }
        self.phase = Phase::Consumed;
        Ok(ValidationListQueryResult { indices: std::mem::take(&mut self.indices), total: self.total })
    }
}

#[cfg(test)]
mod tests;
