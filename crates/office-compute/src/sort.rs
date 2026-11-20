use crate::error::{ComputeError, indices, validate_table};
use std::cmp::Ordering;

/// Stable bottom-up merge sort. A step does bounded writes, allowing the Worker
/// to yield between calls rather than pretending a synchronous sort is cancellable.
pub struct StableNumericSort {
    values: Vec<f64>,
    states: Vec<u8>,
    rows: usize,
    directions: Vec<u8>,
    nulls_last: bool,
    order: Vec<u32>,
    scratch: Vec<u32>,
    width: usize,
    left: usize,
    middle: usize,
    right: usize,
    i: usize,
    j: usize,
    out: usize,
    work: u64,
    total_work: u64,
}

impl StableNumericSort {
    pub fn new(
        values: Vec<f64>,
        states: Vec<u8>,
        rows: usize,
        directions: Vec<u8>,
        nulls_last: bool,
    ) -> Result<Self, ComputeError> {
        validate_table(&values, &states, rows, directions.len())?;
        if directions.iter().any(|direction| *direction > 1) {
            return Err(ComputeError::invalid(
                "Sort directions must be 0=ascending or 1=descending",
            ));
        }
        if states.contains(&2) {
            return Err(ComputeError::new(
                "COMPUTE_UNSUPPORTED_TYPE",
                "Numeric sorting requires numeric/rank keys; text must be collated by the feature adapter",
            ));
        }
        let order = indices(rows)?;
        let mut scratch = Vec::new();
        scratch
            .try_reserve_exact(rows)
            .map_err(|_| ComputeError::memory())?;
        scratch.resize(rows, 0);
        let passes = if rows < 2 {
            0
        } else {
            usize::BITS - (rows - 1).leading_zeros()
        };
        let mut sorter = Self {
            values,
            states,
            rows,
            directions,
            nulls_last,
            order,
            scratch,
            width: 1,
            left: 0,
            middle: 0,
            right: 0,
            i: 0,
            j: 0,
            out: 0,
            work: 0,
            total_work: rows as u64 * passes as u64,
        };
        sorter.start_merge();
        Ok(sorter)
    }

    fn start_merge(&mut self) {
        self.middle = (self.left + self.width).min(self.rows);
        self.right = (self.left + self.width * 2).min(self.rows);
        self.i = self.left;
        self.j = self.middle;
        self.out = self.left;
    }

    fn compare(&self, left: u32, right: u32) -> Ordering {
        for (key, direction) in self.directions.iter().enumerate() {
            let l = key * self.rows + left as usize;
            let r = key * self.rows + right as usize;
            let lv = self.values[l];
            let rv = self.values[r];
            let lnull = self.states[l] != 1 || !lv.is_finite();
            let rnull = self.states[r] != 1 || !rv.is_finite();
            let order = match (lnull, rnull) {
                (true, true) => Ordering::Equal,
                (true, false) => {
                    if self.nulls_last {
                        Ordering::Greater
                    } else {
                        Ordering::Less
                    }
                }
                (false, true) => {
                    if self.nulls_last {
                        Ordering::Less
                    } else {
                        Ordering::Greater
                    }
                }
                (false, false) => {
                    let numeric = lv.partial_cmp(&rv).unwrap_or(Ordering::Equal);
                    if *direction == 1 {
                        numeric.reverse()
                    } else {
                        numeric
                    }
                }
            };
            if order != Ordering::Equal {
                return order;
            }
        }
        Ordering::Equal
    }

    pub fn is_done(&self) -> bool {
        self.rows < 2 || self.width >= self.rows
    }
    pub fn progress(&self) -> f64 {
        if self.total_work == 0 {
            1.0
        } else {
            self.work as f64 / self.total_work as f64
        }
    }
    pub fn step(&mut self, budget: u32) -> bool {
        let mut remaining = budget.clamp(1, 1_000_000);
        while !self.is_done() && remaining > 0 {
            let take_left = self.i < self.middle
                && (self.j >= self.right
                    || self.compare(self.order[self.i], self.order[self.j]) != Ordering::Greater);
            self.scratch[self.out] = if take_left {
                let row = self.order[self.i];
                self.i += 1;
                row
            } else {
                let row = self.order[self.j];
                self.j += 1;
                row
            };
            self.out += 1;
            self.work += 1;
            remaining -= 1;
            if self.out == self.right {
                self.left = self.right;
                if self.left >= self.rows {
                    std::mem::swap(&mut self.order, &mut self.scratch);
                    self.width *= 2;
                    self.left = 0;
                }
                self.start_merge();
            }
        }
        self.is_done()
    }

    pub fn take_result(&mut self) -> Result<Vec<u32>, ComputeError> {
        if !self.is_done() {
            return Err(ComputeError::new(
                "COMPUTE_NOT_READY",
                "Sort has not finished",
            ));
        }
        Ok(std::mem::take(&mut self.order))
    }
}
