use crate::error::{ComputeError, MAX_ELEMENTS, validate_states};

/// Incremental finite-number statistics. Non-numeric values are never coerced.
#[derive(Debug, Default)]
pub struct NumericAggregate {
    processed_count: usize,
    numeric_count: u64,
    non_empty_count: u64,
    empty_count: u64,
    error_count: u64,
    invalid_numeric_count: u64,
    sum: f64,
    compensation: f64,
    running_mean: f64,
    m2: f64,
    min: Option<f64>,
    max: Option<f64>,
}

impl NumericAggregate {
    pub fn push(&mut self, values: &[f64], states: &[u8]) -> Result<(), ComputeError> {
        if values.len() != states.len() {
            return Err(ComputeError::invalid("Values and states lengths differ"));
        }
        let total = self
            .processed_count
            .checked_add(values.len())
            .ok_or_else(ComputeError::memory)?;
        if total > MAX_ELEMENTS {
            return Err(ComputeError::memory());
        }
        // Validate first: failure cannot leave a partially updated accumulator.
        validate_states(states)?;
        self.processed_count = total;
        for (&value, &state) in values.iter().zip(states) {
            if state == 0 {
                self.empty_count += 1;
                continue;
            }
            self.non_empty_count += 1;
            if state == 3 {
                self.error_count += 1;
            }
            if state != 1 {
                continue;
            }
            if !value.is_finite() {
                self.invalid_numeric_count += 1;
                continue;
            }
            self.numeric_count += 1;
            let next = self.sum + value;
            self.compensation += if self.sum.abs() >= value.abs() {
                (self.sum - next) + value
            } else {
                (value - next) + self.sum
            };
            self.sum = next;
            let delta = value - self.running_mean;
            let count = self.numeric_count as f64;
            // Opposite, very large finite inputs can overflow their difference
            // even when the updated mean is representable.
            self.running_mean = if delta.is_finite() {
                self.running_mean + delta / count
            } else {
                self.running_mean * ((count - 1.0) / count) + value / count
            };
            self.m2 += delta * (value - self.running_mean);
            self.min = Some(self.min.map_or(value, |min| min.min(value)));
            self.max = Some(self.max.map_or(value, |max| max.max(value)));
        }
        Ok(())
    }

    /// Slot schema v1: counts(5), sum, mean, min, max, population/sample variance,
    /// sample standard deviation, overflow flag. NaN means no defined result.
    pub fn packed_result(&self) -> Vec<f64> {
        let sum = self.sum + self.compensation;
        let mean = if self.numeric_count == 0 {
            f64::NAN
        } else if sum.is_finite() {
            sum / self.numeric_count as f64
        } else {
            self.running_mean
        };
        // f64::max treats NaN as the other operand: do not turn overflow into
        // an apparently valid zero variance.
        let m2 = if self.m2.is_finite() {
            self.m2.max(0.0)
        } else {
            self.m2
        };
        let population = if self.numeric_count > 0 {
            m2 / self.numeric_count as f64
        } else {
            f64::NAN
        };
        let sample = if self.numeric_count > 1 {
            m2 / (self.numeric_count - 1) as f64
        } else {
            f64::NAN
        };
        let overflow = !sum.is_finite()
            || (self.numeric_count > 0 && (!mean.is_finite() || !population.is_finite()))
            || (self.numeric_count > 1 && !sample.is_finite());
        vec![
            self.numeric_count as f64,
            self.non_empty_count as f64,
            self.empty_count as f64,
            self.error_count as f64,
            self.invalid_numeric_count as f64,
            sum,
            mean,
            self.min.unwrap_or(f64::NAN),
            self.max.unwrap_or(f64::NAN),
            population,
            sample,
            sample.sqrt(),
            u8::from(overflow) as f64,
        ]
    }
}
