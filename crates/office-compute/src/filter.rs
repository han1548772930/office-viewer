use crate::error::{ComputeError, validate_table};

/// Predicates: eq, ne, lt, le, gt, ge, between, notBetween, isEmpty,
/// isNotEmpty, isError. Operands are finite for numeric predicates.
pub struct NumericFilter {
    values: Vec<f64>,
    states: Vec<u8>,
    rows: usize,
    operators: Vec<u8>,
    lower: Vec<f64>,
    upper: Vec<f64>,
    all: bool,
    position: usize,
    matches: Vec<u32>,
}

impl NumericFilter {
    pub fn new(
        values: Vec<f64>,
        states: Vec<u8>,
        rows: usize,
        operators: Vec<u8>,
        lower: Vec<f64>,
        upper: Vec<f64>,
        all: bool,
    ) -> Result<Self, ComputeError> {
        validate_table(&values, &states, rows, operators.len())?;
        if lower.len() != operators.len() || upper.len() != operators.len() {
            return Err(ComputeError::invalid(
                "Each predicate needs lower and upper operand slots",
            ));
        }
        for (key, &operator) in operators.iter().enumerate() {
            if operator > 10 {
                return Err(ComputeError::invalid("Unknown filter operator"));
            }
            if operator < 8 && !lower[key].is_finite() {
                return Err(ComputeError::invalid("Numeric operands must be finite"));
            }
            if (operator == 6 || operator == 7)
                && (!upper[key].is_finite() || upper[key] < lower[key])
            {
                return Err(ComputeError::invalid(
                    "Between upper bound must be finite and at least its lower bound",
                ));
            }
        }
        let mut matches = Vec::new();
        matches
            .try_reserve_exact(rows)
            .map_err(|_| ComputeError::memory())?;
        Ok(Self {
            values,
            states,
            rows,
            operators,
            lower,
            upper,
            all,
            position: 0,
            matches,
        })
    }

    fn matches(&self, row: usize, key: usize) -> bool {
        let index = key * self.rows + row;
        let state = self.states[index];
        let value = self.values[index];
        match self.operators[key] {
            8 => state == 0,
            9 => state != 0,
            10 => state == 3,
            operator if state == 1 && value.is_finite() => match operator {
                0 => value == self.lower[key],
                1 => value != self.lower[key],
                2 => value < self.lower[key],
                3 => value <= self.lower[key],
                4 => value > self.lower[key],
                5 => value >= self.lower[key],
                6 => value >= self.lower[key] && value <= self.upper[key],
                7 => value < self.lower[key] || value > self.upper[key],
                _ => false,
            },
            _ => false,
        }
    }

    pub fn progress(&self) -> f64 {
        if self.rows == 0 {
            1.0
        } else {
            self.position as f64 / self.rows as f64
        }
    }
    pub fn step(&mut self, row_budget: u32) -> bool {
        let end = self
            .position
            .saturating_add(row_budget.clamp(1, 1_000_000) as usize)
            .min(self.rows);
        for row in self.position..end {
            let matches = if self.all {
                (0..self.operators.len()).all(|key| self.matches(row, key))
            } else {
                (0..self.operators.len()).any(|key| self.matches(row, key))
            };
            if matches {
                self.matches.push(row as u32);
            }
        }
        self.position = end;
        self.position == self.rows
    }

    pub fn take_result(&mut self) -> Result<Vec<u32>, ComputeError> {
        if self.position != self.rows {
            return Err(ComputeError::new(
                "COMPUTE_NOT_READY",
                "Filter has not finished",
            ));
        }
        Ok(std::mem::take(&mut self.matches))
    }
}
