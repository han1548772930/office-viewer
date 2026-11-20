use std::fmt;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComputeError {
    pub code: &'static str,
    pub message: String,
}

impl ComputeError {
    pub fn new(code: &'static str, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
        }
    }

    pub fn invalid(message: impl Into<String>) -> Self {
        Self::new("COMPUTE_INVALID_INPUT", message)
    }

    pub fn memory() -> Self {
        Self::new(
            "COMPUTE_MEMORY_LIMIT",
            "Calculation exceeds the configured memory limit",
        )
    }
}

impl fmt::Display for ComputeError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}: {}", self.code, self.message)
    }
}

impl std::error::Error for ComputeError {}

pub const MAX_ELEMENTS: usize = 16_000_000;
pub const MAX_ROWS: usize = 10_000_000;
pub const MAX_KEYS: usize = 32;

pub fn validate_states(states: &[u8]) -> Result<(), ComputeError> {
    if states.iter().any(|state| *state > 3) {
        return Err(ComputeError::invalid(
            "Cell states must be 0=empty, 1=number, 2=other, 3=error",
        ));
    }
    Ok(())
}

pub fn validate_table(
    values: &[f64],
    states: &[u8],
    rows: usize,
    keys: usize,
) -> Result<(), ComputeError> {
    if rows > MAX_ROWS || keys > MAX_KEYS {
        return Err(ComputeError::memory());
    }
    if keys == 0 {
        return Err(ComputeError::invalid("At least one key is required"));
    }
    let length = rows.checked_mul(keys).ok_or_else(ComputeError::memory)?;
    if length > MAX_ELEMENTS {
        return Err(ComputeError::memory());
    }
    if values.len() != length || states.len() != length {
        return Err(ComputeError::invalid(
            "Column-major values/states must have rows × keys elements",
        ));
    }
    validate_states(states)
}

pub fn indices(rows: usize) -> Result<Vec<u32>, ComputeError> {
    let mut result = Vec::new();
    result
        .try_reserve_exact(rows)
        .map_err(|_| ComputeError::memory())?;
    result.extend((0..rows).map(|row| row as u32));
    Ok(result)
}
