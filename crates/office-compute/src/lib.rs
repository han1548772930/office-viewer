mod aggregate;
mod error;
mod filter;
mod sort;
#[cfg(target_arch = "wasm32")]
mod wasm_api;

pub use aggregate::{NumericAggregate, NumericStatistics};
pub use error::{ComputeError, MAX_ELEMENTS, MAX_KEYS, MAX_ROWS};
pub use filter::NumericFilter;
pub use sort::StableNumericSort;
#[cfg(target_arch = "wasm32")]
pub use wasm_api::*;

#[cfg(test)]
mod tests;
