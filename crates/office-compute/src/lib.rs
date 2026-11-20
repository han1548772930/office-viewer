mod aggregate;
mod error;
mod filter;
mod sort;
mod wasm_api;

pub use aggregate::NumericAggregate;
pub use error::{ComputeError, MAX_ELEMENTS, MAX_KEYS, MAX_ROWS};
pub use filter::NumericFilter;
pub use sort::StableNumericSort;
pub use wasm_api::*;

#[cfg(test)]
mod tests;
