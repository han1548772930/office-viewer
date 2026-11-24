mod aggregate;
mod chart;
mod error;
mod filter;
mod sort;
#[cfg(target_arch = "wasm32")]
mod wasm_api;
#[cfg(target_arch = "wasm32")]
mod chart_wasm;

pub use aggregate::{NumericAggregate, NumericStatistics};
pub use chart::{ChartJob, ChartMemoryEstimate, ChartSegment, ChartSeries, ChartStatistics,
    ChartStatisticsParams, ChartStatisticsResult, ChartTrendlineFit, ErrorDirection,
    ErrorSides, ErrorValueType, HistogramBinning, QuartileMethod, TrendType, WaterfallStackType};
pub use error::{ComputeError, MAX_ELEMENTS, MAX_KEYS, MAX_ROWS};
pub use filter::NumericFilter;
pub use sort::StableNumericSort;
#[cfg(target_arch = "wasm32")]
pub use wasm_api::*;
#[cfg(target_arch = "wasm32")]
pub use chart_wasm::*;

#[cfg(test)]
mod tests;
