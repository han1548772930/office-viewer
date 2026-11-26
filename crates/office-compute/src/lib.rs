mod aggregate;
mod chart;
mod error;
mod filter;
mod sort;
mod slicer;
#[cfg(target_arch = "wasm32")]
mod wasm_api;
#[cfg(target_arch = "wasm32")]
mod chart_wasm;
#[cfg(target_arch = "wasm32")]
mod slicer_wasm;

pub use aggregate::{NumericAggregate, NumericStatistics};
pub use chart::{ChartJob, ChartMemoryEstimate, ChartSegment, ChartSeries, ChartStatistics,
    ChartStatisticsParams, ChartStatisticsResult, ChartTrendlineFit, ErrorDirection,
    ErrorSides, ErrorValueType, HistogramBinning, QuartileMethod, TrendType, WaterfallStackType};
pub use error::{ComputeError, MAX_ELEMENTS, MAX_KEYS, MAX_ROWS};
pub use filter::NumericFilter;
pub use sort::StableNumericSort;
pub use slicer::{SlicerMembers, SlicerMembersParams, SlicerMembersResult, SlicerMembersMemory,
    SlicerQueryMode, SlicerSortOrder, SlicerCrossFilter, SLICER_NO_TEXT, SLICER_BLANK, SLICER_NUMBER,
    SLICER_DATE, SLICER_STRING, SLICER_BOOLEAN, SLICER_ERROR};
#[cfg(target_arch = "wasm32")]
pub use wasm_api::*;
#[cfg(target_arch = "wasm32")]
pub use chart_wasm::*;
#[cfg(target_arch = "wasm32")]
pub use slicer_wasm::*;

#[cfg(test)]
mod tests;
