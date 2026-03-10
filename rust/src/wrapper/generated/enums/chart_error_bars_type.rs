use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartErrorBarsType` enum defines the type of a chart series
/// ChartErrorBars.
///
/// The following enum values represent the error bar types that are available
/// in Excel.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/chart_error_bars_types.png">
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartErrorBarsType {
    /// Set a fixed value for the positive and negative error bars. In Excel
    /// this must be > 0.0.
    FixedValue(f64),
    /// Set a percentage for the positive and negative error bars. In Excel this
    /// must be >= 0.0.
    Percentage(f64),
    /// Set a multiple of the standard deviation for the positive and negative
    /// error bars. In Excel this must be >= 0.0.
    StandardDeviation(f64),
    /// Set a the standard error value for the positive and negative error bars.
    /// This is the default.
    StandardError,
    /// Set custom values for the error bars based on a range of worksheet
    /// values like `Sheet1!$B$1:$B$3` (single value) or `Sheet1!$B$1:$B$5` (a
    /// range to match the number of point in the series). Single values are
    /// repeated for each point in the chart, like `FixedValue`. The `plus` and
    /// `minus` values must be set separately using ChartRange instances.
    Custom(ChartRange, ChartRange),
}

impl From<ChartErrorBarsType> for xlsx::ChartErrorBarsType {
    fn from(value: ChartErrorBarsType) -> xlsx::ChartErrorBarsType {
        match value {
            ChartErrorBarsType::FixedValue(v0) => xlsx::ChartErrorBarsType::FixedValue(v0),
            ChartErrorBarsType::Percentage(v0) => xlsx::ChartErrorBarsType::Percentage(v0),
            ChartErrorBarsType::StandardDeviation(v0) => {
                xlsx::ChartErrorBarsType::StandardDeviation(v0)
            }
            ChartErrorBarsType::StandardError => xlsx::ChartErrorBarsType::StandardError,
            ChartErrorBarsType::Custom(v0, v1) => xlsx::ChartErrorBarsType::Custom(v0, v1),
        }
    }
}
