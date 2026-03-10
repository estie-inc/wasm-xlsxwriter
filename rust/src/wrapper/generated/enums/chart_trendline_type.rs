use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartTrendlineType` enum defines the trendline types of a
/// ChartSeries.
///
/// The following are the trendline types supported by Excel.
///
/// <img src="https://rustxlsxwriter.github.io/images/trendline_types.png">
///
/// The trendline type is used in conjunction with the
/// ChartTrendline.setType() method and a ChartSeries.
#[derive(Debug, Clone, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartTrendlineType {
    /// Don't show any trendline for the data series. The default.
    None,
    /// Display an exponential fit trendline.
    Exponential,
    /// Display a linear best fit trendline.
    Linear,
    /// Display a logarithmic best fit trendline.
    Logarithmic,
    /// Display a polynomial fit trendline. The order of the polynomial can be
    /// specified in the range 2-6.
    Polynomial(u8),
    /// Display a power fit trendline.
    Power,
    /// Display a moving average trendline. The period of the moving average can
    /// be specified in the range 2-4.
    MovingAverage(u8),
}

impl From<ChartTrendlineType> for xlsx::ChartTrendlineType {
    fn from(value: ChartTrendlineType) -> xlsx::ChartTrendlineType {
        match value {
            ChartTrendlineType::None => xlsx::ChartTrendlineType::None,
            ChartTrendlineType::Exponential => xlsx::ChartTrendlineType::Exponential,
            ChartTrendlineType::Linear => xlsx::ChartTrendlineType::Linear,
            ChartTrendlineType::Logarithmic => xlsx::ChartTrendlineType::Logarithmic,
            ChartTrendlineType::Polynomial(v0) => xlsx::ChartTrendlineType::Polynomial(v0),
            ChartTrendlineType::Power => xlsx::ChartTrendlineType::Power,
            ChartTrendlineType::MovingAverage(v0) => xlsx::ChartTrendlineType::MovingAverage(v0),
        }
    }
}
