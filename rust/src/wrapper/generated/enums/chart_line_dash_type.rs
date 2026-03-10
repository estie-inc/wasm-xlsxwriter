use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartLineDashType` enum defines the Chart line dash types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartLineDashType {
    /// Solid - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_solid.png">
    Solid,
    /// Round dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_round_dot.png">
    RoundDot,
    /// Square dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_square_dot.png">
    SquareDot,
    /// Dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash.png">
    Dash,
    /// Dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash_dot.png">
    DashDot,
    /// Long dash - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash.png">
    LongDash,
    /// Long dash dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot.png">
    LongDashDot,
    /// Long dash dot dot - chart line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot_dot.png">
    LongDashDotDot,
}

impl From<ChartLineDashType> for xlsx::ChartLineDashType {
    fn from(value: ChartLineDashType) -> xlsx::ChartLineDashType {
        match value {
            ChartLineDashType::Solid => xlsx::ChartLineDashType::Solid,
            ChartLineDashType::RoundDot => xlsx::ChartLineDashType::RoundDot,
            ChartLineDashType::SquareDot => xlsx::ChartLineDashType::SquareDot,
            ChartLineDashType::Dash => xlsx::ChartLineDashType::Dash,
            ChartLineDashType::DashDot => xlsx::ChartLineDashType::DashDot,
            ChartLineDashType::LongDash => xlsx::ChartLineDashType::LongDash,
            ChartLineDashType::LongDashDot => xlsx::ChartLineDashType::LongDashDot,
            ChartLineDashType::LongDashDotDot => xlsx::ChartLineDashType::LongDashDotDot,
        }
    }
}
