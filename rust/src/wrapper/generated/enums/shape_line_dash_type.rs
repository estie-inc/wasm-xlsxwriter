use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ShapeLineDashType` enum defines the Shape line dash types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ShapeLineDashType {
    /// Solid - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_solid.png">
    Solid,
    /// Round dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_round_dot.png">
    RoundDot,
    /// Square dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_square_dot.png">
    SquareDot,
    /// Dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash.png">
    Dash,
    /// Dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash_dot.png">
    DashDot,
    /// Long dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash.png">
    LongDash,
    /// Long dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot.png">
    LongDashDot,
    /// Long dash dot dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot_dot.png">
    LongDashDotDot,
}

impl From<ShapeLineDashType> for xlsx::ShapeLineDashType {
    fn from(value: ShapeLineDashType) -> xlsx::ShapeLineDashType {
        match value {
            ShapeLineDashType::Solid => xlsx::ShapeLineDashType::Solid,
            ShapeLineDashType::RoundDot => xlsx::ShapeLineDashType::RoundDot,
            ShapeLineDashType::SquareDot => xlsx::ShapeLineDashType::SquareDot,
            ShapeLineDashType::Dash => xlsx::ShapeLineDashType::Dash,
            ShapeLineDashType::DashDot => xlsx::ShapeLineDashType::DashDot,
            ShapeLineDashType::LongDash => xlsx::ShapeLineDashType::LongDash,
            ShapeLineDashType::LongDashDot => xlsx::ShapeLineDashType::LongDashDot,
            ShapeLineDashType::LongDashDotDot => xlsx::ShapeLineDashType::LongDashDotDot,
        }
    }
}
