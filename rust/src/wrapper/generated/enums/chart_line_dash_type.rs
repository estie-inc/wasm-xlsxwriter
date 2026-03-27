use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartLineDashType` enum defines the {@link Chart} line dash types.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartLineDashType {
    /// Solid - chart line/border dash type.
    Solid,
    /// Round dot - chart line/border dash type.
    RoundDot,
    /// Square dot - chart line/border dash type.
    SquareDot,
    /// Dash - chart line/border dash type.
    Dash,
    /// Dash dot - chart line/border dash type.
    DashDot,
    /// Long dash - chart line/border dash type.
    LongDash,
    /// Long dash dot - chart line/border dash type.
    LongDashDot,
    /// Long dash dot dot - chart line/border dash type.
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
