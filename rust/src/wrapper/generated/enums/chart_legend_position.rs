use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartLegendPosition` enum defines the {@link Chart} legend positions.
///
/// Excel allows the following positions of the chart legend:
///
/// These positions can be set using the {@link ChartLegend#setPosition} method
/// and the `ChartLegendPosition` enum values.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartLegendPosition {
    /// Chart legend positioned at the right side. The default.
    Right,
    /// Chart legend positioned at the left side.
    Left,
    /// Chart legend positioned at the top.
    Top,
    /// Chart legend positioned at the bottom.
    Bottom,
    /// Chart legend positioned at the top right.
    TopRight,
}

impl From<ChartLegendPosition> for xlsx::ChartLegendPosition {
    fn from(value: ChartLegendPosition) -> xlsx::ChartLegendPosition {
        match value {
            ChartLegendPosition::Right => xlsx::ChartLegendPosition::Right,
            ChartLegendPosition::Left => xlsx::ChartLegendPosition::Left,
            ChartLegendPosition::Top => xlsx::ChartLegendPosition::Top,
            ChartLegendPosition::Bottom => xlsx::ChartLegendPosition::Bottom,
            ChartLegendPosition::TopRight => xlsx::ChartLegendPosition::TopRight,
        }
    }
}
