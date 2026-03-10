use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisLabelPosition` enum defines the {@link Chart} axis label
/// positions.
///
/// This property is used in conjunction with
/// {@link ChartAxis#setLabelPosition}.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisLabelPosition {
    /// Position the axis labels next to the axis. The default.
    NextTo,
    /// Position the axis labels at the top of the chart, for horizontal axes,
    /// or to the right for vertical axes.
    High,
    /// Position the axis labels at the bottom of the chart, for horizontal
    /// axes, or to the left for vertical axes.
    Low,
    /// Turn off the the axis labels.
    None,
}

impl From<ChartAxisLabelPosition> for xlsx::ChartAxisLabelPosition {
    fn from(value: ChartAxisLabelPosition) -> xlsx::ChartAxisLabelPosition {
        match value {
            ChartAxisLabelPosition::NextTo => xlsx::ChartAxisLabelPosition::NextTo,
            ChartAxisLabelPosition::High => xlsx::ChartAxisLabelPosition::High,
            ChartAxisLabelPosition::Low => xlsx::ChartAxisLabelPosition::Low,
            ChartAxisLabelPosition::None => xlsx::ChartAxisLabelPosition::None,
        }
    }
}
