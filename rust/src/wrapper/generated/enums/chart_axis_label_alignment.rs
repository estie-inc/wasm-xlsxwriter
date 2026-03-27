use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartAxisLabelAlignment` enum defines the {@link ChartAxis} crossing point for
/// the opposite axis.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartAxisLabelAlignment {
    /// Center the axis label with the tick mark. This is the default.
    Center,
    /// Set the axis label to the left of the tick mark.
    Left,
    /// Set the axis label to the right of the tick mark.
    Right,
}

impl From<ChartAxisLabelAlignment> for xlsx::ChartAxisLabelAlignment {
    fn from(value: ChartAxisLabelAlignment) -> xlsx::ChartAxisLabelAlignment {
        match value {
            ChartAxisLabelAlignment::Center => xlsx::ChartAxisLabelAlignment::Center,
            ChartAxisLabelAlignment::Left => xlsx::ChartAxisLabelAlignment::Left,
            ChartAxisLabelAlignment::Right => xlsx::ChartAxisLabelAlignment::Right,
        }
    }
}
