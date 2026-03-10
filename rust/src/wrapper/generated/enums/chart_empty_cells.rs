use rust_xlsxwriter as xlsx;
use serde::{Deserialize, Serialize};
use tsify::Tsify;

/// The `ChartEmptyCells` enum defines the Chart empty cell options.
///
/// This enum defines the Excel chart options for handling empty cell in the
/// chart data ranges.
///
/// These options can be set using the Chart.showEmptyCellsAs() method.
#[derive(Debug, Clone, Copy, Serialize, Deserialize, Tsify)]
#[tsify(into_wasm_abi, from_wasm_abi)]
pub enum ChartEmptyCells {
    /// Show empty cells in the chart as gaps. The default.
    Gaps,
    /// Show empty cells in the chart as zeroes.
    Zero,
    /// Show empty cells in the chart connected by a line to the previous point.
    Connected,
}

impl From<ChartEmptyCells> for xlsx::ChartEmptyCells {
    fn from(value: ChartEmptyCells) -> xlsx::ChartEmptyCells {
        match value {
            ChartEmptyCells::Gaps => xlsx::ChartEmptyCells::Gaps,
            ChartEmptyCells::Zero => xlsx::ChartEmptyCells::Zero,
            ChartEmptyCells::Connected => xlsx::ChartEmptyCells::Connected,
        }
    }
}
