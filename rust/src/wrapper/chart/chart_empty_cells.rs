use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// The `ChartEmptyCells` enum defines the Chart empty cell options.
///
/// This enum defines the Excel chart options for handling empty cells in the
/// chart data ranges.
#[wasm_bindgen]
#[derive(Clone, Copy)]
pub enum ChartEmptyCells {
    /// Show empty cells in the chart as gaps. The default.
    Gaps = 0,
    /// Show empty cells in the chart as zeroes.
    Zero = 1,
    /// Show empty cells in the chart connected by a line to the previous point.
    Connected = 2,
}

impl From<ChartEmptyCells> for xlsx::ChartEmptyCells {
    fn from(value: ChartEmptyCells) -> Self {
        match value {
            ChartEmptyCells::Gaps => xlsx::ChartEmptyCells::Gaps,
            ChartEmptyCells::Zero => xlsx::ChartEmptyCells::Zero,
            ChartEmptyCells::Connected => xlsx::ChartEmptyCells::Connected,
        }
    }
}
