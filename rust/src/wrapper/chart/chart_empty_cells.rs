use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

/// The `ChartEmptyCells` enum defines how empty cells are displayed in a chart.
///
/// This enum is used with the {@link Chart.showEmptyCellsAs} method to control
/// the display of empty or missing data in a chart series.
///
#[derive(Clone, Copy, PartialEq, Eq)]
#[wasm_bindgen]
pub enum ChartEmptyCells {
    /// Display empty cells as gaps in the chart. This is the default.
    Gaps,
    /// Display empty cells as zero values.
    Zero,
    /// Connect data points across empty cells with a line.
    Connected,
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

