use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Clone, Copy, PartialEq, Eq)]
#[wasm_bindgen]
pub enum ChartLegendPosition {
    Bottom,
    Left,
    Right,
    Top,
}

impl From<ChartLegendPosition> for xlsx::ChartLegendPosition {
    fn from(position: ChartLegendPosition) -> Self {
        match position {
            ChartLegendPosition::Bottom => xlsx::ChartLegendPosition::Bottom,
            ChartLegendPosition::Left => xlsx::ChartLegendPosition::Left,
            ChartLegendPosition::Right => xlsx::ChartLegendPosition::Right,
            ChartLegendPosition::Top => xlsx::ChartLegendPosition::Top,
        }
    }
}
