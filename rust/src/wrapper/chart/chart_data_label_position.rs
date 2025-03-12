use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Debug, Clone, Copy)]
#[wasm_bindgen]
pub enum ChartDataLabelPosition {
    Default,
    Center,
    Right,
    Left,
    Above,
    Below,
    InsideBase,
    InsideEnd,
    OutsideEnd,
    BestFit,
}

impl From<ChartDataLabelPosition> for xlsx::ChartDataLabelPosition {
    fn from(position: ChartDataLabelPosition) -> Self {
        match position {
            ChartDataLabelPosition::Default => xlsx::ChartDataLabelPosition::Default,
            ChartDataLabelPosition::Center => xlsx::ChartDataLabelPosition::Center,
            ChartDataLabelPosition::Right => xlsx::ChartDataLabelPosition::Right,
            ChartDataLabelPosition::Left => xlsx::ChartDataLabelPosition::Left,
            ChartDataLabelPosition::Above => xlsx::ChartDataLabelPosition::Above,
            ChartDataLabelPosition::Below => xlsx::ChartDataLabelPosition::Below,
            ChartDataLabelPosition::InsideBase => xlsx::ChartDataLabelPosition::InsideBase,
            ChartDataLabelPosition::InsideEnd => xlsx::ChartDataLabelPosition::InsideEnd,
            ChartDataLabelPosition::OutsideEnd => xlsx::ChartDataLabelPosition::OutsideEnd,
            ChartDataLabelPosition::BestFit => xlsx::ChartDataLabelPosition::BestFit,
        }
    }
}
