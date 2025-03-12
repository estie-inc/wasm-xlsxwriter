use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[derive(Debug, Clone, Copy)]
#[wasm_bindgen]
pub enum ChartMarkerType {
    Square,
    Diamond,
    Triangle,
    X,
    Star,
    ShortDash,
    LongDash,
    Circle,
    PlusSign,
}

impl From<ChartMarkerType> for xlsx::ChartMarkerType {
    fn from(marker_type: ChartMarkerType) -> xlsx::ChartMarkerType {
        match marker_type {
            ChartMarkerType::Square => xlsx::ChartMarkerType::Square,
            ChartMarkerType::Diamond => xlsx::ChartMarkerType::Diamond,
            ChartMarkerType::Triangle => xlsx::ChartMarkerType::Triangle,
            ChartMarkerType::X => xlsx::ChartMarkerType::X,
            ChartMarkerType::Star => xlsx::ChartMarkerType::Star,
            ChartMarkerType::ShortDash => xlsx::ChartMarkerType::ShortDash,
            ChartMarkerType::LongDash => xlsx::ChartMarkerType::LongDash,
            ChartMarkerType::Circle => xlsx::ChartMarkerType::Circle,
            ChartMarkerType::PlusSign => xlsx::ChartMarkerType::PlusSign,
        }
    }
}
