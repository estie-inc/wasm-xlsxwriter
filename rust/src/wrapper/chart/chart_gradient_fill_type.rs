use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ChartGradientFillType {
    Linear, 
    Radial,
    Rectangular,
    Path,
}

impl From<ChartGradientFillType> for xlsx::ChartGradientFillType {
    fn from(value: ChartGradientFillType) -> Self {
        match value {
            ChartGradientFillType::Linear => xlsx::ChartGradientFillType::Linear,
            ChartGradientFillType::Radial => xlsx::ChartGradientFillType::Radial,
            ChartGradientFillType::Rectangular => xlsx::ChartGradientFillType::Rectangular,
            ChartGradientFillType::Path => xlsx::ChartGradientFillType::Path,
        }
    }
}
