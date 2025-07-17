use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;

#[wasm_bindgen]
pub struct ChartGradientStop {
    pub(crate) inner: xlsx::ChartGradientStop,
}

#[wasm_bindgen]
impl ChartGradientStop {
    #[wasm_bindgen(constructor)]
    pub fn new(color: &Color, position: u8) -> ChartGradientStop {
        ChartGradientStop {
            inner: xlsx::ChartGradientStop::new(color.inner, position),
        }
    }
}
