use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::{chart::chart_pattern_fill_type::ChartPatternFillType, color::Color};

#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartPatternFill {
    pub(crate) inner: xlsx::ChartPatternFill,
}

#[wasm_bindgen]
impl ChartPatternFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPatternFill {
        ChartPatternFill {
            inner: xlsx::ChartPatternFill::new(),
        }
    }

    #[wasm_bindgen(js_name = "setPattern")]
    pub fn set_pattern(&mut self, pattern: ChartPatternFillType) -> ChartPatternFill {
        self.inner.set_pattern(pattern.into());
        self.clone()
    }

    #[wasm_bindgen(js_name = "setBackgroundColor")]
    pub fn set_background_color(&mut self, color: Color) -> ChartPatternFill {
        self.inner.set_background_color(color);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setForegroundColor")]
    pub fn set_foreground_color(&mut self, color: Color) -> ChartPatternFill {
        self.inner.set_foreground_color(color);
        self.clone()
    }
}
