use std::sync::Arc;

use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::{ChartSeries, Color};

#[wasm_bindgen]
impl ChartSeries {
    /// Set individual colors for each point/segment in the series.
    /// Accepts an array of Color values (e.g., `["Red", { RGB: 0xFF0000 }]`).
    #[wasm_bindgen(js_name = "setPointColors")]
    pub fn set_point_colors(&self, colors: JsValue) -> ChartSeries {
        let colors: Vec<Color> =
            serde_wasm_bindgen::from_value(colors).unwrap_or_default();
        let xlsx_colors: Vec<xlsx::Color> = colors.into_iter().map(xlsx::Color::from).collect();
        let mut lock = self.inner.lock().unwrap();
        lock.set_point_colors(&xlsx_colors);
        ChartSeries {
            inner: Arc::clone(&self.inner),
        }
    }
}
