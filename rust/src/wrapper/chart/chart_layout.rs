use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartLayout {
    pub(crate) inner: xlsx::ChartLayout,
}

#[wasm_bindgen]
impl ChartLayout {
    /// Create a new `ChartLayout` object.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLayout {
        ChartLayout {
            inner: xlsx::ChartLayout::new(),
        }
    }

    /// Set the offset of the chart layout.
    ///
    /// @param {number} x_offset - The x offset of the chart layout.
    /// @param {number} y_offset - The y offset of the chart layout.
    /// @return {ChartLayout} - The ChartLayout instance.
    #[wasm_bindgen(js_name = "setOffset")]
    pub fn set_offset(&mut self, x_offset: f64, y_offset: f64) -> ChartLayout {
        ChartLayout {
            inner: self.inner.clone().set_offset(x_offset, y_offset),
        }
    }

    /// Set the dimensions of the chart layout.
    ///
    /// @param {number} width - The width of the chart layout.
    /// @param {number} height - The height of the chart layout.
    /// @return {ChartLayout} - The ChartLayout instance.
    #[wasm_bindgen(js_name = "setDimensions")]
    pub fn set_dimensions(&mut self, width: f64, height: f64) -> ChartLayout {
        ChartLayout {
            inner: self.inner.clone().set_dimensions(width, height),
        }
    }
}
