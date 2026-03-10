use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `ChartSolidFill` struct represents a the solid fill for a chart element.
///
/// The {@link ChartSolidFill} struct represents the formatting properties for the
/// solid fill of a Chart element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ChartSolidFill` is a sub property of the {@link ChartFormat} struct and is used
/// with the {@link ChartFormat#setSolidFill} method.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartSolidFill {
    pub(crate) inner: Arc<Mutex<xlsx::ChartSolidFill>>,
}

#[wasm_bindgen]
impl ChartSolidFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartSolidFill {
        ChartSolidFill {
            inner: Arc::new(Mutex::new(xlsx::ChartSolidFill::new())),
        }
    }
    /// Set the transparency of a solid fill.
    ///
    /// Set the transparency of a solid fill color for a Chart element. You must
    /// also specify a fill color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ChartSolidFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_transparency(transparency);
        ChartSolidFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
