use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Image` struct is used to create an object to represent an image that
/// can be inserted into a worksheet.
///
/// Output file:
#[derive(Clone)]
#[wasm_bindgen]
pub struct Image {
    pub(crate) inner: Arc<Mutex<xlsx::Image>>,
}

#[wasm_bindgen]
impl Image {
    #[wasm_bindgen(constructor)]
    pub fn new(path: P) -> Image {
        Image {
            inner: Arc::new(Mutex::new(xlsx::Image::new(path))),
        }
    }
    /// Get the width of the image used for the size calculations in Excel.
    ///
    /// Note, this gets the actual pixel width of the image and not the
    /// logical/scaled width set via {@link Image#setWidth}.
    #[wasm_bindgen(js_name = "width", skip_jsdoc)]
    pub fn width(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.width()
    }
    /// Get the height of the image used for the size calculations in Excel. See
    /// the example above.
    ///
    /// Note, this gets the actual pixel height of the image and not the
    /// logical/scaled height set via {@link Image#setHeight}.
    #[wasm_bindgen(js_name = "height", skip_jsdoc)]
    pub fn height(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.height()
    }
    /// Get the width/horizontal DPI of the image used for the size calculations
    /// in Excel. See the example above.
    ///
    /// Excel assumes a default image DPI of 96.0 and scales all other DPIs
    /// relative to that.
    #[wasm_bindgen(js_name = "widthDpi", skip_jsdoc)]
    pub fn width_dpi(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.width_dpi()
    }
    /// Get the height/vertical DPI of the image used for the size calculations
    /// in Excel. See the example above.
    ///
    /// Excel assumes a default image DPI of 96.0 and scales all other DPIs
    /// relative to that.
    #[wasm_bindgen(js_name = "heightDpi", skip_jsdoc)]
    pub fn height_dpi(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.height_dpi()
    }
}
