use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
    #[wasm_bindgen(js_name = "width", skip_jsdoc)]
    pub fn width(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.width()
    }
    #[wasm_bindgen(js_name = "height", skip_jsdoc)]
    pub fn height(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.height()
    }
    #[wasm_bindgen(js_name = "widthDpi", skip_jsdoc)]
    pub fn width_dpi(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.width_dpi()
    }
    #[wasm_bindgen(js_name = "heightDpi", skip_jsdoc)]
    pub fn height_dpi(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.height_dpi()
    }
}
