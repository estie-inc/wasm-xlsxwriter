use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLine {
    pub(crate) inner: Arc<Mutex<xlsx::ChartLine>>,
}

#[wasm_bindgen]
impl ChartLine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLine {
        ChartLine {
            inner: Arc::new(Mutex::new(xlsx::ChartLine::new())),
        }
    }
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: f64) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_width(width);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_transparency(transparency);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> ChartLine {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hidden(enable);
        ChartLine {
            inner: Arc::clone(&self.inner),
        }
    }
}
