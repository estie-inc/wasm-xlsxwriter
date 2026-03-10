use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartLayout {
    pub(crate) inner: Arc<Mutex<xlsx::ChartLayout>>,
}

#[wasm_bindgen]
impl ChartLayout {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartLayout {
        ChartLayout {
            inner: Arc::new(Mutex::new(xlsx::ChartLayout::new())),
        }
    }
    #[wasm_bindgen(js_name = "setOffset", skip_jsdoc)]
    pub fn set_offset(&self, x_offset: f64, y_offset: f64) -> ChartLayout {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_offset(x_offset, y_offset);
        *lock = inner;
        ChartLayout {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDimensions", skip_jsdoc)]
    pub fn set_dimensions(&self, width: f64, height: f64) -> ChartLayout {
        let mut lock = self.inner.lock().unwrap();
        let mut inner = std::mem::take(&mut *lock);
        inner = inner.set_dimensions(width, height);
        *lock = inner;
        ChartLayout {
            inner: Arc::clone(&self.inner),
        }
    }
}
