use crate::wrapper::ChartErrorBarsDirection;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartErrorBars {
    pub(crate) inner: Arc<Mutex<xlsx::ChartErrorBars>>,
}

#[wasm_bindgen]
impl ChartErrorBars {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartErrorBars {
        ChartErrorBars {
            inner: Arc::new(Mutex::new(xlsx::ChartErrorBars::new())),
        }
    }
    #[wasm_bindgen(js_name = "setDirection", skip_jsdoc)]
    pub fn set_direction(&self, direction: ChartErrorBarsDirection) -> ChartErrorBars {
        let mut lock = self.inner.lock().unwrap();
        lock.set_direction(xlsx::ChartErrorBarsDirection::from(direction));
        ChartErrorBars {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setEndCap", skip_jsdoc)]
    pub fn set_end_cap(&self, enable: bool) -> ChartErrorBars {
        let mut lock = self.inner.lock().unwrap();
        lock.set_end_cap(enable);
        ChartErrorBars {
            inner: Arc::clone(&self.inner),
        }
    }
}
