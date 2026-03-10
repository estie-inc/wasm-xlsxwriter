use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

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
    #[wasm_bindgen(js_name = "setTransparency", skip_jsdoc)]
    pub fn set_transparency(&self, transparency: u8) -> ChartSolidFill {
        let mut lock = self.inner.lock().unwrap();
        lock.set_transparency(transparency);
        ChartSolidFill {
            inner: Arc::clone(&self.inner),
        }
    }
}
