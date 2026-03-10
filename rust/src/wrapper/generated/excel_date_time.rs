use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ExcelDateTime {
    pub(crate) inner: Arc<Mutex<xlsx::ExcelDateTime>>,
}

#[wasm_bindgen]
impl ExcelDateTime {
    #[wasm_bindgen(js_name = "toExcel", skip_jsdoc)]
    pub fn to_excel(&self) -> f64 {
        let lock = self.inner.lock().unwrap();
        lock.to_excel()
    }
}
