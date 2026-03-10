use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct DataValidation {
    pub(crate) inner: Arc<Mutex<xlsx::DataValidation>>,
}

#[wasm_bindgen]
impl DataValidation {
    #[wasm_bindgen(constructor)]
    pub fn new() -> DataValidation {
        DataValidation {
            inner: Arc::new(Mutex::new(xlsx::DataValidation::new())),
        }
    }
}
