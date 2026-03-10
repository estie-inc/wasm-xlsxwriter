use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormatFormula {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormatFormula>>,
}

#[wasm_bindgen]
impl ConditionalFormatFormula {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormatFormula {
        ConditionalFormatFormula {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormatFormula::new())),
        }
    }
}
