use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPatternFill {
    pub(crate) inner: Arc<Mutex<xlsx::ChartPatternFill>>,
}

#[wasm_bindgen]
impl ChartPatternFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPatternFill {
        ChartPatternFill {
            inner: Arc::new(Mutex::new(xlsx::ChartPatternFill::new())),
        }
    }
}
