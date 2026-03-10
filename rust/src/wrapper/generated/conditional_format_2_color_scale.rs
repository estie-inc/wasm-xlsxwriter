use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ConditionalFormat2ColorScale {
    pub(crate) inner: Arc<Mutex<xlsx::ConditionalFormat2ColorScale>>,
}

#[wasm_bindgen]
impl ConditionalFormat2ColorScale {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ConditionalFormat2ColorScale {
        ConditionalFormat2ColorScale {
            inner: Arc::new(Mutex::new(xlsx::ConditionalFormat2ColorScale::new())),
        }
    }
}
