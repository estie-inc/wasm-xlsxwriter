use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct FilterCondition {
    pub(crate) inner: Arc<Mutex<xlsx::FilterCondition>>,
}

#[wasm_bindgen]
impl FilterCondition {
    #[wasm_bindgen(constructor)]
    pub fn new() -> FilterCondition {
        FilterCondition {
            inner: Arc::new(Mutex::new(xlsx::FilterCondition::new())),
        }
    }
}
