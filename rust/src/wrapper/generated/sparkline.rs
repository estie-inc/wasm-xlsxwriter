use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Sparkline {
    pub(crate) inner: Arc<Mutex<xlsx::Sparkline>>,
}

#[wasm_bindgen]
impl Sparkline {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Sparkline {
        Sparkline {
            inner: Arc::new(Mutex::new(xlsx::Sparkline::new())),
        }
    }
}
