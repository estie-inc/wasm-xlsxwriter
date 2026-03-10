use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapeLine {
    pub(crate) inner: Arc<Mutex<xlsx::ShapeLine>>,
}

#[wasm_bindgen]
impl ShapeLine {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapeLine {
        ShapeLine {
            inner: Arc::new(Mutex::new(xlsx::ShapeLine::new())),
        }
    }
}
