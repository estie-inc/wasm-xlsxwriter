use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ShapePatternFill {
    pub(crate) inner: Arc<Mutex<xlsx::ShapePatternFill>>,
}

#[wasm_bindgen]
impl ShapePatternFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ShapePatternFill {
        ShapePatternFill {
            inner: Arc::new(Mutex::new(xlsx::ShapePatternFill::new())),
        }
    }
}
