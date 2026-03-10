use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartPoint {
    pub(crate) inner: Arc<Mutex<xlsx::ChartPoint>>,
}

#[wasm_bindgen]
impl ChartPoint {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartPoint {
        ChartPoint {
            inner: Arc::new(Mutex::new(xlsx::ChartPoint::new())),
        }
    }
}
